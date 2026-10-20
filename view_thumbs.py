#!/usr/bin/env python3
"""NAS_IMG 에 저장된 썸네일을 DB 에서 읽어 HTML 한 장으로 만들고 브라우저로 연다.

    python view_thumbs.py                  # 최근 저장 50장
    python view_thumbs.py --limit 200 --bs 디지털새싹
"""
import argparse
import base64
import html
import os
import webbrowser

from image_loader import DEFAULT_CONFIG, connect, load_config

HERE = os.path.dirname(os.path.abspath(__file__))

PAGE = """<!doctype html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>NAS_IMG 썸네일</title>
<style>
:root {{ --bg:#f6f6f4; --card:#fff; --fg:#222; --muted:#777; --line:#e2e2de; --bad:#c0392b; }}
@media (prefers-color-scheme: dark) {{
  :root {{ --bg:#1b1b1d; --card:#26262a; --fg:#eee; --muted:#9a9a9f; --line:#38383d; --bad:#ff7b6b; }}
}}
body {{ margin:0; padding:16px; background:var(--bg); color:var(--fg);
       font:14px/1.45 "Malgun Gothic", system-ui, sans-serif; }}
h1 {{ font-size:18px; margin:0 0 4px; }}
.sum {{ color:var(--muted); margin-bottom:16px; }}
.grid {{ display:grid; gap:12px; grid-template-columns:repeat(auto-fill, minmax(260px, 1fr)); }}
.card {{ background:var(--card); border:1px solid var(--line); border-radius:8px; overflow:hidden; }}
.ph {{ height:240px; display:flex; align-items:center; justify-content:center; background:#0001; }}
.ph img {{ max-width:100%; max-height:240px; }}
.ph .err {{ color:var(--bad); padding:12px; font-size:12px; word-break:break-all; }}
.meta {{ padding:8px 10px; font-size:12px; }}
.meta b {{ font-size:13px; word-break:break-all; }}
.meta div {{ color:var(--muted); word-break:break-all; }}
.d {{ color:var(--bad); font-weight:bold; }}
</style></head><body>
<h1>NAS_IMG 썸네일</h1>
<div class="sum">{summary}</div>
<div class="grid">
{cards}
</div></body></html>
"""

CARD = """<div class="card">
  <div class="ph">{pic}</div>
  <div class="meta">
    <b>#{img_id} {file_name}</b>
    <div>{bs_name} · {bs_year} · bs_id {bs_id} · 상태 <span class="{state_cls}">{db_state}</span></div>
    <div>원본 {size_kb:,}KB · {img_w}×{img_h} · 수정일 {mod_date}</div>
    <div>썸네일 {thumb_kb} · {thumb_w}×{thumb_h} · {thumb_at}</div>
    <div>{folder_path}</div>
  </div>
</div>"""


def main():
    ap = argparse.ArgumentParser(description="NAS_IMG 썸네일을 HTML 로 보기")
    ap.add_argument("--config", default=DEFAULT_CONFIG)
    ap.add_argument("--table", default="NAS_IMG")
    ap.add_argument("--limit", type=int, default=50, help="보여줄 장 수 (기본 50)")
    ap.add_argument("--bs", help="사업명에 이 낱말이 든 것만")
    ap.add_argument("--out", default=os.path.join(HERE, "thumb_view.html"), help="만들 HTML 파일")
    ap.add_argument("--no-open", action="store_true", help="브라우저를 열지 않음")
    args = ap.parse_args()

    conn = connect(load_config(args.config))
    cur = conn.cursor()
    where, params = "WHERE thumb IS NOT NULL OR thumb_error IS NOT NULL", []
    if args.bs:
        where += " AND bs_name LIKE %s"
        params.append(f"%{args.bs}%")
    cur.execute(f"SELECT COUNT(*), SUM(thumb IS NOT NULL), SUM(LENGTH(thumb)), SUM(db_state='D') "
                f"FROM {args.table} {where}", params)
    total, with_thumb, thumb_bytes, deleted = cur.fetchone()
    cur.execute(
        f"SELECT img_id, bs_id, bs_name, bs_year, file_name, folder_path, size_kb, mod_date,"
        f" img_w, img_h, thumb, thumb_w, thumb_h, thumb_at, thumb_error, db_state"
        f" FROM {args.table} {where} ORDER BY img_id LIMIT %s", params + [args.limit])
    cols = [d[0] for d in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    conn.close()

    cards = []
    for r in rows:
        thumb = r.pop("thumb")
        if thumb:
            pic = f'<img src="data:image/jpeg;base64,{base64.b64encode(thumb).decode()}" alt="">'
            thumb_kb = f"{len(thumb) / 1024:.1f}KB"
        else:
            pic = f'<div class="err">썸네일 실패<br>{html.escape(r["thumb_error"] or "")}</div>'
            thumb_kb = "-"
        fields = {k: html.escape(str(v if v is not None else "-")) for k, v in r.items()}
        fields["size_kb"] = r["size_kb"] or 0
        cards.append(CARD.format(pic=pic, thumb_kb=thumb_kb,
                                 state_cls="d" if r["db_state"] == "D" else "", **{
                                     k: v for k, v in fields.items() if k != "size_kb"},
                                 size_kb=fields["size_kb"]))

    total = total or 0
    summary = (f"썸네일 있는 줄 {int(with_thumb or 0):,} / 처리한 줄 {total:,} · "
               f"썸네일 합계 {(thumb_bytes or 0) / 1024 / 1024:.2f}MB · 삭제(D) {int(deleted or 0):,} · "
               f"이 페이지 {len(rows):,}장")
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(PAGE.format(summary=summary, cards="\n".join(cards)))
    print(f"만들었습니다: {args.out}")
    if not args.no_open:
        webbrowser.open("file:///" + os.path.abspath(args.out).replace("\\", "/"))


if __name__ == "__main__":
    main()
