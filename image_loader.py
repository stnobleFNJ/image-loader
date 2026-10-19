#!/usr/bin/env python3
"""NAS이미지_사업별분류.xlsx 를 cafe24 MySQL/MariaDB 데이터베이스에 저장한다.

    「사업별 요약」 시트 → NAS_BS   (사업명 하나당 한 줄)
    「이미지 목록」 시트 → NAS_IMG  (폴더 경로 + 파일명 하나당 한 줄)

테이블이 이미 있으면 그대로 쓴다. 같은 키의 줄이 있으면 값을 견주어
바뀐 것만 UPDATE 하고, 같으면 건드리지 않으며, 없는 줄만 INSERT 한다.
"""
import argparse
import getpass
import hashlib
import json
import os
import sys
import time
from datetime import date, datetime
from decimal import Decimal

DEFAULT_XLSX = r"D:\FNJ_WORK\10.회사관련자료\11.기타자료\2026_이미지메타검색\NAS이미지_사업별분류.xlsx"
DEFAULT_CONFIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "db_config.json")

BS_SHEET = "사업별 요약"
IMG_SHEET = "이미지 목록"
YEARS = ["2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026"]

# (DB 열, 엑셀 머리글) — 머리글 이름으로 찾으므로 시트의 열 순서가 바뀌어도 된다.
BS_COLUMNS = [
    ("bs_name", "사업명"),
    ("bs_type", "구분"),
    ("bs_years", "사업 연도"),
    ("history_no", "이력 No"),
    ("client", "발주기관"),
    ("folder_cnt", "폴더 수"),
    ("image_cnt", "이미지 수(중복 제거)"),
    ("dup_cnt", "중복 사본 수"),
    ("size_mb", "용량(MB)"),
] + [(f"img_{y}", y) for y in YEARS] + [("img_unknown", "연도미상")]

IMG_COLUMNS = [
    ("bs_name", "사업명"),
    ("bs_year", "년도"),
    ("bs_type", "구분"),
    ("history_no", "이력 No"),
    ("folder_path", "폴더 경로"),
    ("file_name", "파일명"),
    ("size_kb", "크기(KB)"),
    ("mod_date", "수정일"),
    ("copy_cnt", "같은 사진 수(사본 포함)"),
]
# NAS_IMG 에 쓰는 열: 엑셀 열 앞에 bs_id(NAS_BS 번호), 파일명 뒤에 file_ext 를 더한다.
IMG_DB_COLUMNS = ["bs_id", "bs_name", "bs_year", "bs_type", "history_no", "folder_path",
                  "file_name", "file_ext", "size_kb", "mod_date", "copy_cnt"]


def schema(bs_table, img_table):
    year_cols = "\n".join(f"        img_{y} INT NULL," for y in YEARS)
    return [
        f"""
        CREATE TABLE IF NOT EXISTS {bs_table} (
            bs_id        INT AUTO_INCREMENT PRIMARY KEY,
            bs_name      VARCHAR(100) NOT NULL,
            bs_type      VARCHAR(20)  NULL,
            bs_years     VARCHAR(100) NULL,
            history_no   VARCHAR(100) NULL,
            client       VARCHAR(200) NULL,
            folder_cnt   INT NULL,
            image_cnt    INT NULL,
            dup_cnt      INT NULL,
            size_mb      DECIMAL(12,1) NULL,
{year_cols}
            img_unknown  INT NULL,
            created_at   DATETIME NOT NULL,
            updated_at   DATETIME NOT NULL,
            UNIQUE KEY uk_bs_name (bs_name)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """,
        f"""
        CREATE TABLE IF NOT EXISTS {img_table} (
            img_id       BIGINT AUTO_INCREMENT PRIMARY KEY,
            bs_id        INT NULL,
            bs_name      VARCHAR(100) NOT NULL,
            bs_year      SMALLINT NULL,
            bs_type      VARCHAR(20)  NULL,
            history_no   INT NULL,
            folder_path  VARCHAR(500) NOT NULL,
            file_name    VARCHAR(255) NOT NULL,
            file_ext     VARCHAR(10)  NULL,
            size_kb      INT NULL,
            mod_date     DATE NULL,
            copy_cnt     INT NULL,
            path_hash    BINARY(20) NOT NULL,
            created_at   DATETIME NOT NULL,
            updated_at   DATETIME NOT NULL,
            UNIQUE KEY uk_path_hash (path_hash),
            KEY ix_bs_id (bs_id),
            KEY ix_bs (bs_name, bs_year),
            KEY ix_file_name (file_name(100)),
            KEY ix_file_ext (file_ext)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
        """,
    ]


def load_config(path):
    if not os.path.exists(path):
        sys.exit(f"DB 설정 파일이 없습니다: {path}\n"
                 "db_config.example.json 을 db_config.json 으로 복사해 값을 채우세요.")
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg.setdefault("port", 3306)
    if not cfg.get("password"):
        cfg["password"] = os.environ.get("CAFE24_DB_PASSWORD") or getpass.getpass("DB 비밀번호: ")
    return cfg


def connect(cfg):
    try:
        import pymysql
    except ImportError:
        sys.exit("pymysql 이 없습니다.  pip install -r requirements.txt  로 설치하세요.")
    return pymysql.connect(
        host=cfg["host"], port=int(cfg["port"]), user=cfg["user"],
        password=cfg["password"], database=cfg["database"],
        charset="utf8mb4", autocommit=False, connect_timeout=15,
    )


def norm(v):
    """엑셀 값과 DB 값을 같은 모양으로 맞춰 견줄 수 있게 한다."""
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        return v or None
    if isinstance(v, Decimal):
        return float(v)
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return v


def read_sheet(wb, sheet, columns, limit):
    ws = wb[sheet]
    rows = ws.iter_rows(values_only=True)
    header = [norm(h) for h in next(rows)]
    try:
        idx = [header.index(title) for _, title in columns]
    except ValueError as e:
        sys.exit(f"「{sheet}」 시트에 필요한 열이 없습니다: {e}")
    out = []
    for r in rows:
        if not any(c is not None for c in r):
            continue
        out.append(tuple(norm(r[i]) if i < len(r) else None for i in idx))
        if limit and len(out) >= limit:
            break
    return out


def bs_rows(raw):
    rows = {}
    for r in raw:
        r = list(r)
        r[3] = None if r[3] is None else str(r[3])          # 이력 No: "50, 54, 75"
        r[8] = None if r[8] is None else round(float(r[8]), 1)
        rows[r[0]] = tuple(r)
    return rows


def img_key(folder, name):
    return hashlib.sha1(f"{folder}\\{name}".encode("utf-8")).digest()


def img_rows(raw, bs_ids):
    """엑셀 줄을 IMG_DB_COLUMNS 순서로 바꾼다. bs_ids: {사업명: NAS_BS.bs_id}"""
    rows = {}
    for name, year, bs_type, hist, folder, fname, size_kb, mod, copies in raw:
        year = year if isinstance(year, int) else None      # "연도미상" → NULL
        mod = None if mod is None else str(mod)[:10]        # 수정일: YYYY-MM-DD
        ext = fname.rsplit(".", 1)[1].lower()[:10] if "." in fname else None
        rows[img_key(folder, fname)] = (bs_ids.get(name), name, year, bs_type, hist, folder,
                                        fname, ext, size_kb, mod, copies)
    return rows


def sync(conn, table, key_col, cols, rows, fetch_key, batch, required=()):
    """rows: {키: 값 튜플}. 기존 값과 견주어 INSERT / UPDATE / 그대로 를 나눈다."""
    cur = conn.cursor()
    cur.execute(f"SHOW COLUMNS FROM {table}")
    missing = sorted({key_col, *cols, *required} - {r[0] for r in cur.fetchall()})
    if missing:
        sys.exit(f"기존 {table} 테이블에 필요한 열이 없습니다: {', '.join(missing)}\n"
                 "예전 형식의 테이블입니다. 이름을 바꾸거나 지운 뒤 다시 실행하세요.")
    cur.execute(f"SELECT {fetch_key}, {', '.join(cols)} FROM {table}")
    existing = {r[0]: tuple(norm(v) for v in r[1:]) for r in cur.fetchall()}

    inserts, updates, same = [], [], 0
    for key, vals in rows.items():
        old = existing.get(key)
        if old is None:
            inserts.append((key, vals))
        elif old != vals:
            updates.append((key, vals))
        else:
            same += 1

    extra_key = key_col not in cols                      # NAS_IMG: path_hash 는 값 목록에 없음
    ins_cols = ([key_col] if extra_key else []) + cols + ["created_at", "updated_at"]
    ins_sql = (f"INSERT INTO {table} ({', '.join(ins_cols)}) VALUES "
               f"({', '.join(['%s'] * (len(ins_cols) - 2))}, NOW(), NOW())")
    upd_sql = (f"UPDATE {table} SET {', '.join(f'{c}=%s' for c in cols)}, updated_at=NOW() "
               f"WHERE {key_col}=%s")

    for i in range(0, len(inserts), batch):
        part = inserts[i:i + batch]
        cur.executemany(ins_sql, [((k,) if extra_key else ()) + v for k, v in part])
        conn.commit()
        print(f"\r  {table} 신규 {min(i + batch, len(inserts)):,}/{len(inserts):,}", end="", flush=True)
    if inserts:
        print()
    for i in range(0, len(updates), batch):
        part = updates[i:i + batch]
        cur.executemany(upd_sql, [v + (k,) for k, v in part])
        conn.commit()
        print(f"\r  {table} 변경 {min(i + batch, len(updates)):,}/{len(updates):,}", end="", flush=True)
    if updates:
        print()
    print(f"  {table}: 신규 {len(inserts):,} / 변경 {len(updates):,} / 그대로 {same:,}")


def main():
    ap = argparse.ArgumentParser(description="NAS이미지_사업별분류.xlsx 를 cafe24 DB 에 저장")
    ap.add_argument("xlsx", nargs="?", default=DEFAULT_XLSX, help="엑셀 파일")
    ap.add_argument("--config", default=DEFAULT_CONFIG, help="DB 설정 파일 (기본: db_config.json)")
    ap.add_argument("--batch", type=int, default=1000, help="한 번에 보낼 줄 수 (기본 1000)")
    ap.add_argument("--limit", type=int, default=0, help="시트마다 앞에서부터 이 줄 수만 저장 (기본 0 = 전체)")
    ap.add_argument("--only", choices=["bs", "img"], help="한 테이블만 저장")
    ap.add_argument("--table-suffix", default="", help="테이블 이름 뒤에 붙일 말 (예: _TEST → NAS_BS_TEST)")
    args = ap.parse_args()

    if not os.path.exists(args.xlsx):
        sys.exit(f"엑셀 파일이 없습니다: {args.xlsx}")
    try:
        import openpyxl
    except ImportError:
        sys.exit("openpyxl 이 없습니다.  pip install -r requirements.txt  로 설치하세요.")

    bs_table = "NAS_BS" + args.table_suffix
    img_table = "NAS_IMG" + args.table_suffix
    started = time.time()

    print(f"엑셀 파일 : {args.xlsx}")
    wb = openpyxl.load_workbook(args.xlsx, read_only=True, data_only=True)
    bs = img_raw = None
    if args.only in (None, "bs"):
        bs = bs_rows(read_sheet(wb, BS_SHEET, BS_COLUMNS, args.limit))
        print(f"  「{BS_SHEET}」 {len(bs):,}줄")
    if args.only in (None, "img"):
        img_raw = read_sheet(wb, IMG_SHEET, IMG_COLUMNS, args.limit)
        print(f"  「{IMG_SHEET}」 {len(img_raw):,}줄")
    wb.close()

    conn = connect(load_config(args.config))
    try:
        cur = conn.cursor()
        for ddl in schema(bs_table, img_table):
            cur.execute(ddl)
        conn.commit()
        if bs is not None:
            sync(conn, bs_table, "bs_name", [c for c, _ in BS_COLUMNS],
                 bs, "bs_name", args.batch, required=("bs_id",))
        if img_raw is not None:
            # NAS_BS 를 먼저 반영한 뒤 그 번호(bs_id)를 이미지에 붙인다.
            cur.execute(f"SELECT bs_name, bs_id FROM {bs_table}")
            bs_ids = dict(cur.fetchall())
            img = img_rows(img_raw, bs_ids)
            no_bs = sum(1 for v in img.values() if v[0] is None)
            if no_bs:
                print(f"  {bs_table}에 없는 사업명 {no_bs:,}줄은 bs_id 를 비워 둡니다. "
                      f"{bs_table}을 채운 뒤 다시 실행하면 채워집니다.")
            sync(conn, img_table, "path_hash", IMG_DB_COLUMNS,
                 img, "path_hash", args.batch, required=("img_id",))
    except KeyboardInterrupt:
        conn.rollback()
        print("\n중단했습니다. 다시 실행하면 이미 들어간 줄은 그대로 두고 이어서 넣습니다.")
        sys.exit(1)
    finally:
        conn.close()
    print(f"걸린 시간 : {time.time() - started:,.0f}초")


if __name__ == "__main__":
    main()
