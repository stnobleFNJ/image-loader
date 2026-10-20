#!/usr/bin/env python3
"""NAS이미지_사업별분류.xlsx 를 cafe24 MySQL/MariaDB 데이터베이스에 저장한다.

    「사업별 요약」 시트 → NAS_BS   (사업명 하나당 한 줄)
    「이미지 목록」 시트 → NAS_IMG  (폴더 경로 + 파일명 하나당 한 줄, NAS 이미지의 썸네일 포함)

테이블이 이미 있으면 그대로 쓴다. 같은 키의 줄이 있으면 값을 견주어
바뀐 것만 UPDATE 하고, 같으면 건드리지 않으며, 없는 줄만 INSERT 한다.
실행할 때마다 기존 줄을 모두 db_state='D' 로 바꾸고 엑셀 내용을 새 줄(db_state='Y')로 등록한다.
한 번의 실행에서 넣은 줄은 같은 load_id 를 가진다. 같은 파일의 썸네일은 이전 줄에서 옮겨 쓴다.
"""
import argparse
import getpass
import hashlib
import io
import json
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from decimal import Decimal

DEFAULT_XLSX = r"D:\FNJ_WORK\10.회사관련자료\11.기타자료\2026_이미지메타검색\NAS이미지_사업별분류.xlsx"
DEFAULT_CONFIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "db_config.json")

BS_SHEET = "사업별 요약"
IMG_SHEET = "이미지 목록"
YEARS = ["2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026"]

THUMB_SIZE = 320        # 썸네일 긴 변(px). 실측 평균 13KB, 최대 25KB
THUMB_QUALITY = 80
THUMB_BATCH = 100       # NAS_IMG 를 한 번에 보내는 줄 수 (썸네일 포함 약 2MB)

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

# 새로 등록할 때 넣는 DB 열 (값 튜플 순서)
BS_DB_COLUMNS = [c for c, _ in BS_COLUMNS] + ["db_state"]
IMG_DB_COLUMNS = ["path_hash", "bs_id", "bs_name", "bs_year", "bs_type", "history_no", "folder_path",
                  "file_name", "file_ext", "size_kb", "mod_date", "copy_cnt", "db_state"]

# 예전에 만든 테이블에 없으면 ALTER TABLE 로 더하는 열
ADD_COLUMNS = {
    "NAS_BS": [
        ("db_state", "CHAR(1) NOT NULL DEFAULT 'Y'"),
        ("load_id", "INT NOT NULL DEFAULT 0"),
    ],
    "NAS_IMG": [
        ("db_state", "CHAR(1) NOT NULL DEFAULT 'Y'"),
        ("load_id", "INT NOT NULL DEFAULT 0"),
        ("img_w", "INT NULL"),
        ("img_h", "INT NULL"),
        ("thumb", "MEDIUMBLOB NULL"),
        ("thumb_w", "SMALLINT NULL"),
        ("thumb_h", "SMALLINT NULL"),
        ("thumb_at", "DATETIME NULL"),
        ("thumb_error", "VARCHAR(200) NULL"),
    ],
}
# (예전 UNIQUE 이름, 지금 UNIQUE 이름, 키 열) — 지금은 (키, load_id) 로 중복을 막는다.
UNIQUE_KEYS = {
    "NAS_BS": ("uk_bs_name", "uk_bs_name_load", "bs_name"),
    "NAS_IMG": ("uk_path_hash", "uk_path_hash_load", "path_hash"),
}


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
            db_state     CHAR(1) NOT NULL DEFAULT 'Y',
            load_id      INT NOT NULL DEFAULT 0,
            created_at   DATETIME NOT NULL,
            updated_at   DATETIME NOT NULL,
            UNIQUE KEY uk_bs_name_load (bs_name, load_id),
            KEY ix_load (load_id, db_state)
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
            img_w        INT NULL,
            img_h        INT NULL,
            thumb        MEDIUMBLOB NULL,
            thumb_w      SMALLINT NULL,
            thumb_h      SMALLINT NULL,
            thumb_at     DATETIME NULL,
            thumb_error  VARCHAR(200) NULL,
            path_hash    BINARY(20) NOT NULL,
            db_state     CHAR(1) NOT NULL DEFAULT 'Y',
            load_id      INT NOT NULL DEFAULT 0,
            created_at   DATETIME NOT NULL,
            updated_at   DATETIME NOT NULL,
            UNIQUE KEY uk_path_hash_load (path_hash, load_id),
            KEY ix_load (load_id, db_state),
            KEY ix_bs_id (bs_id),
            KEY ix_bs (bs_name, bs_year),
            KEY ix_file_name (file_name(100)),
            KEY ix_file_ext (file_ext),
            KEY ix_db_state (db_state)
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
    """엑셀 줄을 BS_DB_COLUMNS 순서로 바꾼다. 같은 사업명이 또 나오면 뒤의 것을 쓴다."""
    rows = {}
    for r in raw:
        r = list(r)
        r[3] = None if r[3] is None else str(r[3])          # 이력 No: "50, 54, 75"
        r[8] = None if r[8] is None else round(float(r[8]), 1)
        rows[r[0]] = tuple(r) + ("Y",)
    return list(rows.values())


def img_key(folder, name):
    return hashlib.sha1(f"{folder}\\{name}".encode("utf-8")).digest()


def img_rows(raw, bs_ids):
    """엑셀 줄을 IMG_DB_COLUMNS 순서로 바꾼다. bs_ids: {사업명: NAS_BS.bs_id}
    같은 폴더·파일명이 또 나오면 뒤의 것을 쓴다."""
    rows = {}
    for name, year, bs_type, hist, folder, fname, size_kb, mod, copies in raw:
        year = year if isinstance(year, int) else None      # "연도미상" → NULL
        mod = None if mod is None else str(mod)[:10]        # 수정일: YYYY-MM-DD
        ext = fname.rsplit(".", 1)[1].lower()[:10] if "." in fname else None
        key = img_key(folder, fname)
        rows[key] = (key, bs_ids.get(name), name, year, bs_type, hist, folder,
                     fname, ext, size_kb, mod, copies, "Y")
    return list(rows.values())


def ensure_columns(conn, table, base):
    """예전에 만든 테이블에 새 열(db_state·load_id·썸네일 등)이 없으면 더하고,
    키 하나짜리 UNIQUE 를 (키, load_id) UNIQUE 로 바꾼다."""
    cur = conn.cursor()
    cur.execute(f"SHOW COLUMNS FROM {table}")
    have = {r[0] for r in cur.fetchall()}
    for col, ddl in ADD_COLUMNS[base]:
        if col not in have:
            cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ddl}")
            print(f"  {table} 에 {col} 열을 더했습니다.")
    old_uk, new_uk, key_col = UNIQUE_KEYS[base]
    cur.execute(f"SHOW INDEX FROM {table}")
    names = {r[2] for r in cur.fetchall()}
    if new_uk not in names:
        drop = f"DROP INDEX {old_uk}, " if old_uk in names else ""
        cur.execute(f"ALTER TABLE {table} {drop}ADD UNIQUE KEY {new_uk} ({key_col}, load_id), "
                    f"ADD KEY ix_load (load_id, db_state)")
        print(f"  {table} 의 중복 방지 기준을 ({key_col}, load_id) 로 바꿨습니다.")
    conn.commit()


def next_load_id(conn, tables):
    """이번 실행 번호 = 두 테이블을 통틀어 가장 큰 load_id + 1"""
    cur = conn.cursor()
    cur.execute(" UNION ALL ".join(f"SELECT MAX(load_id) FROM {t}" for t in tables))
    return max((r[0] or 0) for r in cur.fetchall()) + 1


def reload(conn, table, cols, rows, load_id, batch, extend=None, extra_cols=()):
    """기존 줄을 모두 db_state='D' 로 바꾸고 rows 를 새 줄(load_id)로 넣는다.
    rows: [값 튜플] — cols 순서.
    extend: 한 묶음의 줄을 받아 extra_cols 값을 덧붙여 돌려주는 함수 (NAS_IMG 썸네일)."""
    cols = cols + list(extra_cols)
    cur = conn.cursor()
    cur.execute(f"SHOW COLUMNS FROM {table}")
    missing = sorted({*cols, "load_id"} - {r[0] for r in cur.fetchall()})
    if missing:
        sys.exit(f"기존 {table} 테이블에 필요한 열이 없습니다: {', '.join(missing)}\n"
                 "예전 형식의 테이블입니다. 이름을 바꾸거나 지운 뒤 다시 실행하세요.")
    gone = cur.execute(f"UPDATE {table} SET db_state='D', updated_at=NOW() WHERE db_state<>'D'")
    conn.commit()

    ins_cols = cols + ["load_id", "created_at", "updated_at"]
    ins_sql = (f"INSERT INTO {table} ({', '.join(ins_cols)}) VALUES "
               f"({', '.join(['%s'] * (len(cols) + 1))}, NOW(), NOW())")
    started = time.time()
    for i in range(0, len(rows), batch):
        part = [r + (load_id,) for r in rows[i:i + batch]]
        if extend:
            part = [r[:-1] + x + (load_id,) for r, x in zip(part, extend(part))]
        cur.executemany(ins_sql, part)
        conn.commit()
        n = min(i + batch, len(rows))
        rate = n / max(time.time() - started, 0.001)
        print(f"\r  {table} 등록 {n:,}/{len(rows):,} ({rate:,.1f}줄/초)", end="", flush=True)
    if rows:
        print()
    print(f"  {table}: 기존 {gone:,}줄 → D / 새로 등록 {len(rows):,}줄 (load_id {load_id})")


# ---------------------------------------------------------------- 썸네일

def make_thumb(path, size):
    """NAS 이미지를 읽어 (JPEG 바이트, 썸네일 너비, 높이, 원본 너비, 높이) 를 돌려준다."""
    from PIL import Image, ImageOps
    with Image.open(path) as im:
        w, h = im.size
        if im.getexif().get(0x0112) in (5, 6, 7, 8):     # 세로로 돌려 찍은 사진
            w, h = h, w
        im.draft("RGB", (size * 2, size * 2))           # JPEG 는 작게 풀어 빠르게 읽음
        im = ImageOps.exif_transpose(im)
        im.thumbnail((size, size))
        if im.mode in ("RGBA", "LA", "P"):
            rgba = im.convert("RGBA")
            bg = Image.new("RGB", rgba.size, (255, 255, 255))
            bg.paste(rgba, mask=rgba.split()[3])
            im = bg
        elif im.mode != "RGB":
            im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=THUMB_QUALITY, optimize=True)
    return buf.getvalue(), im.width, im.height, w, h


THUMB_COLS = ["img_w", "img_h", "thumb", "thumb_w", "thumb_h", "thumb_at", "thumb_error"]
NO_THUMB = (None,) * len(THUMB_COLS)


def setup_pillow():
    try:
        from PIL import Image, ImageFile
        import pillow_heif
        pillow_heif.register_heif_opener()               # 아이폰 HEIC
    except ImportError:
        sys.exit("Pillow / pillow-heif 가 없습니다.  pip install -r requirements.txt  로 설치하세요.")
    Image.MAX_IMAGE_PIXELS = 400_000_000
    ImageFile.LOAD_TRUNCATED_IMAGES = True               # 끝이 몇십 바이트 잘린 JPEG(카카오톡 등)도 읽음


def retry_thumbs(conn, table, size, workers):
    """db_state='Y' 인데 썸네일이 없는 줄만 NAS 에서 다시 읽어 채운다 (load_id 그대로).
    그래도 만들지 못한 줄은 D 로 바꾼다."""
    setup_pillow()
    cur = conn.cursor()
    cur.execute(f"SELECT img_id, folder_path, file_name FROM {table} WHERE db_state='Y' AND thumb IS NULL")
    todo = cur.fetchall()
    print(f"  썸네일 없는 Y 줄: {len(todo):,}")

    def job(row):
        try:
            return row[0], make_thumb(os.path.join(row[1], row[2]), size), None
        except Exception as e:
            return row[0], None, f"{type(e).__name__}: {e}"[:200]

    ok_sql = (f"UPDATE {table} SET thumb=%s, thumb_w=%s, thumb_h=%s, img_w=%s, img_h=%s,"
              f" thumb_at=NOW(), thumb_error=NULL WHERE img_id=%s")
    err_sql = f"UPDATE {table} SET thumb_error=%s, thumb_at=NOW(), db_state='D' WHERE img_id=%s"
    made = failed = 0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for i in range(0, len(todo), THUMB_BATCH):
            oks, errs = [], []
            for img_id, res, err in pool.map(job, todo[i:i + THUMB_BATCH]):
                if res:
                    oks.append(res + (img_id,))
                else:
                    errs.append((err, img_id))
            if oks:
                cur.executemany(ok_sql, oks)
            if errs:
                cur.executemany(err_sql, errs)
            conn.commit()
            made += len(oks)
            failed += len(errs)
            print(f"\r  썸네일 다시 만들기 {made + failed:,}/{len(todo):,}", end="", flush=True)
    if todo:
        print()
    print(f"  썸네일: 새로 만듦 {made:,} / 여전히 실패 {failed:,}" + (" (D 로 바꿈)" if failed else ""))


class Thumbnailer:
    """NAS_IMG 에 넣을 한 묶음의 줄마다 THUMB_COLS 값을 만든다.
    지난 실행(D 줄)에 같은 파일(경로·크기·수정일)의 썸네일이 있으면 그것을 쓰고,
    없으면 NAS 에서 이미지를 읽어 만든다."""

    def __init__(self, conn, table, load_id, size, workers, make_new=True):
        if make_new:
            setup_pillow()
        self.cur = conn.cursor()
        self.table, self.load_id, self.size, self.make_new = table, load_id, size, make_new
        self.pool = ThreadPoolExecutor(max_workers=workers)
        self.reused = self.made = self.failed = self.bytes = 0

    def _job(self, row):
        path = os.path.join(row[6], row[7])              # folder_path, file_name
        try:
            thumb, tw, th, w, h = make_thumb(path, self.size)
            return (w, h, thumb, tw, th, datetime.now().replace(microsecond=0), None)
        except Exception as e:                           # 파일 없음·깨진 파일·지원 안 하는 형식
            return (None, None, None, None, None, datetime.now().replace(microsecond=0),
                    f"{type(e).__name__}: {e}"[:200])

    def __call__(self, part):
        # part 의 줄: IMG_DB_COLUMNS 순서 — 0 path_hash, 9 size_kb, 10 mod_date
        keys = [r[0] for r in part]
        self.cur.execute(
            f"SELECT path_hash, size_kb, mod_date, {', '.join(THUMB_COLS)} FROM {self.table}"
            f" WHERE path_hash IN ({', '.join(['%s'] * len(keys))})"
            f" AND load_id < %s AND thumb IS NOT NULL", keys + [self.load_id])
        prev = {(r[0], r[1], norm(r[2])): tuple(r[3:]) for r in self.cur.fetchall()}

        out = [prev.get((r[0], r[9], r[10])) for r in part]
        self.reused += sum(1 for x in out if x)
        todo = [i for i, x in enumerate(out) if x is None]
        if self.make_new:
            for i, res in zip(todo, self.pool.map(self._job, [part[i] for i in todo])):
                out[i] = res
                if res[2]:
                    self.made += 1
                    self.bytes += len(res[2])
                else:
                    self.failed += 1
        return [x or NO_THUMB for x in out]

    def finish(self, conn):
        self.pool.shutdown()
        # 새 줄로 옮겨 간 썸네일은 D 줄에서 비워 DB 용량이 실행마다 늘지 않게 한다.
        self.cur.execute(f"UPDATE {self.table} SET thumb=NULL WHERE load_id < %s AND thumb IS NOT NULL",
                         (self.load_id,))
        # 썸네일을 만들지 못한 줄(파일 없음·깨진 파일)은 쓰지 않는 줄로 본다.
        if self.make_new:
            self.cur.execute(f"UPDATE {self.table} SET db_state='D' WHERE load_id=%s AND thumb IS NULL"
                             f" AND thumb_error IS NOT NULL", (self.load_id,))
        conn.commit()
        avg = self.bytes / self.made / 1024 if self.made else 0
        print(f"  썸네일: 새로 만듦 {self.made:,} (평균 {avg:.1f}KB) / 지난 것 씀 {self.reused:,}"
              f" / 실패 {self.failed:,}"
              + ("  — 실패한 줄은 D 로 두었고, 이유는 thumb_error 열에 있습니다." if self.failed else ""))


def main():
    ap = argparse.ArgumentParser(description="NAS이미지_사업별분류.xlsx 를 cafe24 DB 에 저장")
    ap.add_argument("xlsx", nargs="?", default=DEFAULT_XLSX, help="엑셀 파일")
    ap.add_argument("--config", default=DEFAULT_CONFIG, help="DB 설정 파일 (기본: db_config.json)")
    ap.add_argument("--batch", type=int, default=1000, help="한 번에 보낼 줄 수 (기본 1000)")
    ap.add_argument("--limit", type=int, default=0,
                    help="시트마다 앞에서부터 이 줄 수만 저장 (기본 0 = 전체)")
    ap.add_argument("--only", choices=["bs", "img"], help="한 테이블만 저장")
    ap.add_argument("--no-thumb", action="store_true", help="썸네일을 만들지 않음")
    ap.add_argument("--thumb-size", type=int, default=THUMB_SIZE, help=f"썸네일 긴 변 px (기본 {THUMB_SIZE})")
    ap.add_argument("--workers", type=int, default=8, help="썸네일을 동시에 읽는 개수 (기본 8)")
    ap.add_argument("--retry-thumbs", action="store_true",
                    help="엑셀은 읽지 않고, 썸네일이 없는 Y 줄만 다시 만듦 (load_id·상태 그대로)")
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

    if args.retry_thumbs:
        conn = connect(load_config(args.config))
        try:
            retry_thumbs(conn, img_table, args.thumb_size, args.workers)
        finally:
            conn.close()
        print(f"걸린 시간 : {time.time() - started:,.0f}초")
        return

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
        ensure_columns(conn, bs_table, "NAS_BS")
        ensure_columns(conn, img_table, "NAS_IMG")
        load_id = next_load_id(conn, [bs_table, img_table])
        print(f"  이번 실행 번호(load_id): {load_id}")
        if bs is not None:
            reload(conn, bs_table, BS_DB_COLUMNS, bs, load_id, args.batch)
        if img_raw is not None:
            # 지금 쓰는(Y) NAS_BS 의 번호(bs_id)를 이미지에 붙인다.
            cur.execute(f"SELECT bs_name, bs_id FROM {bs_table} WHERE db_state='Y'")
            bs_ids = dict(cur.fetchall())
            img = img_rows(img_raw, bs_ids)
            no_bs = sum(1 for v in img if v[1] is None)
            if no_bs:
                print(f"  {bs_table}에 없는 사업명 {no_bs:,}줄은 bs_id 를 비워 둡니다.")
            # 줄을 넣을 때 썸네일도 함께 넣는다. 한 줄에 썸네일 약 20KB 가 붙으므로
            # DB 한 번 전송 한도(16MB)를 넘지 않게 THUMB_BATCH 줄씩 보낸다.
            thumbs = Thumbnailer(conn, img_table, load_id, args.thumb_size, args.workers,
                                 make_new=not args.no_thumb)
            reload(conn, img_table, IMG_DB_COLUMNS, img, load_id, min(args.batch, THUMB_BATCH),
                   extend=thumbs, extra_cols=THUMB_COLS)
            thumbs.finish(conn)
    except KeyboardInterrupt:
        conn.rollback()
        print("\n중단했습니다. 처음부터 다시 실행하세요 (이번에 넣던 줄은 다음 실행 때 D 가 됩니다).")
        sys.exit(1)
    finally:
        conn.close()
    print(f"걸린 시간 : {time.time() - started:,.0f}초")


if __name__ == "__main__":
    main()
