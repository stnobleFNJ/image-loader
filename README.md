# image-loader

`NAS이미지_사업별분류.xlsx`를 cafe24 MySQL(MariaDB) 데이터베이스에 저장하고,
NAS의 이미지를 읽어 썸네일을 `NAS_IMG`에 BLOB으로 넣습니다.

| 시트 | 테이블 | 키 |
|---|---|---|
| 사업별 요약 | `NAS_BS` | 사업명 + 실행 번호 (`bs_name`, `load_id` UNIQUE, 번호는 `bs_id` 자동 증가) |
| 이미지 목록 | `NAS_IMG` | 폴더 경로 + 파일명 + 실행 번호 (SHA1 `path_hash`, `load_id` UNIQUE, 번호는 `img_id` 자동 증가) |

- `NAS_IMG.bs_id`는 사업명으로 찾은 지금 쓰는(`Y`) `NAS_BS.bs_id`입니다. `bs_year`는 숫자이고 엑셀의 "연도미상"은 NULL,
  `file_ext`는 파일명에서 뽑은 확장자(소문자)입니다.
- 썸네일: 긴 변 320px JPEG(품질 80), 실측 평균 13~17KB. `thumb`(MEDIUMBLOB), `thumb_w`·`thumb_h`,
  원본 해상도 `img_w`·`img_h`. 읽지 못한 파일은 `thumb_error`에 이유가 남고 다시 시도하지 않습니다.
- `db_state`: 지금 쓰는 줄은 `Y`, 지난 실행에서 넣은 줄은 `D`. `load_id`는 실행마다 1씩 늘어나는 번호입니다.

## 준비

```
pip install -r requirements.txt
copy db_config.example.json db_config.json
```

`db_config.json`에 cafe24 DB 접속 정보를 적습니다. `password`를 비워 두면 환경변수
`CAFE24_DB_PASSWORD`를 쓰고, 그것도 없으면 실행할 때 묻습니다.
썸네일을 만들려면 이 PC에서 `\\192.168.0.6\...` NAS 경로가 열려야 합니다.

## 실행

```
python image_loader.py                        # 기본 엑셀 파일, 두 시트 모두 + 썸네일
python image_loader.py --limit 10             # 시트마다 앞 10줄만 (이때도 기존 줄은 모두 D)
python image_loader.py --only bs              # 사업별 요약만
python image_loader.py --no-thumb             # 썸네일 없이 목록만
python image_loader.py --thumb-size 480       # 썸네일 크기 바꾸기 (이미 만든 것은 그대로)
python image_loader.py --table-suffix _TEST   # NAS_BS_TEST, NAS_IMG_TEST 에 시험 저장

python view_thumbs.py                         # 저장된 썸네일을 HTML 로 만들어 브라우저로 보기
python view_thumbs.py --limit 200 --bs 디지털새싹
```

## 저장 방식

- 실행할 때마다 기존 줄을 모두 `db_state='D'`로 바꾸고, 엑셀 내용을 새 줄(`db_state='Y'`, 새 `load_id`)로 등록합니다.
  지난 줄은 지우지 않으므로 `load_id`별로 이력이 남습니다. `--only bs`/`--only img`면 그 테이블만 바뀝니다.
- 썸네일은 `NAS_IMG` 줄을 넣을 때 함께 넣습니다(100줄씩). 지난 실행에 같은 파일(경로·크기·수정일이 같음)의
  썸네일이 있으면 그것을 쓰고, 없거나 파일이 바뀐 것만 NAS 에서 새로 읽습니다.
  다 넣은 뒤 `D` 줄의 썸네일은 비웁니다(DB 용량이 실행마다 늘지 않게).
- 테이블이 없으면 만들고, 예전에 만든 테이블에 `db_state`·`load_id`·썸네일 열이 없으면 더하고
  중복 방지 기준을 (키, `load_id`)로 바꿉니다.
- 도중에 멈췄으면 처음부터 다시 실행하면 됩니다. 멈춘 실행에서 넣던 줄은 다음 실행 때 `D`가 됩니다.
