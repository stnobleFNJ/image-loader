# image-loader

`NAS이미지_사업별분류.xlsx`를 cafe24 MySQL(MariaDB) 데이터베이스에 저장합니다.

| 시트 | 테이블 | 키 |
|---|---|---|
| 사업별 요약 | `NAS_BS` | 사업명 (`bs_name` UNIQUE, 번호는 `bs_id` 자동 증가) |
| 이미지 목록 | `NAS_IMG` | 폴더 경로 + 파일명 (SHA1 `path_hash` UNIQUE, 번호는 `img_id` 자동 증가) |

`NAS_IMG.bs_id`는 사업명으로 찾은 `NAS_BS.bs_id`입니다. `bs_year`는 숫자이고 엑셀의 "연도미상"은 NULL,
`file_ext`는 파일명에서 뽑은 확장자(소문자)입니다.

## 준비

```
pip install -r requirements.txt
copy db_config.example.json db_config.json
```

`db_config.json`에 cafe24 DB 접속 정보를 적습니다. `password`를 비워 두면 환경변수
`CAFE24_DB_PASSWORD`를 쓰고, 그것도 없으면 실행할 때 묻습니다.

## 실행

```
python image_loader.py                        # 기본 엑셀 파일, 두 시트 모두
python image_loader.py 다른파일.xlsx
python image_loader.py --only img --limit 50  # 이미지 목록 앞 50줄만
python image_loader.py --only bs              # 사업별 요약만
python image_loader.py --table-suffix _TEST   # NAS_BS_TEST, NAS_IMG_TEST 에 시험 저장
```

## 저장 방식

- 테이블이 없으면 만들고, 있으면 그대로 씁니다.
- 같은 키의 줄이 이미 있으면 값을 비교해서 바뀐 것만 UPDATE 하고(`updated_at` 갱신), 같으면 건드리지 않습니다.
  없는 줄만 INSERT 합니다. 실행이 끝나면 테이블마다 `신규 / 변경 / 그대로` 건수를 보여 줍니다.
- 기존 테이블에 필요한 열이 없으면(예전 형식) 아무것도 쓰지 않고 멈춥니다.
- 엑셀에 없는 줄은 DB에서 지우지 않습니다.
