# Streamlit Community Cloud 배포 방법

## 1. GitHub 저장소 루트에 아래 파일이 있는지 확인

- `streamlit_app.py`
- `requirements.txt`
- `.streamlit/config.toml`
- `app/`
- `sample_data/`

## 2. Streamlit Community Cloud

Create app에서 다음을 선택합니다.

```text
Repository : 본인의 accounting-review-assistant 저장소
Branch     : main
Main file  : streamlit_app.py
```

그다음 Deploy를 누릅니다.

## 3. 주의

Streamlit Community Cloud의 로컬 SQLite는 영구 저장을 보장하지 않습니다.
이 버전은 공개 포트폴리오 데모를 목적으로 합니다.

## 4. GitHub 업로드 제외

다음 폴더는 올리지 마세요. `.gitignore`에도 포함되어 있습니다.

```text
__pycache__/
data/
uploads/
exports/
staging/
.venv/
```
