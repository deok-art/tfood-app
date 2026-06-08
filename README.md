# 식품이력등록 데이터 취합 앱

Streamlit 기반 엑셀 취합용 웹 애플리케이션입니다.

## 설치

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 실행

```powershell
python -m streamlit run app.py
```

## 현재 범위

- 엑셀 클립보드 TSV 붙여넣기
- 여러 엑셀 파일 업로드
- 파일명/첫 행 기반 거래처 자동 인식 placeholder
- 인식 실패 시 거래처 수동 선택
- pass-through 가공 후 엑셀 다운로드
- 서버 저장 없음. 입력과 산출물은 Streamlit 세션 메모리에서만 처리합니다.

거래처별 실제 가공 규칙은 `processors.py`의 `VENDOR_ALIASES`와 `PROCESSOR_REGISTRY`에 추가합니다.
