# transformer-design-tools
변압기 구조설계 업무 자동화 도구

## make_cube.py (Solid Edge 100 mm 정육면체)
- 설치: `pip install pywin32`
- 실행: `python make_cube.py` (Windows + Solid Edge 2024 이상)
- 오류 `Invalid class string` → Solid Edge 미설치/미등록. Solid Edge를 한 번 직접 실행 후 재시도
- 오류 `constants` 의 `AttributeError` → 상수 타입 라이브러리 미생성. `python -m win32com.client.makepy`에서 "Solid Edge Constants Type Library" 선택 후 재실행
