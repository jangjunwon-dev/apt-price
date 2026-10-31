# 🏢 아파트 매매 실거래가 조회기 (Public Data API + Streamlit)

국토교통부 공공데이터 API로 **지역·기간별 아파트 매매 실거래가**를 조회하고,
통계·그래프로 확인한 뒤 **정리된 엑셀 파일**로 내려받는 웹앱입니다.

## 주요 기능
- 시·도/시·군·구 선택 또는 지역 코드 직접 입력, 최대 24개월 기간 조회
- 해제(취소) 거래 제외, 거래금액·면적 숫자 정리, 평형·전용 평당가 계산
- 거래 건수, 평균·중위 거래가, 평균 전용 평당가 요약
- 월별 평균 거래가 추이 / 월별 거래 건수 그래프, 거래 많은 단지 TOP 10
- 아파트 이름 검색, 전용면적 필터, 열 정리된 엑셀 다운로드

## 사용 기술
Python, Requests, XML 파싱(ElementTree), pandas, openpyxl, Streamlit

## 데이터 출처
공공데이터포털 - 국토교통부_아파트 매매 실거래가 상세 자료 (Open API)

## 로컬 실행
```
pip install -r requirements.txt
streamlit run app.py
```
인증키는 코드에 넣지 않고, 화면에서 입력하거나 `.streamlit/secrets.toml`에
`SERVICE_KEY = "발급받은 인증키"` 형태로 저장해 사용합니다.
