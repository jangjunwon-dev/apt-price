"""
아파트 매매 실거래가 조회기 - Streamlit 웹앱
데이터: 공공데이터포털 '국토교통부_아파트 매매 실거래가 상세 자료' API

설치 (노트북 셀에서):
    %pip install streamlit requests pandas openpyxl
실행 (터미널에서, app.py가 있는 폴더에서):
    & "C:\\Users\\junwo\\AppData\\Local\\Programs\\Python\\Python315\\python.exe" -m streamlit run app.py

인증키: 공공데이터포털 마이페이지의 '일반 인증키(Decoding)'를 사용
    - 로컬: .streamlit/secrets.toml 에 SERVICE_KEY = "키" 저장 (GitHub에 올리지 말 것!)
    - Streamlit 클라우드: 앱 설정 > Secrets 에 같은 내용 입력
"""

import io
import xml.etree.ElementTree as ET
from urllib.parse import unquote
from datetime import date

import pandas as pd
import requests
import streamlit as st

API_URL = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev"

# 법정동 코드 앞 5자리 (시군구 코드). 목록에 없는 지역은 '직접 입력'으로 조회
REGIONS = {
    "서울특별시": {
        "종로구": "11110", "중구": "11140", "용산구": "11170", "성동구": "11200",
        "광진구": "11215", "동대문구": "11230", "중랑구": "11260", "성북구": "11290",
        "강북구": "11305", "도봉구": "11320", "노원구": "11350", "은평구": "11380",
        "서대문구": "11410", "마포구": "11440", "양천구": "11470", "강서구": "11500",
        "구로구": "11530", "금천구": "11545", "영등포구": "11560", "동작구": "11590",
        "관악구": "11620", "서초구": "11650", "강남구": "11680", "송파구": "11710",
        "강동구": "11740",
    },
    "부산광역시": {"해운대구": "26350", "수영구": "26500"},
    "대구광역시": {"수성구": "27260"},
    "인천광역시": {"연수구": "28185"},
    "대전광역시": {"유성구": "30200"},
    "세종특별자치시": {"세종시": "36110"},
    "경기도": {"성남시 분당구": "41135", "수원시 영통구": "41117", "용인시 수지구": "41465"},
    "전북특별자치도": {"전주시 완산구": "52111", "전주시 덕진구": "52113"},
}

COLUMN_MAP = {  # API 응답 태그 → 한글 열 이름
    "aptNm": "아파트", "umdNm": "법정동", "jibun": "지번", "aptDong": "동",
    "floor": "층", "excluUseAr": "전용면적(㎡)", "dealAmount": "거래금액(만원)",
    "buildYear": "건축년도", "dealingGbn": "거래유형", "cdealType": "해제여부",
}


# ---------------- API 호출 ----------------
def month_list(start: date, end: date):
    """시작~끝 사이의 'YYYYMM' 문자열 목록"""
    months, y, m = [], start.year, start.month
    while (y, m) <= (end.year, end.month):
        months.append(f"{y}{m:02d}")
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return months


def parse_items(xml_text):
    """XML 응답에서 거래 목록과 전체 건수 추출"""
    root = ET.fromstring(xml_text)
    code = root.findtext(".//resultCode", default="")
    if code not in ("000", "00"):
        msg = root.findtext(".//resultMsg", default="알 수 없는 오류")
        raise RuntimeError(f"API 오류 ({code}): {msg}")
    total = int(root.findtext(".//totalCount", default="0") or 0)
    rows = [{child.tag: (child.text or "").strip() for child in item}
            for item in root.iter("item")]
    return rows, total


@st.cache_data(ttl=60 * 60 * 6, show_spinner=False)
def fetch_month(service_key, lawd_cd, deal_ymd):
    """한 달치 거래를 페이지를 넘기며 모두 가져옴 (6시간 캐시)"""
    rows, page = [], 1
    while True:
        params = {"serviceKey": service_key, "LAWD_CD": lawd_cd,
                  "DEAL_YMD": deal_ymd, "pageNo": page, "numOfRows": 1000}
        res = requests.get(API_URL, params=params, timeout=20)
        res.raise_for_status()
        items, total = parse_items(res.text)
        rows.extend(items)
        if len(rows) >= total or not items:
            break
        page += 1
    return rows


def clean(rows):
    """원본 응답을 분석하기 좋은 표로 정리"""
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["계약일"] = pd.to_datetime(
        df["dealYear"] + "-" + df["dealMonth"].str.zfill(2) + "-" + df["dealDay"].str.zfill(2),
        errors="coerce")
    df = df.rename(columns=COLUMN_MAP)
    df["거래금액(만원)"] = pd.to_numeric(df["거래금액(만원)"].str.replace(",", ""), errors="coerce")
    df["전용면적(㎡)"] = pd.to_numeric(df["전용면적(㎡)"], errors="coerce")
    df["평형(평)"] = (df["전용면적(㎡)"] / 3.3058).round(1)
    df["전용평당가(만원)"] = (df["거래금액(만원)"] / df["평형(평)"]).round(0)
    df["층"] = pd.to_numeric(df.get("층"), errors="coerce")
    order = ["계약일", "아파트", "법정동", "지번", "동", "층", "전용면적(㎡)", "평형(평)",
             "거래금액(만원)", "전용평당가(만원)", "건축년도", "거래유형", "해제여부"]
    df = df[[c for c in order if c in df.columns]]
    return df.sort_values("계약일", ascending=False).reset_index(drop=True)


def to_excel_bytes(df):
    """엑셀 변환 + 열 너비 자동 조정 + 제목 행 고정"""
    out = df.copy()
    out["계약일"] = out["계약일"].dt.strftime("%Y-%m-%d")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="실거래가")
        ws = writer.sheets["실거래가"]
        for col in ws.columns:
            max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)
        ws.freeze_panes = "A2"
    return buffer.getvalue()


def get_service_key():
    try:
        return st.secrets["SERVICE_KEY"]
    except Exception:
        return None


# ---------------- 화면 ----------------
st.set_page_config(page_title="아파트 실거래가 조회기", page_icon="🏢", layout="wide")
st.title("🏢 아파트 매매 실거래가 조회기")
st.caption("국토교통부 공공데이터 API로 지역·기간별 아파트 실거래가를 조회하고 엑셀로 내려받습니다.")

with st.sidebar:
    st.header("조회 설정")
    service_key = get_service_key()
    if not service_key:
        service_key = st.text_input("인증키 (Decoding)", type="password",
                                    help="공공데이터포털 마이페이지의 일반 인증키(Decoding)")

    sido = st.selectbox("시·도", list(REGIONS.keys()) + ["직접 입력"])
    if sido == "직접 입력":
        lawd_cd = st.text_input("시군구 코드 5자리", max_chars=5, placeholder="예: 11680")
        region_name = lawd_cd
    else:
        sgg = st.selectbox("시·군·구", list(REGIONS[sido].keys()))
        lawd_cd, region_name = REGIONS[sido][sgg], f"{sido} {sgg}"

    today = date.today()
    default_start = (pd.Timestamp(today) - pd.DateOffset(months=3)).date().replace(day=1)
    start = st.date_input("시작 월", value=default_start, max_value=today)
    end = st.date_input("끝 월", value=today, max_value=today)
    exclude_cancel = st.checkbox("해제(취소)된 거래 제외", value=True)
    run = st.button("조회하기", type="primary", use_container_width=True)
    st.caption("※ 실거래 신고는 계약 후 30일 이내라 최근 한 달은 건수가 적을 수 있습니다.")
    st.caption("※ 평당가는 전용면적 기준이라, 공급면적 기준인 일반 시세 정보보다 높게 나옵니다.")

if run:
    months = month_list(start, end)
    if not service_key:
        st.error("인증키를 입력해 주세요.")
    elif not (lawd_cd and lawd_cd.isdigit() and len(lawd_cd) == 5):
        st.error("시군구 코드는 숫자 5자리여야 합니다.")
    elif start > end:
        st.error("시작 월이 끝 월보다 늦습니다.")
    elif len(months) > 24:
        st.error("한 번에 최대 24개월까지 조회할 수 있습니다.")
    else:
        service_key = unquote(service_key.strip())  # Encoding 키를 넣어도 자동으로 Decoding 키로 변환
        all_rows = []
        bar = st.progress(0.0, text="조회 중...")
        try:
            for i, ym in enumerate(months, start=1):
                all_rows.extend(fetch_month(service_key, lawd_cd, ym))
                bar.progress(i / len(months), text=f"{ym[:4]}년 {ym[4:]}월 조회 중... ({i}/{len(months)})")
        except Exception as e:
            st.error(f"조회 중 오류가 났습니다: {e}")
        bar.empty()
        df = clean(all_rows)
        if not df.empty and exclude_cancel and "해제여부" in df.columns:
            df = df[df["해제여부"] != "O"]
        st.session_state["apt_df"] = df
        st.session_state["apt_title"] = f"{region_name} · {months[0][:4]}.{months[0][4:]}~{months[-1][:4]}.{months[-1][4:]}"

if "apt_df" in st.session_state:
    df = st.session_state["apt_df"]
    st.subheader(st.session_state["apt_title"])
    if df.empty:
        st.warning("조회된 거래가 없습니다. 지역 코드나 기간을 확인해 주세요.")
    else:
        # 필터
        c1, c2 = st.columns(2)
        keyword = c1.text_input("아파트 이름 검색", placeholder="예: 래미안")
        a_min, a_max = float(df["전용면적(㎡)"].min()), float(df["전용면적(㎡)"].max())
        area = c2.slider("전용면적(㎡)", a_min, a_max, (a_min, a_max)) if a_min < a_max else (a_min, a_max)
        f = df[(df["전용면적(㎡)"] >= area[0]) & (df["전용면적(㎡)"] <= area[1])]
        if keyword:
            f = f[f["아파트"].str.contains(keyword, case=False, na=False)]

        if f.empty:
            st.info("조건에 맞는 거래가 없습니다.")
        else:
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("거래 건수", f"{len(f):,}건")
            m2.metric("평균 거래가", f"{f['거래금액(만원)'].mean() / 10000:.2f}억")
            m3.metric("중위 거래가", f"{f['거래금액(만원)'].median() / 10000:.2f}억")
            m4.metric("평균 전용 평당가", f"{f['전용평당가(만원)'].mean():,.0f}만원")

            monthly = (f.set_index("계약일")["거래금액(만원)"].resample("MS").agg(["mean", "count"]))
            monthly.index = monthly.index.strftime("%Y-%m")
            g1, g2 = st.columns(2)
            g1.markdown("**월별 평균 거래가 (만원)**")
            g1.line_chart(monthly["mean"].round(0))
            g2.markdown("**월별 거래 건수**")
            g2.bar_chart(monthly["count"])

            top = (f.groupby("아파트").agg(거래건수=("거래금액(만원)", "size"),
                                         평균가_만원=("거래금액(만원)", "mean"))
                     .sort_values("거래건수", ascending=False).head(10).round(0))
            st.markdown("**거래가 많은 단지 TOP 10**")
            st.dataframe(top, use_container_width=True)

            st.markdown("**전체 거래 내역**")
            show = f.copy()
            show["계약일"] = show["계약일"].dt.strftime("%Y-%m-%d")
            st.dataframe(show, use_container_width=True, hide_index=True)

            st.download_button(
                "📥 엑셀로 내려받기", data=to_excel_bytes(f),
                file_name="apt_trades.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
