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

import base64
import io
import xml.etree.ElementTree as ET
from urllib.parse import unquote
from datetime import date

import pandas as pd
import requests
import streamlit as st

API_URL = "https://apis.data.go.kr/1613000/RTMSDataSvcAptTradeDev/getRTMSDataSvcAptTradeDev"

# 법정동 코드 앞 5자리 (시군구 코드), 행정안전부 법정동코드 2026.10.26 기준
# 이후 행정구역이 바뀐 지역은 '직접 입력'으로 새 코드를 넣어 조회
REGIONS = {
    "서울특별시": {"종로구": "11110", "중구": "11140", "용산구": "11170", "성동구": "11200", "광진구": "11215", "동대문구": "11230", "중랑구": "11260", "성북구": "11290", "강북구": "11305", "도봉구": "11320", "노원구": "11350", "은평구": "11380", "서대문구": "11410", "마포구": "11440", "양천구": "11470", "강서구": "11500", "구로구": "11530", "금천구": "11545", "영등포구": "11560", "동작구": "11590", "관악구": "11620", "서초구": "11650", "강남구": "11680", "송파구": "11710", "강동구": "11740"},
    "부산광역시": {"중구": "26110", "서구": "26140", "동구": "26170", "영도구": "26200", "부산진구": "26230", "동래구": "26260", "남구": "26290", "북구": "26320", "해운대구": "26350", "사하구": "26380", "금정구": "26410", "강서구": "26440", "연제구": "26470", "수영구": "26500", "사상구": "26530", "기장군": "26710"},
    "대구광역시": {"중구": "27110", "동구": "27140", "서구": "27170", "남구": "27200", "북구": "27230", "수성구": "27260", "달서구": "27290", "달성군": "27710", "군위군": "27720"},
    "인천광역시": {"제물포구": "28125", "영종구": "28155", "미추홀구": "28177", "연수구": "28185", "남동구": "28200", "부평구": "28237", "계양구": "28245", "서해구": "28275", "검단구": "28290", "강화군": "28710", "옹진군": "28720"},
    "대전광역시": {"동구": "30110", "중구": "30140", "서구": "30170", "유성구": "30200", "대덕구": "30230"},
    "울산광역시": {"중구": "31110", "남구": "31140", "동구": "31170", "북구": "31200", "울주군": "31710"},
    "세종특별자치시": {"세종시": "36110"},
    "경기도": {"수원시 장안구": "41111", "수원시 권선구": "41113", "수원시 팔달구": "41115", "수원시 영통구": "41117", "성남시 수정구": "41131", "성남시 중원구": "41133", "성남시 분당구": "41135", "의정부시": "41150", "안양시 만안구": "41171", "안양시 동안구": "41173", "부천시 원미구": "41192", "부천시 소사구": "41194", "부천시 오정구": "41196", "광명시": "41210", "평택시": "41220", "동두천시": "41250", "안산시 상록구": "41271", "안산시 단원구": "41273", "고양시 덕양구": "41281", "고양시 일산동구": "41285", "고양시 일산서구": "41287", "과천시": "41290", "구리시": "41310", "남양주시": "41360", "오산시": "41370", "시흥시": "41390", "군포시": "41410", "의왕시": "41430", "하남시": "41450", "용인시 처인구": "41461", "용인시 기흥구": "41463", "용인시 수지구": "41465", "파주시": "41480", "이천시": "41500", "안성시": "41550", "김포시": "41570", "화성시 만세구": "41591", "화성시 효행구": "41593", "화성시 병점구": "41595", "화성시 동탄구": "41597", "광주시": "41610", "양주시": "41630", "포천시": "41650", "여주시": "41670", "연천군": "41800", "가평군": "41820", "양평군": "41830"},
    "강원특별자치도": {"춘천시": "51110", "원주시": "51130", "강릉시": "51150", "동해시": "51170", "태백시": "51190", "속초시": "51210", "삼척시": "51230", "홍천군": "51720", "횡성군": "51730", "영월군": "51750", "평창군": "51760", "정선군": "51770", "철원군": "51780", "화천군": "51790", "양구군": "51800", "인제군": "51810", "고성군": "51820", "양양군": "51830"},
    "충청북도": {"청주시 상당구": "43111", "청주시 서원구": "43112", "청주시 흥덕구": "43113", "청주시 청원구": "43114", "충주시": "43130", "제천시": "43150", "보은군": "43720", "옥천군": "43730", "영동군": "43740", "증평군": "43745", "진천군": "43750", "괴산군": "43760", "음성군": "43770", "단양군": "43800"},
    "충청남도": {"천안시 동남구": "44131", "천안시 서북구": "44133", "공주시": "44150", "보령시": "44180", "아산시": "44200", "서산시": "44210", "논산시": "44230", "계룡시": "44250", "당진시": "44270", "금산군": "44710", "부여군": "44760", "서천군": "44770", "청양군": "44790", "홍성군": "44800", "예산군": "44810", "태안군": "44825"},
    "전북특별자치도": {"전주시 완산구": "52111", "전주시 덕진구": "52113", "군산시": "52130", "익산시": "52140", "정읍시": "52180", "남원시": "52190", "김제시": "52210", "완주군": "52710", "진안군": "52720", "무주군": "52730", "장수군": "52740", "임실군": "52750", "순창군": "52770", "고창군": "52790", "부안군": "52800"},
    "전남광주통합특별시": {"목포시": "12110", "여수시": "12130", "순천시": "12150", "나주시": "12170", "광양시": "12190", "동구": "12210", "서구": "12240", "남구": "12270", "북구": "12300", "광산구": "12330", "담양군": "12710", "곡성군": "12720", "구례군": "12730", "고흥군": "12740", "보성군": "12750", "화순군": "12760", "장흥군": "12770", "강진군": "12780", "해남군": "12790", "영암군": "12800", "무안군": "12810", "함평군": "12820", "영광군": "12830", "장성군": "12840", "완도군": "12850", "진도군": "12860", "신안군": "12870"},
    "경상북도": {"포항시 남구": "47111", "포항시 북구": "47113", "경주시": "47130", "김천시": "47150", "안동시": "47170", "구미시": "47190", "영주시": "47210", "영천시": "47230", "상주시": "47250", "문경시": "47280", "경산시": "47290", "의성군": "47730", "청송군": "47750", "영양군": "47760", "영덕군": "47770", "청도군": "47820", "고령군": "47830", "성주군": "47840", "칠곡군": "47850", "예천군": "47900", "봉화군": "47920", "울진군": "47930", "울릉군": "47940"},
    "경상남도": {"창원시 의창구": "48121", "창원시 성산구": "48123", "창원시 마산합포구": "48125", "창원시 마산회원구": "48127", "창원시 진해구": "48129", "진주시": "48170", "통영시": "48220", "사천시": "48240", "김해시": "48250", "밀양시": "48270", "거제시": "48310", "양산시": "48330", "의령군": "48720", "함안군": "48730", "창녕군": "48740", "고성군": "48820", "남해군": "48840", "하동군": "48850", "산청군": "48860", "함양군": "48870", "거창군": "48880", "합천군": "48890"},
    "제주특별자치도": {"제주시": "50110", "서귀포시": "50130"},
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
    """보기 좋게 정리된 엑셀 만들기: 열 너비·숫자 서식·제목 행 꾸미기·필터"""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    out = df.copy()
    out["계약일"] = out["계약일"].dt.strftime("%Y-%m-%d")
    for col in ["층", "건축년도"]:  # 글자로 저장돼 있던 숫자를 진짜 숫자로
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce").astype("Int64")
    # 값이 하나도 없는 열은 빼기 (예: 과거 데이터의 거래유형·해제여부)
    out = out[[c for c in out.columns if out[c].astype(str).str.strip().replace({"<NA>": "", "nan": ""}).ne("").any()]]

    number_formats = {"거래금액(만원)": "#,##0", "전용평당가(만원)": "#,##0",
                      "전용면적(㎡)": "0.00", "평형(평)": "0.0"}
    center_cols = {"계약일", "층", "건축년도", "동", "거래유형", "해제여부"}

    def text_width(v):  # 한글은 영문보다 두 배 넓게 계산
        return sum(2 if ord(ch) > 127 else 1 for ch in str(v))

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="실거래가")
        ws = writer.sheets["실거래가"]

        header_fill = PatternFill("solid", fgColor="1F4E78")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24

        for idx, name in enumerate(out.columns, start=1):
            letter = get_column_letter(idx)
            values = [name] + out[name].astype(str).head(500).tolist()
            ws.column_dimensions[letter].width = min(max(text_width(v) for v in values) + 3, 45)
            fmt = number_formats.get(name)
            for cell in ws[letter][1:]:
                if fmt:
                    cell.number_format = fmt
                if name in center_cols:
                    cell.alignment = Alignment(horizontal="center")

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions  # 제목 행에 필터 버튼
    return buffer.getvalue()


def get_service_key():
    try:
        return st.secrets["SERVICE_KEY"]
    except Exception:
        return None



# ---------------- 상단 로고 배너 (누르면 처음 화면으로) ----------------
BANNER_SVG = """<svg xmlns='http://www.w3.org/2000/svg' width='900' height='120' viewBox='0 0 900 120'>
<defs><linearGradient id='g' x1='0' y1='0' x2='1' y2='0'>
<stop offset='0' stop-color='#1F4E78'/><stop offset='1' stop-color='#2E86C1'/></linearGradient></defs>
<rect width='900' height='120' rx='18' fill='url(#g)'/>
<g fill='#ffffff' opacity='0.95'>
<rect x='34' y='38' width='26' height='58' rx='3'/><rect x='64' y='22' width='34' height='74' rx='3'/><rect x='102' y='48' width='22' height='48' rx='3'/>
</g>
<g fill='#2E86C1'>
<rect x='39' y='46' width='6' height='6'/><rect x='49' y='46' width='6' height='6'/><rect x='39' y='58' width='6' height='6'/><rect x='49' y='58' width='6' height='6'/>
<rect x='70' y='30' width='7' height='7'/><rect x='84' y='30' width='7' height='7'/><rect x='70' y='44' width='7' height='7'/><rect x='84' y='44' width='7' height='7'/>
<rect x='70' y='58' width='7' height='7'/><rect x='84' y='58' width='7' height='7'/><rect x='107' y='56' width='5' height='5'/><rect x='115' y='56' width='5' height='5'/>
</g>
<text x='150' y='62' font-family='Pretendard, Noto Sans KR, Malgun Gothic, sans-serif' font-size='34' font-weight='800' fill='#ffffff'>아파트 매매 실거래가 조회기</text>
<text x='152' y='92' font-family='Pretendard, Noto Sans KR, Malgun Gothic, sans-serif' font-size='16' fill='#D6EAF8'>국토교통부 공공데이터 · 지역·기간별 실거래가 조회 · 엑셀 다운로드</text>
</svg>"""


def show_banner():
    """배너 이미지를 링크로 감싸서, 누르면 앱 첫 주소로 새로 열림 (모든 설정·결과 초기화)"""
    b64 = base64.b64encode(BANNER_SVG.encode("utf-8")).decode()
    st.markdown(
        f'<a href="./" target="_self" title="처음 화면으로">'
        f'<img src="data:image/svg+xml;base64,{b64}" style="width:100%;max-width:900px;border-radius:18px;"></a>',
        unsafe_allow_html=True)


def show_home():
    """조회 전 처음 화면"""
    st.markdown("### 👋 어서 오세요")
    st.write("전국 아파트의 **실제 매매 거래가격**을 지역과 기간별로 확인하고, 정리된 엑셀로 내려받을 수 있습니다.")
    c1, c2, c3 = st.columns(3)
    with c1.container(border=True):
        st.markdown("**① 지역 고르기**")
        st.caption("왼쪽에서 시·도와 시·군·구를 선택하세요. 2026년 최신 행정구역 기준입니다.")
    with c2.container(border=True):
        st.markdown("**② 기간 정하고 조회**")
        st.caption("2006년부터 고를 수 있고, 한 번에 최대 24개월까지 조회됩니다.")
    with c3.container(border=True):
        st.markdown("**③ 분석하고 저장**")
        st.caption("평균가·평당가·월별 추이를 보고, 단지별로 골라 엑셀로 내려받으세요.")
    st.info("📱 휴대폰에서는 왼쪽 위 **>** 버튼을 눌러 조회 설정을 여세요. 위쪽 로고를 누르면 언제든 이 화면으로 돌아옵니다.")
    st.caption("※ 개인정보 보호를 위해 호수는 공개되지 않으며, 동 정보는 등기가 끝난 거래에만 표시됩니다.")


# ---------------- 화면 ----------------
st.set_page_config(page_title="아파트 실거래가 조회기", page_icon="🏢", layout="wide")
show_banner()

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
    start = st.date_input("시작 월", value=default_start, min_value=date(2006, 1, 1), max_value=today)
    end = st.date_input("끝 월", value=today, min_value=date(2006, 1, 1), max_value=today)
    exclude_cancel = st.checkbox("해제(취소)된 거래 제외", value=True)
    run = st.button("조회하기", type="primary", use_container_width=True)
    debug = st.checkbox("API 응답 원문 확인 (문제 확인용)", value=False)
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
        if debug:
            try:
                key = service_key
                raw = requests.get(API_URL, params={"serviceKey": key, "LAWD_CD": lawd_cd,
                                   "DEAL_YMD": months[0], "pageNo": 1, "numOfRows": 10}, timeout=20)
                with st.expander("API 응답 원문 (첫 달, 최대 10건)", expanded=True):
                    st.write(f"요청 지역 코드: {lawd_cd} / 계약년월: {months[0]} / HTTP 상태: {raw.status_code}")
                    st.code(raw.text[:3000], language="xml")
            except Exception as e:
                st.error(f"원문 확인 중 오류: {e}")
        df = clean(all_rows)
        if not df.empty and exclude_cancel and "해제여부" in df.columns:
            df = df[df["해제여부"] != "O"]
        st.session_state["apt_df"] = df
        st.session_state["apt_title"] = f"{region_name} · {months[0][:4]}.{months[0][4:]}~{months[-1][:4]}.{months[-1][4:]}"

if "apt_df" not in st.session_state:
    show_home()
else:
    df = st.session_state["apt_df"]
    st.subheader(st.session_state["apt_title"])
    if df.empty:
        st.warning("조회된 거래가 없습니다. 지역 코드나 기간을 확인해 주세요.")
    else:
        # 필터
        c1, c2 = st.columns(2)
        names = sorted(df["아파트"].dropna().unique().tolist())
        picked = c1.multiselect(
            f"단지 선택 (이 기간 거래된 단지 {len(names)}곳)", names,
            placeholder="이름 일부를 입력해 찾기 · 비우면 전체",
            help="조회 기간에 실제 거래가 있었던 단지만 목록에 나옵니다. 목록에 없으면 그 기간 거래가 없었다는 뜻입니다.")
        a_min, a_max = float(df["전용면적(㎡)"].min()), float(df["전용면적(㎡)"].max())
        area = c2.slider("전용면적(㎡)", a_min, a_max, (a_min, a_max)) if a_min < a_max else (a_min, a_max)
        f = df[(df["전용면적(㎡)"] >= area[0]) & (df["전용면적(㎡)"] <= area[1])]
        if picked:
            f = f[f["아파트"].isin(picked)]

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
