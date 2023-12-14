import streamlit as st
from deta import Deta
import datetime
import pandas as pd
import io
import openpyxl
import matplotlib.pyplot as plt
from matplotlib import font_manager, rc
from dotenv import load_dotenv
import os

# 환경 변수에서 DETA_PROJECT_KEY를 로드
# .env 파일에서 환경 변수 로드
load_dotenv()

DETA_PROJECT_KEY = os.getenv("DETA_PROJECT_KEY")
if DETA_PROJECT_KEY is None:
    raise Exception("DETA_PROJECT_KEY 환경 변수가 설정되지 않았습니다.")
deta = Deta(DETA_PROJECT_KEY)
db = deta.Base("count")

# 데이터베이스 CRUD 함수
def insert_record(data):
    # '지출일자' 키가 있는지 확인하고 문자열로 변환
    if "2.지출일자" in data:
        data["2.지출일자"] = data["2.지출일자"].isoformat()
    else:
        raise ValueError("날짜 정보가 누락되었습니다.")

    # '원천징수 대상자' 키가 없으면 빈 딕셔너리로 초기화
    if "6.원천징수 대상자" not in data:
        data["6.원천징수 대상자"] = {}

    # 데이터베이스에 기록 추가
    return db.put(data)

def get_record(key):
    return db.get(key)

def fetch_records():
    return db.fetch().items

def update_record(key, updates):
    return db.update(updates, key)

def delete_record(key):
    return db.delete(key)

def filter_records(branch, account_title, start_date, end_date):
    all_records = db.fetch().items
    filtered_records = [
        record for record in all_records
        if record.get("1.지점명") == branch and
            (account_title == "모든 계정과목" or record.get("3.계정과목") == account_title) and
            start_date <= datetime.datetime.fromisoformat(record.get("2.지출일자")).date() <= end_date
    ]
    return filtered_records

# 지점명 필터링 함수
def filter_records(branch, account_title, start_date, end_date):
    all_records = db.fetch().items
    if branch != "모든 지점":
        filtered_records = [
            record for record in all_records
            if record.get("1.지점명") == branch and
               (account_title == "모든 계정과목" or record.get("3.계정과목") == account_title) and
               start_date <= datetime.datetime.fromisoformat(record.get("2.지출일자")).date() <= end_date
        ]
    else:
        filtered_records = [
            record for record in all_records
            if (account_title == "모든 계정과목" or record.get("3.계정과목") == account_title) and
               start_date <= datetime.datetime.fromisoformat(record.get("2.지출일자")).date() <= end_date
        ]
    return filtered_records


def visualize_data(records):
    # 데이터 분석 및 시각화
    df = pd.DataFrame(records)
    if not df.empty:
        # '3.계정과목'에 따른 '5.총출금액(원천세 제외)'의 합계 계산
        summary = df.groupby('3.계정과목')['5.총출금액(원천세 제외)'].sum()
        summary.plot(kind='bar')
        plt.title('계정과목별 총 지출')
        plt.xlabel('계정과목')
        plt.ylabel('총 지출금액')
        plt.xticks(rotation=45)
        st.pyplot(plt)
    else:
        st.write("데이터가 없습니다.")

def download_excel(data):
    # 데이터를 판다스 DataFrame으로 변환
    df = pd.DataFrame(data)

    # 엑셀 파일로 변환
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
        # writer.save() 호출이 필요 없음

    # 파일 포인터를 처음으로 되돌림
    output.seek(0)
    return output


# Streamlit 애플리케이션
def main():
    st.title("💸소종사업 회계 관리 시스템")

    # 서울 자치구 목록
    seoul_districts = [
        "강남지점", "강동지점", "강북지점", "강서지점", "관악지점",
        "광진지점", "구로지점", "금천지점", "노원지점", "도봉지점",
        "동대문지점", "동작지점", "마포지점", "서대문지점", "서초지점",
        "성동지점", "성북지점", "송파지점", "양천지점", "영등포지점",
        "용산지점", "은평지점", "종로지점", "명동지점", "중랑지점"
    ]

    # 지점명 입력 및 자동완성 기능
    user_input = st.text_input("지점명 입력")
    if user_input:
        # 사용자 입력에 따라 지점명 필터링
        filtered_districts = [district for district in seoul_districts if user_input in district]

        # 필터링된 지점명을 선택할 수 있는 selectbox 제공
        if filtered_districts:
            branch_name = st.selectbox("지점명 선택", filtered_districts)
        else:
            st.write("일치하는 지점명이 없습니다.")

    # 계정과목 및 예산 코드
    account_titles = {
        "자영업지원센터 운영": ["조사연구비", "광고선전비"],
        "종합지원 포털 운영": ["종합지원 포털 서비스 유지관리", "종합지원 포털 서비스 고도화"],
        "우리마을가게 상권분석서비스": ["우리마을가게 고도화", "우리마을가게 유지관리"],
        "소상공인 역량강화": ["온라인 교육 시스템", "소상공인 교육", "현장멘토링"],
        "자영업 클리닉 지원": ["자영업클리닉 컨설팅"],
        "위탁관리수수료": ["위탁관리수수료"],
        "위기 소상공인 조기발굴 및 선제지원": ["위기 소상공인 컨설팅", "위기 소상공인 이행비용"],
        "중장년 소상공인 디지털 전환": ["컨설팅 비용", "디지털 전환 교육", "디지털 전환비용", "디지털 정착비용"],
        "소상공인 사업재기 및 안전한 폐업지원": ["사업재기 컨설팅", "사업재기 폐업지원금"],
        "외부전문가 구성 및 운영": ["업종닥터 운영비", "외부전문가 교육비", "우수 멘토 행사", "디지털 전환 운영비"],
        "서울형 다시서기 4.0 프로젝트": ["재도전 씨앗자금", "자영업클리닉(다시서기)"]
    }

    # 계정과목 선택
    account_title = st.selectbox("계정과목 선택", list(account_titles.keys()))

    # 선택된 계정과목에 따라 귀속코드 업데이트
    budget_codes = account_titles[account_title]
    budget_code = st.selectbox("귀속코드 선택", budget_codes)

    if 'withholding_tax' not in st.session_state:
        st.session_state['withholding_tax'] = False

    if 'confirm_submit' not in st.session_state:
        st.session_state['confirm_submit'] = False

    withholding_tax = st.checkbox("✔️원천징수 여부(컨설팅비용 등 개인에게 수당 지급하면서, 그 금액이 125,000원 이상인 경우)", value=st.session_state['withholding_tax'], key='withholding_tax')
    # 지출 결의서 입력 폼
    with st.form("expense_form"):

        # 초기화: 만약 session_state에 해당 키가 없다면 초기값 설정
        if 'selected_branch' not in st.session_state:
            st.session_state['selected_branch'] = seoul_districts[0] # 첫 번째 지점으로 초기화
        if 'selected_account_title' not in st.session_state:
            st.session_state['selected_account_title'] = list(account_titles.keys())[0] # 첫 번째 계정과목으로 초기화
        if 'start_date' not in st.session_state:
            st.session_state['start_date'] = datetime.date.today() - datetime.timedelta(days=30)
        if 'end_date' not in st.session_state:
            st.session_state['end_date'] = datetime.date.today()

        # 날짜 선택 (기본값: 오늘 날짜)
        date = st.date_input("날짜 선택", datetime.date.today())
        amount = st.number_input("금액", min_value=0)

        withholding_names = []
        names_input = ""

        if st.session_state['withholding_tax']:
            names_input = st.text_area("원천징수 대상자 이름 (쉼표로 구분, 컨트롤+엔터로 입력하실 것)", "홍길동,임꺽정")
            withholding_names = []
            withholding_amounts = {}

        if names_input:
            withholding_names = [name.strip() for name in names_input.split(',')]
            for name in withholding_names:
                withholding_amounts[name] = st.number_input(f"[{name}] 원천징수액", min_value=0, key=name)

        # 원천징수액을 설명에 추가
        default_description_lines = [
            f"- 사업접수번호 :",
            *[f"[{name}] 원천징수액 : {withholding_amounts.get(name, '')}" for name in withholding_names]
        ]
        default_description = "\n".join(default_description_lines)

        description = st.text_area("상세 설명(사업접수번호 및 원천징수 상세내역 입력)", default_description)
        submit_button = st.form_submit_button("제출")
        if submit_button and not st.session_state['confirm_submit']:
            st.session_state['confirm_submit'] = True

    # 원천징수 여부 확인 후 데이터 제출 처리
    if st.session_state['confirm_submit']:
        confirm = st.radio("원천징수 여부를 확인하셨습니까?", ('','예', '아니오'))
        income_tax = amount * 0.08 if withholding_tax else 0
        local_tax = amount * 0.008 if withholding_tax else 0
        net_amount = amount - income_tax - local_tax
        if confirm == '예':
            # 데이터베이스에 기록 추가
            record = {
                "1.지점명": branch_name,
                "2.지출일자": date,
                "3.계정과목": account_title,
                "4.예산귀속코드": budget_code,
                "5.총출금액(원천세 제외)": net_amount,
                "6.원천징수 대상자": [
                    {"이름": name, "원천징수액": withholding_amounts.get(name, 0)} for name in withholding_names
                ],
                "7.기타소득세": income_tax,
                "8.기타지방소득세": local_tax,
                "9.상세설명": description
            }
            try:
                insert_record(record)
                st.success("기록이 성공적으로 저장되었습니다.")
                st.session_state['confirm_submit'] = False  # 상태 초기화
            except Exception as e:
                st.error(f"기록 저장에 실패했습니다: {e}")
                st.session_state['confirm_submit'] = False  # 상태 초기화
        elif confirm == '아니오':
            st.warning("제출을 중단합니다. 원천징수 여부를 확인해주세요.")
            st.session_state['confirm_submit'] = False  # 상태 초기화

    # 기록 조회 및 수정/삭제
    st.header("입력한 내용 불러오기")
    with st.expander("불러오기 옵션"):
        st.session_state['selected_branch'] = st.selectbox("지점명 선택", seoul_districts, key="branch_select", index=seoul_districts.index(st.session_state['selected_branch']))
        all_account_titles = ["모든 계정과목"] + list(account_titles.keys())
        st.session_state['selected_account_title'] = st.selectbox("계정과목 선택", all_account_titles, key="account_title_select", index=all_account_titles.index(st.session_state['selected_account_title']))
        st.session_state['start_date'] = st.date_input("시작 날짜", st.session_state['start_date'], key="start_date_select")
        st.session_state['end_date'] = st.date_input("종료 날짜", st.session_state['end_date'], key="end_date_select")
        search_button = st.button("조회", key="search_button")

    if search_button:
        st.session_state['filtered_records'] = filter_records(st.session_state['selected_branch'], st.session_state['selected_account_title'], st.session_state['start_date'], st.session_state['end_date'])
        df = pd.DataFrame(st.session_state['filtered_records'])
        towrite = io.BytesIO()
        df.to_excel(towrite, index=False, engine='openpyxl')  # index=False로 설정하여 인덱스 제외
        towrite.seek(0)
        st.dataframe(df)  # 데이터프레임 표시
        st.download_button(label="엑셀로 다운로드",
                        data=towrite,
                        file_name="filtered_records.xlsx",
                        mime="application/vnd.ms-excel")
    # 시각화 버튼
    if st.button('시각화') and 'filtered_records' in st.session_state:
        visualize_data(st.session_state['filtered_records'])
        # 다운로드 버튼 생성


    st.title("기간별 / 지점별 원천징수 명세 불러오기")


    # 원천징수 대상자 세액 데이터 다운로드 버튼 클릭 시 세션 초기화
    # 세션 상태 초기화
    if 'start_date_withholding' not in st.session_state:
        st.session_state['start_date_withholding'] = datetime.date.today() - datetime.timedelta(days=30)
    if 'end_date_withholding' not in st.session_state:
        st.session_state['end_date_withholding'] = datetime.date.today()

    # 폼을 사용한 입력 필드
    with st.form("my_form"):
        start_date = st.date_input("시작 날짜", st.session_state['start_date_withholding'], key="start_date_withholding")
        end_date = st.date_input("종료 날짜", st.session_state['end_date_withholding'], key="end_date_withholding")
        selected_branch = st.selectbox("지점명 선택", ["모든 지점"] + seoul_districts)  # seoul_districts는 사전에 정의된 지역 목록

        # '제출' 버튼
        submitted = st.form_submit_button("원천징수 대상자 세액 데이터 불러오기")

    if submitted:
        filtered_records = filter_records(selected_branch, st.session_state['selected_account_title'], start_date, end_date)
        fetched_data = []
        for record in filtered_records:
            fetched_record = db.get(record["key"])
            withholding_data = fetched_record.get("6.원천징수 대상자", {})
            fetched_data.append({
                "1.지점명": fetched_record["1.지점명"],
                "2.지출일자": fetched_record["2.지출일자"],
                "5.총출금액(원천세 제외)": fetched_record["5.총출금액(원천세 제외)"],
                "3.계정과목": fetched_record["3.계정과목"],
                "6.원천징수 대상자": withholding_data
            })

        if fetched_data:
            expanded_records = []
            for record in fetched_data:
                withholding_data = record["6.원천징수 대상자"]

                # withholding_data가 리스트인지 확인
                if isinstance(withholding_data, list):
                    for item in withholding_data:
                        expanded_record = record.copy()
                        expanded_record["원천징수 대상자 이름"] = item.get("이름")
                        expanded_record["원천징수 금액"] = item.get("원천징수액")
                        expanded_records.append(expanded_record)
                else:
                    # withholding_data가 리스트가 아니거나 비어 있는 경우의 처리
                    # 예: expanded_records.append({...})
                    pass

            st.dataframe(expanded_records)
            
            excel_output = download_excel(expanded_records)
            st.download_button(label="원천징수 대상자 세액 엑셀 다운로드",
                            data=excel_output.getvalue(),  # 변경된 부분
                            file_name="원천징수 명세.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️선택한 기간 및 지점에 원천징수 대상자 데이터가 없습니다.")

if __name__ == "__main__":
    main()

