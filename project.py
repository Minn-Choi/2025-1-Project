import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from tkinter import Tk, filedialog
from tkinter.simpledialog import askinteger
from tkinter.messagebox import showinfo


def select_file():
    """사용자 파일 선택"""
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="엑셀 파일 선택",
        filetypes=[("Excel Files", "*.xls;*.xlsx")]
    )
    return file_path


def get_user_input(prompt):
    """사용자로부터 정수 입력 받기"""
    return askinteger("입력 요청", prompt)


def get_user_inputs_for_dict(keys, title):
    """사용자로부터 키별 값을 입력받아 딕셔너리로 반환"""
    values = {}
    for key in keys:
        value = askinteger(f"{title}", f"{key}의 정원을 입력하세요:") or 0
        values[key] = value
    return values


def preprocess_data(input_df):
    """데이터 전처리"""
    department_mapping = {
        "85": "임원",
        "84": "안전감사실",
        "81": "전략기획실",
        "82": "경영지원부",
        "83": "체육사업부",
        "86": "주차사업부",
        "91": "시설관리부",
        "89": "사회서비스단",
    }
    input_df["부서.1"] = input_df["부서"].astype(str).str[:2].map(department_mapping)

    input_df["직급"] = input_df["직급"].replace(
        {"이사장": "임원", "본부장": "임원", "체력측정사": "전문지도직", "수영강사": "전문지도직",
         "헬스강사": "전문지도직", "테니스강사": "전문지도직", "운동처방사": "전문지도직"}
    )

    input_df = input_df[~input_df["직급"].isin(["기간제근로", "휴직대체(7"])]
    return input_df


def custom_sort(data, order):
    """사용자 정의 순서를 기반으로 데이터 정렬"""
    return sorted(data.items(), key=lambda x: order.index(x[0]) if x[0] in order else len(order))


def process_data(file_path):
    """사용자 파일에서 부서.1별, 직급별 현원을 계산"""
    input_df = pd.read_excel(file_path)

    input_df = preprocess_data(input_df)

    if "부서.1" not in input_df.columns or "직급" not in input_df.columns:
        print("파일에 '부서.1' 또는 '직급' 열이 없습니다.")
        return

    department_counts = input_df["부서.1"].value_counts()

    position_counts = input_df["직급"].value_counts()

    return department_counts, position_counts


def create_excel_file(department_counts, position_counts, total_quota, position_quota, department_quota, detailed_quota):
    """엑셀 파일 생성"""
    department_order = ["임원", "안전감사실", "전략기획실", "경영지원부", "체육사업부", "주차사업부", "시설관리부", "사회서비스단"]
    position_order = ["임원","3급","4급","5급","6급","7급","전문지도직", "시설안내원", "주차관리원", "환경관리원", "사무보조직"]

    wb = Workbook()
    ws = wb.active
    ws.title = "정원 및 현황"

    ws.append(["구분", "항목", "현원", "정원", "과부족"])

    total_current = sum(department_counts.values)
    total_surplus_deficit = total_current - total_quota
    ws.append(["전체", "전체 정원", total_current, total_quota, total_surplus_deficit])

    ws.append(["부서별", None, None, None, None])
    for department, count in custom_sort(department_counts, department_order):
        quota = department_quota.get(department, 0)
        surplus_deficit = count - quota
        ws.append(["", department, count, quota, surplus_deficit])

    ws.append(["직급별", None, None, None, None])
    for position, count in custom_sort(position_counts, position_order):
        quota = position_quota.get(position, 0)
        surplus_deficit = count - quota
        ws.append(["", position, count, quota, surplus_deficit])

    ws.append(["부서별 세부 정원", None, None, None, None])
    for department, allocation in custom_sort(detailed_quota, department_order):
        ws.append(["", department, None, None, None])
        for category, value in sorted(allocation.items()):
            ws.append(["", f"  {category}", None, value, None])

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=Side(style="thin"),
                                 right=Side(style="thin"),
                                 top=Side(style="thin"),
                                 bottom=Side(style="thin"))
            if cell.row == 1 or cell.column == 1:
                cell.font = Font(bold=True)
            if cell.column == 5 and isinstance(cell.value, (int, float)):
                if cell.value < 0:
                    cell.fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")  # 과부족 음수
                elif cell.value > 0:
                    cell.fill = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")  # 과부족 양수

    wb.save("정원_현황.xlsx")
    print("정원 현황 파일이 생성되었습니다: 정원_현황.xlsx")


def main():
    file_path = select_file()
    if not file_path:
        print("파일이 선택되지 않았습니다.")
        return

    department_counts, position_counts = process_data(file_path)

    total_quota = get_user_input("전체 정원의 수를 입력하세요:")
    position_quota = get_user_inputs_for_dict(position_counts.keys(), "직급별 정원 입력")
    department_quota = get_user_inputs_for_dict(department_counts.keys(), "부서별 정원 입력")

    detailed_quota = {}
    for department in department_counts.keys():
        showinfo("부서별 세부 정원", f"{department}에 대해 정원을 입력합니다.")
        detailed_quota[department] = get_user_inputs_for_dict(
            ["임원~7급", "시설안내원~사무보조"],
            f"{department}의 세부 정원 입력"
        )

    create_excel_file(department_counts, position_counts, total_quota, position_quota, department_quota, detailed_quota)


if __name__ == "__main__":
    main()