import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from tkinter import Tk, Toplevel, Label, Entry, Button, filedialog, Checkbutton, BooleanVar, Listbox, END, Frame
from openpyxl.chart import BarChart, Reference
import datetime

def select_file():
    """사용자 파일 선택"""
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="엑셀 파일 선택",
        filetypes=[("Excel Files", "*.xls;*.xlsx")]
    )
    return file_path

def get_sorting_order(departments, positions):
    """정렬 순서를 선택하거나 사용자 지정으로 입력받는 함수"""
    root = Toplevel()
    root.title("정렬 순서 선택")
    root.geometry("500x500")  # 창 크기 설정

    use_custom_order = BooleanVar()
    use_custom_order.set(False)

    Label(root, text="정렬 방식을 선택하세요", font=("Arial", 14)).pack(pady=10)
    Checkbutton(root, text="사용자 지정 정렬", variable=use_custom_order).pack()

    # 사용자 지정 정렬을 위한 프레임
    custom_order_frame = Frame(root)
    custom_order_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # 부서 순서
    Label(custom_order_frame, text="부서 순서").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    department_listbox = Listbox(custom_order_frame, selectmode="extended", exportselection=False, height=10)
    for department in departments:
        department_listbox.insert(END, department)
    department_listbox.grid(row=1, column=0, padx=5, pady=5)

    Label(custom_order_frame, text="직급 순서").grid(row=0, column=1, padx=5, pady=5, sticky="w")
    position_listbox = Listbox(custom_order_frame, selectmode="extended", exportselection=False, height=10)
    for position in positions:
        position_listbox.insert(END, position)
    position_listbox.grid(row=1, column=1, padx=5, pady=5)

    def move_up(listbox):
        selection = listbox.curselection()
        if not selection or selection[0] == 0:
            return
        for index in selection:
            value = listbox.get(index)
            listbox.delete(index)
            listbox.insert(index - 1, value)
            listbox.selection_set(index - 1)

    def move_down(listbox):
        selection = listbox.curselection()
        if not selection or selection[-1] == listbox.size() - 1:
            return
        for index in reversed(selection):
            value = listbox.get(index)
            listbox.delete(index)
            listbox.insert(index + 1, value)
            listbox.selection_set(index + 1)

    Button(custom_order_frame, text="⬆️ 위로", command=lambda: move_up(department_listbox)).grid(row=2, column=0, pady=5)
    Button(custom_order_frame, text="⬇️ 아래로", command=lambda: move_down(department_listbox)).grid(row=3, column=0, pady=5)

    Button(custom_order_frame, text="⬆️ 위로", command=lambda: move_up(position_listbox)).grid(row=2, column=1, pady=5)
    Button(custom_order_frame, text="⬇️ 아래로", command=lambda: move_down(position_listbox)).grid(row=3, column=1, pady=5)

    custom_order = {"department": [], "position": []}

    def submit():
        if use_custom_order.get():
            custom_order["department"] = [department_listbox.get(i) for i in range(department_listbox.size())]
            custom_order["position"] = [position_listbox.get(i) for i in range(position_listbox.size())]
        root.quit()  # mainloop 종료
        root.destroy()  # 창 닫기

    Button(root, text="확인", command=submit).pack(pady=10)

    root.mainloop()
    return use_custom_order.get(), custom_order

def get_quota_input(departments, positions):
    """부서 및 직급별 정원을 한 화면에서 입력받는 함수"""
    root = Toplevel()
    root.title("정원 입력")

    entries = {}
    row = 0

    Label(root, text="정원을 입력하세요", font=("Arial", 14)).grid(row=row, column=0, columnspan=3 + len(positions), pady=10)
    row += 1

    Label(root, text="부서").grid(row=row, column=0, padx=10, pady=5)
    for col, position in enumerate(positions, start=1):
        Label(root, text=position).grid(row=row, column=col, padx=10, pady=5)
    row += 1

    for department in departments:
        Label(root, text=department).grid(row=row, column=0, padx=10, pady=5, sticky="e")
        for col, position in enumerate(positions, start=1):
            entry = Entry(root)
            entry.grid(row=row, column=col, padx=10, pady=5)
            entries[(department, position)] = entry
        row += 1

    detailed_quota = {}

    def submit():
        nonlocal detailed_quota
        for (department, position), entry in entries.items():
            value = entry.get()
            if department not in detailed_quota:
                detailed_quota[department] = {}
            detailed_quota[department][position] = int(value) if value.isdigit() else 0
        root.destroy()

    Button(root, text="입력 완료", command=submit).grid(row=row, column=0, columnspan=3 + len(positions), pady=10)
    root.wait_window()
    return detailed_quota

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

    sub_department_mapping = {
        "8100": "전략기획부_부장",
        "8101": "기획조정팀",
        "8103": "소통경영팀",
        "8200": "경영지원부_부장",
        "8201": "인사총무팀",
        "8202": "재정지원팀",
        "8300": "체육사업부_부장",
        "8301": "체육사업1팀",
        "8302": "체육사업2팀",
        "8352": "충무",
        "8351": "무학봉",
        "8358": "훈련원공원",
        "8354": "회현",
        "8353": "손기정",
        "8361": "남산",
        "8360": "장충",
        "8400": "안전감사실_부장",
        "8401": "청렴감사팀",
        "8402": "안전보건팀",
        "8600": "주차사업부_부장",
        "8601": "주차사업1팀",
        "8602": "주차사업2팀",
        "8900": "사회서비스단",
        "9100": "시설관리부_부장",
        "9101": "시설관리1팀",
        "9102": "시설관리2팀",
        "9103": "공공시설팀",
        "9155": "중구구민회관",
        "9157": "공중공원화장실",
        "9152": "중구종합복지센터",
        "9153": "신당누리센터",
        "9154": "교육지원센터",
        "9151": "중립종합복지센터",
        "9156": "충무창업큐브",
    }

    input_df["부서.1"] = input_df["부서"].astype(str).str[:2].map(department_mapping)

    input_df["세부부서"] = input_df["부서"].astype(str).map(sub_department_mapping).fillna("기타")

    input_df["직급"] = input_df["직급"].replace(
        {"이사장": "임원", "본부장": "임원", "수영강사": "전문지도직", "헬스강사": "전문지도직", "테니스강사": "전문지도직"}
    )

    input_df = input_df[~input_df["직급"].isin(["기간제근로", "휴직대체(7", "운동처방사", "체력측정사"])]

    return input_df

def traverse_hierarchy_and_count(hierarchy, input_df):
    updated_hierarchy = {}

    for code, data in hierarchy.items():
        sub_departments = data["sub_departments"]
        current_count = input_df[input_df["부서"].astype(str) == code].shape[0]
        sub_department_data = traverse_hierarchy_and_count(sub_departments, input_df)
        sub_total_count = sum(sub_data["현원"] for sub_data in sub_department_data.values())
        total_count = current_count + sub_total_count
        updated_hierarchy[code] = {
            "name": data["name"],
            "현원": total_count,
            "sub_departments": sub_department_data
        }
    return updated_hierarchy

def write_hierarchy_to_excel(ws, hierarchy, parent=None):
    """계층 구조 데이터를 엑셀에 기록"""
    for code, data in hierarchy.items():
        name = data["name"]
        current_count = data["현원"]
        sub_departments = data["sub_departments"]
        ws.append([parent, name, current_count])
        write_hierarchy_to_excel(ws, sub_departments, parent=name)

def create_excel_file(
    department_counts,
    position_counts,
    grouped_counts,
    total_quota,
    position_quota,
    department_quota,
    detailed_quota,
    sub_department_counts,
    department_hierarchy,
    input_df,
    department_order,  # 사용자 지정 부서 순서
    position_order     # 사용자 지정 직급 순서
):
    """엑셀 파일 생성 및 데이터 기록"""

    for department, positions in detailed_quota.items():
        department_quota[department] = sum(positions.values())

    # 직급별 총 정원 계산
    for position in position_order:
        position_quota[position] = sum(
            detailed_quota[department].get(position, 0) for department in detailed_quota
        )

    wb = Workbook()
    ws_main = wb.active
    ws_main.title = "정원 및 현황"

    ws_main.append(["구분", "항목", "현원", "정원", "과부족"])

    total_current = sum(department_counts.values)
    total_quota = sum(department_quota.values())
    total_surplus_deficit = total_current - total_quota
    ws_main.append(["전체", "전체 정원", total_current, total_quota, total_surplus_deficit])

    ws_main.append(["부서별", None, None, None, None])
    start_row = ws_main.max_row + 1  # 그래프 데이터의 시작 위치
    for department, count in sorted(department_counts.items(), key=lambda x: department_order.index(x[0])):
        quota = department_quota.get(department, 0)
        surplus_deficit = count - quota
        ws_main.append(["", department, count, quota, surplus_deficit])

    ws_main.append(["직급별", None, None, None, None])
    for position, count in sorted(position_counts.items(), key=lambda x: position_order.index(x[0])):
        quota = position_quota.get(position, 0)
        surplus_deficit = count - quota
        ws_main.append(["", position, count, quota, surplus_deficit])

    ws_main.append(["부서별 세부 정원 및 현원", None, None, None, None])
    for department, allocation in detailed_quota.items():
        ws_main.append(["", department, None, None, None])
        for category, quota in allocation.items():
            current_count = grouped_counts.get((department, category), 0)  # 부서-직급 현원 가져오기
            surplus_deficit = current_count - quota
            ws_main.append(["", f"  {category}", current_count, quota, surplus_deficit])

    chart = BarChart()
    chart.title = "부서별 현원 및 정원"
    chart.x_axis.title = "부서"
    chart.y_axis.title = "인원"

    data = Reference(ws_main, min_col=3, max_col=4, min_row=start_row - 1, max_row=start_row + len(department_counts) - 1)
    categories = Reference(ws_main, min_col=2, min_row=start_row, max_row=start_row + len(department_counts) - 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    ws_main.add_chart(chart, "G2")

    ws_hierarchy = wb.create_sheet(title="부서 계층 구조")
    ws_hierarchy.append(["상위 부서", "하위 부서", "현원"])

    updated_hierarchy = traverse_hierarchy_and_count(department_hierarchy, input_df)
    write_hierarchy_to_excel(ws_hierarchy, updated_hierarchy)

    for ws in [ws_main, ws_hierarchy]:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=5):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(left=Side(style="thin"),
                                     right=Side(style="thin"),
                                     top=Side(style="thin"),
                                     bottom=Side(style="thin"))
                if cell.row == 1:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")

    file_name = f"정원_현황_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(file_name)
    print(f"엑셀 파일이 생성되었습니다: {file_name}")

def main():
    file_path = select_file()
    if not file_path:
        print("파일이 선택되지 않았습니다.")
        return

    input_df = pd.read_excel(file_path)
    print("원본 데이터 열:", input_df.columns)  # 원본 데이터 열 확인

    input_df = preprocess_data(input_df)
    print("전처리 후 데이터 열:", input_df.columns) 
    print(input_df[["부서", "부서.1", "세부부서"]].head())  

    if "세부부서" not in input_df.columns:
        print("세부부서 열이 없습니다. 데이터 전처리를 확인하세요.")
        return

    department_counts = input_df["부서.1"].value_counts()
    position_counts = input_df["직급"].value_counts()
    grouped_counts = input_df.groupby(["부서.1", "직급"]).size()
    sub_department_counts = input_df["세부부서"].value_counts()

    total_quota = 100 
    departments = department_counts.index.tolist()
    positions = position_counts.index.tolist()

    use_custom_order, custom_order = get_sorting_order(departments, positions)

    if use_custom_order:
        departments = custom_order["department"]
        positions = custom_order["position"]

    detailed_quota = get_quota_input(departments, positions)

    department_hierarchy = {
     "85": {"name": "임원", "sub_departments": {}},
    "84": {
        "name": "안전감사실",
        "sub_departments": {
            "8400": {"name": "안전감사실_부장", "sub_departments": {}},
            "8401": {"name": "청렴감사팀", "sub_departments": {}},
            "8402": {"name": "안전보건팀", "sub_departments": {}}
        }
    },
    "81": {
        "name": "전략기획실",
        "sub_departments": {
            "8100": {"name": "전략기획부_부장", "sub_departments": {}},
            "8101": {"name": "기획조정팀", "sub_departments": {}},
            "8103": {"name": "소통경영팀", "sub_departments": {}}
        }
    },
    "82": {
        "name": "경영지원부",
        "sub_departments": {
            "8200": {"name": "경영지원부_부장", "sub_departments": {}},
            "8201": {"name": "인사총무팀", "sub_departments": {}},
            "8202": {"name": "재정지원팀", "sub_departments": {}}
        }
    },
    "83": {
        "name": "체육사업부",
        "sub_departments": {
            "8300": {"name": "체육사업부_부장", "sub_departments": {}},
            "8301": {
                "name": "체육사업1팀",
                "sub_departments": {
                    "8352": {"name": "충무", "sub_departments": {}},
                    "8351": {"name": "무학봉", "sub_departments": {}},
                    "8358": {"name": "훈련원공원", "sub_departments": {}}
                }
            },
            "8302": {
                "name": "체육사업2팀",
                "sub_departments": {
                    "8354": {"name": "회현", "sub_departments": {}},
                    "8353": {"name": "손기정", "sub_departments": {}},
                    "8361": {"name": "남산", "sub_departments": {}},
                    "8360": {"name": "장충", "sub_departments": {}}
                }
            }
        }
    },
    "86": {
        "name": "주차사업부",
        "sub_departments": {
            "8600": {"name": "주차사업부_부장", "sub_departments": {}},
            "8601": {"name": "주차사업1팀", "sub_departments": {}},
            "8602": {"name": "주차사업2팀", "sub_departments": {}}
        }
    },
    "91": {
        "name": "시설관리부",
        "sub_departments": {
            "9100": {"name": "시설관리부_부장", "sub_departments": {}},
            "9101": {"name": "시설관리1팀", "sub_departments": {}},
            "9102": {
                "name": "시설관리2팀",
                "sub_departments": {
                    "9151": {"name": "중립종합복지센터", "sub_departments": {}},
                    "9152": {"name": "중구종합복지센터", "sub_departments": {}},
                    "9153": {"name": "신당누리센터", "sub_departments": {}},
                    "9154": {"name": "교육지원센터", "sub_departments": {}}
                }
            },
            "9103": {
                "name": "공공시설팀",
                "sub_departments": {
                    "9155": {"name": "중구구민회관", "sub_departments": {}},
                    "9156": {"name": "충무창업큐브", "sub_departments": {}},
                    "9157": {"name": "공중공원화장실", "sub_departments": {}}
                }
            }
        }
    },
    "89": {"name": "사회서비스단", "sub_departments": {}}
}

    department_counts = pd.Series({dept: department_counts[dept] for dept in departments if dept in department_counts})
    position_counts = pd.Series({pos: position_counts[pos] for pos in positions if pos in position_counts})

    create_excel_file(
        department_counts,
        position_counts,
        grouped_counts,
        total_quota,
        {},  # position_quota
        {},  # department_quota
        detailed_quota,
        sub_department_counts,
        department_hierarchy,
        input_df,
        departments,  # 사용자 지정 부서 순서
        positions     # 사용자 지정 직급 순서
    )

if __name__ == "__main__":
    main()