import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os
import matplotlib.pyplot as plt
from matplotlib.widgets import Cursor

root = tk.Tk()
root.title("엑셀 저장 테스트")
root.geometry("500x400")
root.config(bg="#f0f0f0")

def show_message():
    messagebox.showinfo("안내", "버튼이 클릭되었습니다!")

def update_label():
    if check_var.get():
        label_check.config(text="체크박스가 선택됨")
    else:
        label_check.config(text="체크박스를 선택하세요")

def save_to_excel():
    file_name = "user_data.xlsx"
    
    if os.path.exists(file_name):
        wb = load_workbook(file_name)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "User Data"
        ws.append(["이름", "슬라이더 값", "체크박스 상태"])

    name = entry.get()
    slider_value = slider.get()
    check_status = "선택됨" if check_var.get() else "선택 안 됨"

    ws.append([name, slider_value, check_status])
    wb.save(file_name)
    messagebox.showinfo("저장 완료", f"{file_name}에 데이터가 저장되었습니다!")
    plot_graph() 

def plot_graph():
    file_name = "user_data.xlsx"
    
    if not os.path.exists(file_name):
        messagebox.showerror("오류", f"{file_name} 파일이 없습니다!")
        return

    wb = load_workbook(file_name)
    ws = wb.active
    
    checked_count = 0
    unchecked_count = 0
    names_checked = []
    names_unchecked = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        name, slider_value, check_status = row
        if check_status == "선택됨":
            checked_count += 1
            names_checked.append(name)
        else:
            unchecked_count += 1
            names_unchecked.append(name)

    labels = ['선택됨', '선택 안 됨']
    values = [checked_count, unchecked_count]

    plt.rcParams['font.family'] = 'Malgun Gothic'
    
    fig, ax = plt.subplots(figsize=(6, 6))
    bars = ax.bar(labels, values, color=['green', 'red'])
    
    def on_hover(event):
        for bar, names, label in zip(bars, [names_checked, names_unchecked], labels):
            contains, _ = bar.contains(event) 
            if contains:
                bar.set_facecolor('orange') 
                if label == '선택됨':
                    plt.title(f'{label} - {", ".join(names_checked)}')
                else:
                    plt.title(f'{label} - {", ".join(names_unchecked)}') 
            else:
                if label == '선택됨':
                    bar.set_facecolor('green')
                else:
                    bar.set_facecolor('red')
                plt.title('체크박스 상태에 따른 인원 수') 
            fig.canvas.draw_idle()



    fig.canvas.mpl_connect("motion_notify_event", on_hover)
    
    ax.set_xlabel('체크박스 상태')
    ax.set_ylabel('인원 수')
    plt.show()

label = tk.Label(root, text="안녕하세요!", font=("Helvetica", 16, "bold"), fg="blue", bg="#f0f0f0")
label.pack(pady=20)

entry_label = tk.Label(root, text="이름을 입력하세요:", font=("Arial", 12), bg="#f0f0f0")
entry_label.pack(pady=10)
entry = tk.Entry(root, font=("Arial", 12), width=25)
entry.pack(pady=5)

slider_label = tk.Label(root, text="슬라이더 값:", font=("Arial", 12), bg="#f0f0f0")
slider_label.pack(pady=10)
slider = tk.Scale(root, from_=0, to=100, orient="horizontal", font=("Arial", 12))
slider.pack(pady=10)

check_var = tk.BooleanVar()
checkbox = tk.Checkbutton(root, text="체크박스를 선택하세요", font=("Arial", 12), variable=check_var, command=update_label)
checkbox.pack(pady=10)

label_check = tk.Label(root, text="체크박스를 선택하세요", font=("Arial", 12), bg="#f0f0f0")
label_check.pack(pady=5)

save_button = tk.Button(root, text="엑셀로 저장", font=("Arial", 12), bg="#FF5733", fg="white", relief="raised", command=save_to_excel)
save_button.pack(pady=20)

root.mainloop()
