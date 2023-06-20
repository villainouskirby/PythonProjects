from tkinter import *
import tkinter.ttk as ttk
import tkinter.messagebox as msgbox
from tkinter import filedialog
import os
import sys
import random
import win32com.client as win32

title_path = sys.argv[0]
title = os.path.splitext(os.path.basename(title_path))

root = Tk()
root.title(title[0])

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def switch_frame(frame):
    global c_frame
    c_frame.pack_forget()
    c_frame = frame
    frame.pack(fill="both", expand=True)

def save_names():
    names = open("name.txt", "w", encoding="utf-8")
    names.write(name_txt.get("1.0", END))
    names.close()

def get_names():
    global all_lst
    all_lst = list(map(str, name_txt.get("1.0", END).replace(" ", "").strip().split(",")))
    switch_frame(page2)

def set_name():
    name_lst.clear()
    idx = 0
    btns = seats_frame.grid_slaves()[::-1]
    seats_num = len(btns)
    for i in all_lst:
        while idx < seats_num:
            btn = btns[idx]
            btn.config(text="")
            idx += 1
            if btn['state'] == NORMAL:
                btn.config(text=i)
                name_lst.append(i)
                break
    while idx < len(btns):
        btn = btns[idx]
        btn.config(text="")
        idx += 1

def change(btn):
    global c_btn, c_text, c_bool
    if c_bool:
        c_btn.config(text=btn['text'], bg="#F0F0F0")
        btn.config(text=c_text)
        c_bool = False
    else:
        btn.config(bg="#81F79F")
        c_btn = btn
        c_text = btn['text']
        c_bool = True
    

def disable(btn):
    if btn['state'] == NORMAL:
        btn.config(state=DISABLED, bg="#000000")
    else:
        btn.config(state=NORMAL, bg="#F0F0F0")
    set_name()

def seats():
    try:
        a = int(row.get())
        b = int(column.get())
    except ValueError:
        a=b=0
    if a < 1 or b < 1:
        msgbox.showerror("에러", "1 이상의 숫자를 입력해 주십시오.")
        return
    if a*b > 1000:
        msgbox.showerror("에러", "숫자가 너무 큽니다.")
        return

    board.pack(pady=5)

    for i in seats_frame.grid_slaves():
        i.config(state=NORMAL, bg='#F0F0F0', text='')
        i.grid_forget()
    for i in range(a):
        for j in range(b):
            exec(f"b{i*b+j}.grid(row={j}, column={i})")

    confirm_btn.pack(pady=5)
    data_frame.pack(fill="x", expand=True, padx=5, pady=5)
    set_name()

def shuffle():
    random.shuffle(name_lst)
    idx = 0
    btns = seats_frame.grid_slaves()[::-1]
    for i in name_lst:
        while idx < len(btns):
            btn = btns[idx]
            btn.config(text="")
            idx += 1
            if btn['state'] == NORMAL:
                btn.config(text=i)
                break

def open_file():
    file_selected = filedialog.askopenfilename(title="열기",filetypes=(("hwp 파일", "*.hwp"), ("모든 파일", "*.*")))
    if file_selected == data_entry.get() or file_selected == '':
        return

    data_entry.config(state="normal")
    data_entry.delete(0, END)
    data_entry.insert(0, file_selected)
    data_entry.config(state="readonly")

def set_datas():
    dest = data_entry.get()
    if dest == '':
        msgbox.showerror("에러", "선택된 파일이 없습니다.")
        return
    hwp=win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.Open(dest,"HWP","forceopen:true")
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    option=hwp.HParameterSet.HFindReplace
    l = len(name_lst)
    for i, j in enumerate(name_lst[::-1]):
        option.FindString = "pos"+str(l - i)
        option.ReplaceString = j
        option.IgnoreMessage = 1
        hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.SaveAs(saved_dest)
    hwp.Clear(3)
    hwp.Quit()

    del_print.pack(side="right", padx=5, pady=5)
    check_not_print.pack(side="right", padx=5, pady=5)

    msgbox.showinfo("알림", "변환이 완료되었습니다.")

def printer():
    os.startfile(saved_dest, "print")

def file_config():
    os.startfile(saved_dest)

# 데이터 관리
page1 = Frame(root)
page2 = Frame(root)
page3 = Frame(root)
name_path = resource_path("name.txt")
names = open(name_path, "r", encoding="utf-8").read().strip()
all_lst = []
name_lst = []
c_frame = page1
c_btn = 0
c_text = ""
c_bool = False
saved_dest = os.path.dirname(__file__) + f"\\{title[0]}.hwp"
switch_frame(page1)

# page1

# 이름 칸
name_frame = ttk.Labelframe(page1, text="이름")
name_frame.pack(side="top", fill="x", expand=True, padx=5, pady=5)

name_lst_frame = Frame(name_frame)
name_lst_frame.pack(side="top", fill="x", expand=True, padx=5, pady=5)

names_scrollbar = Scrollbar(name_lst_frame)
names_scrollbar.pack(side="right", fill="y")

name_txt = Text(name_lst_frame, width=48, height=5, yscrollcommand=names_scrollbar.set)
name_txt.insert(END, names)
name_txt.pack(side="top", fill="x", expand=True)

names_scrollbar.config(command=name_txt.yview)

save_frame = Frame(page1)
save_frame.pack(fill="x")
save_btn = Button(save_frame, text="저장", padx=20, pady=5, command=save_names)
save_btn.pack(side="left", padx=5, pady=5)

next_btn = Button(save_frame, text="다음", padx=20, pady=5, command=get_names)
next_btn.pack(side="right", padx=5, pady=5)

# page 2

# 자리 수 정하기
seat_num_frame = Frame(page2)
seat_num_frame.pack(side="top", padx=5)

row_txt = Label(seat_num_frame, text="가로 : ")
row_txt.pack(side="left", pady=5)

row = Entry(seat_num_frame, width=5)
row.pack(side="left", pady=5)

column = Label(seat_num_frame, text="   세로 : ")
column.pack(side="left", pady=5)

column = Entry(seat_num_frame, width=5)
column.pack(side="left", pady=5)

mid_frame = Frame(page2)
mid_frame.pack()

seat_btn = Button(mid_frame, text="생성하기", padx=20, pady=5, command=seats)
seat_btn.pack(padx=5, pady=5)

board = Label(mid_frame, text="교탁")

seats_frame = Frame(page2)
seats_frame.pack(expand=True, pady=5)

for i in range(1000):
    exec(f"b{i} = Button(seats_frame, width=10, height=2, command=lambda:change(b{i}))")
    exec(f"b{i}.bind('<Button-3>', lambda event: disable(b{i}))")

confirm_frame = Frame(page2)
confirm_frame.pack(fill="x", expand=True)

confirm_btn = Button(confirm_frame, padx=20, pady=5, text="섞기", command=shuffle)

data_frame = ttk.Labelframe(confirm_frame, text="파일")

data_entry = Entry(data_frame, state="readonly")
data_entry.pack(side="left", fill="x", expand=True, ipady=4, padx=5, pady=5)

set_btn = Button(data_frame, text="변환하기", padx=10, command=set_datas)
set_btn.pack(side="right", padx=5, pady=5)

find_btn = Button(data_frame, text="찾기", padx=10, command=open_file)
find_btn.pack(side="right", padx=5, pady=5)

end_frame = Frame(page2)
end_frame.pack(fill="x", expand=True)

prev_btn = Button(end_frame, text="이전", padx=20, pady=12, command=lambda:switch_frame(page1))
prev_btn.pack(side="left", padx=5, pady=5)

del_print = Button(end_frame, text="바로\n출력하기", padx=20, pady=5, command=printer)

check_not_print = Button(end_frame, text="파일\n확인하기", padx=20, pady=5,  command=file_config)

##########################################################################

path = os.path.join(os.path.dirname(__file__),'kirby.ico')
if os.path.isfile(path):
    root.iconbitmap(path)

root.resizable(False, False)
root.mainloop()