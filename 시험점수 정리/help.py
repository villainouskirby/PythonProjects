from tkinter import *

h = Toplevel()
h.title("도움말")

def h_next():
    h_update()

def h_prev():
    h_update()

def h_update():
    pass

img_idx = 0

img_lst = []

img1 = PhotoImage(file="project_file/1.png")
img2 = PhotoImage(file="project_file/2.png")
img3 = PhotoImage(file="project_file/3.png")
img4 = PhotoImage(file="project_file/4.png")

img_lst.append(img1)
img_lst.append(img2)
img_lst.append(img3)
img_lst.append(img4)

txt_lst = ["1. 찾아보기 버튼을 누르고 사용할 엑셀파일을 찾으십시오.", "2. 사용할 sheet의 이름을 선택하십시오.", "3. 넣을 값에 해당하는 열 제목을 클릭하십시오.\n누른 순서대로 출력이 되고, 잘못 선택하였을 시 지울수 있습니다.\n값이 비어있었다면 \'Unnamed: 0\' 꼴로 값이 들어가는데,\n이는 그대로 출력 되므로 본래의 파일에서 임의로 수정해야 합니다."]

img_frame = Frame(h)
img_frame.pack(fill="x")

img_label = Label(img_frame, image=img3, padx=10, pady=10)
img_label.pack()

txt_label = Label(img_frame, text=txt_lst[2])
txt_label.pack()

h_btn_frame = Frame(h)
h_btn_frame.pack(fill="x")

n_btn = Button(h_btn_frame, command=h_next)
n_btn.pack(side="left")

p_btn = Button(h_btn_frame, command=h_prev)
p_btn.pack(side="right")

h.mainloop()