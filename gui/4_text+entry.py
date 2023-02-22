from tkinter import *

root = Tk()
root.title("SP IMPORT")
root.geometry("640x480")  #가로 * 세로
root.resizable(False, False)  # x(너비) y(높이) 값 변경불가 창크기 고정  

txt = Text(root,width=30, height=5)
txt.pack()

txt.insert(END,"글자를 입력하세요")


e = Entry(root, width=30)  #엔터 입력불가, 한줄로 뭔가 입력받을 때 아이디, 이름 / 텍스트트 여러줄 가능
e.pack()
e.insert(0,"한 줄만 입력해요")


def btncmd():
    # 내용 출력
    print(txt.get("1.0", END)) #1은 첫번째 라인 0은 첫번째 컬럼 
    print(e.get())

    #내용 삭제
    txt.delete("1.0", END)
    e.delete(0, END)

btn = Button(root, text="클릭", command=btncmd)
btn.pack()



root.mainloop()