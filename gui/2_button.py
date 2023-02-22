from tkinter import *

root = Tk()
root.title("SP IMPORT")
root.geometry("640x480")  #가로 * 세로
root.resizable(False, False)  # x(너비) y(높이) 값 변경불가 창크기 고정  

btn1 = Button(root, text="버튼1")
btn1.pack()

btn2 = Button(root, padx=5, pady=10, text="버튼2222222")
btn2.pack()

btn3 = Button(root, padx=5, pady=10, text="버튼2")
btn3.pack()

btn4 = Button(root, width=10, height=3, text="버튼4")  # 버튼 크기가 고정된 크기로 출력(내용이 길면 짤림)
btn4.pack()

btn5 = Button(root, fg="red", bg="yellow", text='버튼5')
btn5.pack()

photo = PhotoImage(file="gui/imgtest.png") #경로를 로컬로 하는게 나은듯?
btn6 = Button(root, image=photo)
btn6.pack()


def btncmd():
    print("버튼이 클릭되었어요")

btn7 = Button(root, text="동작하는 버튼", command=btncmd)
btn7.pack()



root.mainloop()
