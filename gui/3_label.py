from tkinter import *

root = Tk()
root.title("SP IMPORT")
root.geometry("640x480")  #가로 * 세로
root.resizable(False, False)  # x(너비) y(높이) 값 변경불가 창크기 고정  

label1 = Label(root,text="안녕하세요")
label1.pack()

photo=PhotoImage(file="gui/imgtest.png")
label2 = Label(root, image=photo)
label2.pack()

def change():
    label1.config(text="또 만나요")

    global photo2
    photo2 = PhotoImage(file="gui/imgtest2.png")
    label2.config(image=photo2)


btn = Button(root, text="클릭", command=change)
btn.pack()



root.mainloop()