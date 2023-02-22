import tkinter.ttk as ttk   
from tkinter import *  #__all__? filedialog는 서브 모듈이기 때문에 별도 임포트 해줘야 됨 
from tkinter import filedialog


root = Tk()
root.title("SP IMPORT")
root.resizable(False,False)

# 이미지 넣기


#파일추가
def add_file():
    files = filedialog.askopenfilename(title="xlsx 파일을 선택하세요", \
        filetypes=(("xlsx 파일", "*.xlsx"),("All files", "*.*")), \
        initialdir="C:/")
    #사용자가 선택한 파일출력
    for file in files:
        txt_xls_path.insert(END, file)

#PDF저장경로
def savepdf_path():
    savepdf_selected = filedialog.askdirectory()
    if savepdf_selected == '': #사용자가 취소를 누를 때
        return
    txt_savepdf_path.delete(0, END)
    txt_savepdf_path.insert(0,savepdf_selected)

#MergedPDF저장경로
def savemerged_path():
    savemerged_selected = filedialog.askdirectory()
    if savemerged_selected == '': #사용자가 취소를 누를 때
        return
    txt_mergedpdf_path.delete(0, END)
    txt_mergedpdf_path.insert(0,savemerged_selected)




#프레임1 (엑셀파일선택, 경로)
xls_frame = LabelFrame(root, text="Select xlsx file")
xls_frame.pack(fill="x", padx=5, pady=5, ipady=5)

txt_xls_path = Entry(xls_frame)
txt_xls_path.pack(side="left", fill="x", expand="True", ipady=4) #

btn_xls_path = Button(xls_frame, text="...xlsx file",padx=5, pady=5, width=12, command = add_file)
btn_xls_path.pack(side="right")



# #옵션 프레임
# frame_option = LabelFrame(root, text="Option")
# frame_option.pack()

# #SCALING 
# lbl_scaling = Label(frame_option, text="SCALING", width=3)
# lbl_scaling.pack(side="left")
# #SCALING콤보
# opt_scaling = ["","",""]
# cmb_scaling = ttk.Combobox(frame_option,state="readonly", values=opt_scaling)
#timesleep1 
#timesleep2 
#timesleep3


#Web to PDF 저장
savepdf_frame = LabelFrame(root, text="Web to PDF")
savepdf_frame.pack(fill="x", padx=5, pady=5, ipady=5)

txt_savepdf_path = Entry(savepdf_frame)
txt_savepdf_path.pack(side="left", fill="x", expand="True", ipady=4) #

btn_start = Button(savepdf_frame, text="START",padx=5, pady=5, width=12)
btn_start.pack(side="right")

btn_savepdf_path = Button(savepdf_frame, text="...Folder to save",padx=5, pady=5, width=12, command=savepdf_path)
btn_savepdf_path.pack(side="right")


#Slit PDF
splitpdf_frame = LabelFrame(root, text="Split PDF")
splitpdf_frame.pack(fill="x",padx=5, pady=5, ipady=5)

btn_splitpdf = Button(splitpdf_frame, text="START", padx=5, pady=5, width="20")
btn_splitpdf.pack()



# Merged PDF 저장경로
mergedpdf_frame = LabelFrame(root, text="Folder to merfed PDF")
mergedpdf_frame.pack(fill="x", padx=5, pady=5, ipady=5)

txt_mergedpdf_path = Entry(mergedpdf_frame)
txt_mergedpdf_path.pack(side="left", fill="x", expand="True", ipady=4) #

btn_mergestart = Button(mergedpdf_frame, text="START",padx=5, pady=5, width=12)
btn_mergestart.pack(side="right")

btn_mergedpdf_path = Button(mergedpdf_frame, text="...Folder to save",padx=5, pady=5, width=12, command=savemerged_path)
btn_mergedpdf_path.pack(side="right")





# 진행상황 profress bar
frame_progress = LabelFrame(root, text="진행상황")
frame_progress.pack(fill="x")



# 종료 프레임
close_frame = Frame(root)
close_frame.pack(side="right", padx=5, pady=5)

btn_close = Button(close_frame, padx=5, pady=5, text="CLOSE", width=12)
btn_close.pack()


p_var = DoubleVar()
progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=p_var)
progress_bar.pack(fill="x")





root.mainloop()

