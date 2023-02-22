import tkinter.ttk as ttk   
from tkinter import*

root = Tk()
root.title("SP IMPORT")
root.resizable(False,False)

# 이미지 넣기


#프레임1 (엑셀파일선택, 경로)
xls_frame = LabelFrame(root, text="Select xls file")
xls_frame.pack(fill="x", padx=5, pady=5, ipady=5)

txt_xls_path = Entry(xls_frame)
txt_xls_path.pack(side="left", fill="x", expand="True", ipady=4) #

btn_xls_path = Button(xls_frame, text="...xls file",padx=5, pady=5, width=12)
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

btn_savepdf_path = Button(savepdf_frame, text="...Folder to save",padx=5, pady=5, width=12)
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

btn_mergedpdf_path = Button(mergedpdf_frame, text="...Folder to save",padx=5, pady=5, width=12)
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

