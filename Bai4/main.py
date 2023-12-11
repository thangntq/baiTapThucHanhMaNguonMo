from customtkinter import *
import thong_ke as tk
import pandas as pd
from numpy import array
from tkinter import *
from PIL import Image, ImageTk
from tkinter import filedialog, messagebox

win = CTk()
win.title('Make report 2.0')
win.geometry('929x841')
win.resizable(False, False)
win.configure(bg = '#FAFAFA')

def xu_ly_du_lieu():
    file_name = input4.get("1.0", "end").strip()
    if file_name != "":
        try:
            df=pd.read_csv(file_name,index_col=0,header = 0)
            in_data = array(df.iloc[:,:])
            return in_data
        except FileNotFoundError:
            messagebox.showwarning("Lỗi file", "File không hợp lệ hoặc file không tồn tại.")
    else: 
        messagebox.showwarning('Cảnh báo','Chưa nhập file')
#title
CTkLabel(win, text = "Phần mềm xuất báo cáo", font = ('Times new roman', 14), fg_color= '#BA37E8', text_color = 'Black', width = 280, height = 72).place(x = 324, y = 18)

display_data = CTkTextbox(win,font = ('Times New Roman', 13), fg_color = "white", height = 330, width = 847, border_width = 2)
display_data.configure(fg_color = "white")
display_data.place(x = 41, y = 476)

CTkButton(win, text = 'Hiển thị dữ liệu', font = ('Times new roman', 15), text_color = 'black', fg_color = '#EFEA75', width = 137, height = 46, border_width = 2, border_color= '#000').place(x = 41, y = 420)

#tên giảng viên
CTkLabel(win, text = "Tên giảng viên", font = ('Times new roman', 13), fg_color= '#8CA1EE', text_color = 'Black', width = 137, height = 46).place(x = 41, y = 118)
input1 = CTkTextbox(win,font = ('Times New Roman', 13), fg_color = "white", height = 46, width = 240, border_width = 2)
input1.configure(fg_color = "white")
input1.place(x = 209, y = 118)

#tên học phần
CTkLabel(win, text = "Tên học phần", font = ('Times new roman', 13), fg_color= '#8CA1EE', text_color = 'Black', width = 137, height = 46).place(x = 41, y = 192)
input2 = CTkTextbox(win,font = ('Times New Roman', 13), fg_color = "white", height = 46, width = 240, border_width = 2)
input2.configure(fg_color = "white")
input2.place(x = 209, y = 192)

#nhập tên khoa
CTkLabel(win, text = "Tên Khoa", font = ('Times new roman', 13), fg_color= '#8CA1EE', text_color = 'Black', width = 137, height = 46).place(x = 41, y = 259)
input3 = CTkTextbox(win,font = ('Times New Roman', 13), fg_color = "white", height = 46, width = 240, border_width = 2)
input3.configure(fg_color = "white")
input3.place(x = 209, y = 259)

#Nhập file điểm
def open_file_dialog():
    file_path = filedialog.askopenfilename()
    if file_path:
        # Lấy tên tệp từ đường dẫn và hiển thị lên nhãn (Label)
        file_name = file_path.split("/")[-1]  # Lấy tên tệp từ đường dẫn
        input4.delete("1.0", "end")
        input4.insert(INSERT, file_name)
CTkButton(win, text = "Nhập file", command = open_file_dialog, font = ('Times new roman', 13), fg_color= '#8CA1EE', text_color = 'Black', width = 137, height = 46, border_width = 2, border_color= '#000').place(x = 41, y = 333)
input4 = CTkTextbox(win,font = ('Times New Roman', 13), fg_color = "white", height = 46, width = 240, border_width = 2)
input4.configure(fg_color = "white")
input4.place(x = 209, y = 333)

CTkLabel(win, text = "Định dạng font", font = ('Times new roman', 13), fg_color= '#8CA1EE', text_color = 'Black', width = 137, height = 46).place(x = 480, y = 118)
font = StringVar(value='Times New Roman')
CTkOptionMenu(win,variable = font, width = 240, height = 46, button_color='#572AD6', button_hover_color='white',  hover = True, values = ['Times New Roman', 'Arial']).place(x = 648, y = 118)

CTkLabel(win, text = "Kích thước chữ", font = ('Times new roman', 13), fg_color= '#8CA1EE', text_color = 'Black', width = 137, height = 46).place(x = 480, y = 192)
size = StringVar(value='14')
CTkOptionMenu(win, variable=size, width = 240, height = 46, button_color='#572AD6', button_hover_color='white',  hover = True, values = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24']).place(x = 648, y = 192)

CTkLabel(win, text = "Định dạng file", font = ('Times new roman', 13), fg_color= '#8CA1EE', text_color = 'Black', width = 137, height = 46).place(x = 480, y = 259)
type = StringVar(value='Word')
CTkOptionMenu(win, variable = type, width = 240, height = 46, button_color='#572AD6', button_hover_color='white',  hover = True, values = ['Word', 'PDF']).place(x = 648, y = 259)

#nhập tên file muốn lưu
CTkLabel(win, text = "Tên file lưu", font = ('Times new roman', 13), fg_color= '#8CA1EE', text_color = 'Black', width = 137, height = 46).place(x = 480, y = 333)
input6 = CTkTextbox(win,font = ('Times New Roman', 13), fg_color = "white", height = 46, width = 240, border_width = 2)
input6.configure(fg_color = "white")
input6.place(x = 648, y = 333)

#lưu ý
""" note = '*Lưu ý: \n- Nhập tên giảng viên phải nhập cả học hàm học vị \n- File điểm được nhập là file dưới định dạng csv \n- Các mục của file bao gồm STT, tên lớp, Sĩ số, loại A+ => loại F, chuẩn L1, l2, Cuối kỳ'
input5 = Text(win,font = ('Times New Roman', 13), height = 5, width = 50, relief='solid', borderwidth = 2)
input5.configure(fg = 'black', bg = "white")
input5.place(x = 270, y = 370)
input5.insert(INSERT, note) """

def export_report():
    teacher_name = input1.get("1.0", "end").strip()
    subject_name = input2.get("1.0", "end").strip()
    faculty_name = input3.get("1.0", "end").strip()
    filesave = input6.get("1.0", "end").strip()
    if (teacher_name != "" or subject_name != "" or faculty_name != "" or filesave != ""):
        data = xu_ly_du_lieu()
        tk.xuat_report(data, teacher_name, subject_name, faculty_name, filesave, font.get(), int(size.get()), type.get())
    else:
        messagebox.showwarning('Cảnh báo','Hãy nhập đầy đủ thông tin')
CTkButton(win, text = 'Xuất report', command = export_report, font = ('Times new roman', 15), text_color = 'black', fg_color = '#EFEA75', width = 137, height = 46, border_width = 2, border_color= '#000').place(x = 312, y = 420)
CTkButton(win, text = 'Thoat', command = exit, font = ('Times new roman', 15), text_color = 'black', fg_color = '#EFEA75', width = 137, height = 46, border_width = 2, border_color= '#000').place(x = 480, y = 420)

win.mainloop()