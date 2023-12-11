import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from customtkinter import *
from tkinter import messagebox
from PIL import Image
import os

win = CTk()
win.geometry('909x503')
win.title('bai 10')

displayScreen = CTkLabel(win,corner_radius=10, font = ('Times new roman', 13), fg_color= 'white', width = 475, height = 481)
displayScreen.place(x = 415, y = 8)

title = CTkLabel(win,corner_radius=10, text = "Tieu chi thong ke", font = ('Times new roman', 13), fg_color= '#E333bd', text_color = 'black', width = 171, height = 53)
title.place(x = 120, y = 14)

def open_file_dialog():
    file_path = filedialog.askopenfilename()
    if file_path:
        # Lấy tên tệp từ đường dẫn và hiển thị lên nhãn (Label)
        file_name = file_path.split("/")[-1]  # Lấy tên tệp từ đường dẫn
        input4.delete("1.0", "end")
        input4.insert(INSERT, file_name)

def xu_ly_du_lieu():
    file_name = input4.get("1.0", "end").strip()
    if file_name != "":
        try:
            df=pd.read_csv(file_name,index_col=0,header = 0)
            in_data = np.array(df.iloc[:,:])
            return in_data
        except FileNotFoundError:
            messagebox.showwarning("Lỗi file", "File không hợp lệ hoặc file không tồn tại.")
    else: 
        messagebox.showwarning('Cảnh báo','Chưa nhập file')

def displayPicture(fileName):
    img = CTkImage(light_image=Image.open(f'{fileName}'), size=(475, 481))
    image = CTkButton(win, corner_radius=6, text = '', image=img, width=475, height=481)
    image.place(x = 415, y = 8)

#muc 1: tong so sinh vien
def sum_sv(in_data):
    tong_sv = np.sum(in_data[:, 1])
    return tong_sv

#muc 2: tong so sinh vien dat tung loai diem
def tong_sv_type(in_data):
    Text = "Loại A,Loại B+,Loại B,Loại C+,Loại C,Loại D+,Loại D,Loại F"
    loai_diem = Text.split(",")
    tong_Ldiem = []
    for i in range(3, 11):
        tong_Ldiem.append(np.sum(in_data[:, i]))
    result = dict(zip(loai_diem, tong_Ldiem))
    data = result
    categories = list(data.keys())
    values = list(data.values())
    plt.figure(1, figsize=(8, 8))
    plt.bar(categories, values)
    plt.xlabel('Loại điểm')
    plt.ylabel('Số lượng')
    plt.title('Biểu đồ thể hiện số lượng từng loại điểm')
    plt.yticks(np.arange(0, 100, 5))
    plt.savefig('muc_2.png')
    #plt.show()
    plt.close()
    return ''
def button1():
    data = xu_ly_du_lieu()
    tong_sv_type(data)
    displayPicture('muc_2.png')
    os.remove(f'muc_2.png')


#muc 3.1: Phần trăm số SV đạt của môn học
def phan_tram_dat(in_data):
    part = ['sv đạt', 'sv trượt']
    rate = []
    sv_dat = []
    for i in range(1, 9):
        n = np.array(in_data[i, :])
        sv_dat.append(np.sum(n[3:10]))
    phan_tram_sv_dat = np.sum(sv_dat) / sum_sv(in_data)
    rate.append(phan_tram_sv_dat)
    rate.append(1 - phan_tram_sv_dat)
    colors = ['lightgreen', 'lightcoral']
    explode = (0, 0.1)  # Để phân cách một phần nhỏ (loại Đủ) ra xa     
    plt.figure(2, figsize=(8, 8))
    plt.pie(rate, explode=explode, labels=part, colors=colors, autopct='%1.1f%%', shadow=True, startangle=140)
    plt.title('Số lượng sinh viên đạt và không đạt\n')
    plt.axis('equal')  # Đảm bảo biểu đồ hình tròn hoàn chỉnh
    plt.savefig('muc_3.png')
    #plt.show()
    plt.close()
    return ''
def button2():
    data = xu_ly_du_lieu()
    phan_tram_dat(data)
    displayPicture('muc_3.png')
    os.remove(f'muc_3.png')
    
#muc 6: Tìm TBC số sv đạt điểm A,B.. của cả 9 lớp
def tbc_loai_diem(in_data):
    Text = "Loại A,Loại B+,Loại B,Loại C+,Loại C,Loại D+,Loại D,Loại F"
    loai_diem1 = Text.split(",")
    tbc = []
    for i in range(3, 11):
        tbc.append(round(np.mean((in_data[:, i])), 0))
    mydict = dict(zip(loai_diem1, tbc))
    data = mydict
    categories = list(data.keys())
    values = list(data.values())
    plt.figure(3, figsize=(8, 8))
    plt.bar(categories, values)
    plt.xlabel('Loại điểm')
    plt.ylabel('Số lượng sinh viên')
    plt.title('Biểu đồ thể hiện \ntrung bình sinh viên đạt từng loại điểm của 9 lớp')
    #plt.show()
    plt.savefig('muc_7.png')
    plt.close()
    return mydict

def button3():
    data = xu_ly_du_lieu()
    tbc_loai_diem(data)
    displayPicture('muc_7.png')
    os.remove(f'muc_7.png')

#muc 9: do thi pho diem tung lop
def draw_graph_mark(in_data):
    Ma_lop = list(in_data[:, 0])
    Pho_diem = []
    Pd_lop = {}
    Text = "Loại A,Loại B+,Loại B,Loại C+,Loại C,Loại D+,Loại D,Loại F"
    loai_diem1 = Text.split(",")
    for i in range(0, 9):
        info = list(in_data[i, :])
        type_diem = []
        for j in range(3, 11):
            type_diem.append(info[j])
        my_dict1 = dict(zip(loai_diem1, type_diem))
        Pho_diem.append(my_dict1)
    Pd_lop = dict(zip(Ma_lop, Pho_diem))
    
    data = Pd_lop
    tenlop = list(data.keys())
    Pho_diem = list(data[tenlop[0]].keys())
    x = np.arange(len(tenlop))
    bottom_values = [0] * len(tenlop)
    for loai_diem in Pho_diem:
        values = [data[lop][loai_diem] for lop in tenlop]
        plt.bar(x, values, label=loai_diem, bottom=bottom_values)
        bottom_values = [bottom + value for bottom, value in zip(bottom_values, values)]
    plt.xlabel('Lớp')
    plt.ylabel('số lượng điểm')
    plt.title('Phổ điểm của từng lớp')
    plt.xticks(x, tenlop)
    plt.yticks(np.arange(0, 65, 2))
    plt.legend()
    fig = plt.gcf()
    fig.set_size_inches(19, 8)
    plt.savefig('muc_10.png', dpi = 300) 
    #plt.show()
    plt.close()
    return ''

def button4():
    data = xu_ly_du_lieu()
    draw_graph_mark(data)
    displayPicture('muc_10.png')
    os.remove(f'muc_10.png')


CTkButton(win, corner_radius=11, text = "Nhập file", command = open_file_dialog, font = ('Times new roman', 13), fg_color= '#8CA1EE', text_color = 'Black', width = 88, height = 47, border_width = 2, border_color= '#4d8cec').place(x = 19, y = 91)
input4 = CTkTextbox(win,font = ('Times New Roman', 13), fg_color = "white", height = 46, width = 240, border_width = 2)
input4.configure(fg_color = "white")
input4.place(x = 120, y = 91)

CTkButton(win, corner_radius=45, text = '1', command = button1, font = ('Times new roman', 15), text_color = 'black', fg_color = '#3082fe', width = 89, height = 89).place(x = 70, y = 187)
CTkButton(win, corner_radius=45, text = '2', command = button2, font = ('Times new roman', 15), text_color = 'black', fg_color = '#3082fe', width = 89, height = 89).place(x = 251, y = 187)
CTkButton(win, corner_radius=45, text = '3', command = button3, font = ('Times new roman', 15), text_color = 'black', fg_color = '#3082fe', width = 89, height = 89).place(x = 70, y = 319)
CTkButton(win, corner_radius=45, text = '4', command = button4, font = ('Times new roman', 15), text_color = 'black', fg_color = '#3082fe', width = 89, height = 89).place(x = 251, y = 319)

win.mainloop()
