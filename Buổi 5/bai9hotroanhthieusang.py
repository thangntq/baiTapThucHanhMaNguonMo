import cv2
import numpy as np

# Hàm cân bằng histogram để cải thiện độ sáng tổng thể của ảnh
def enhance_brightness_contrast(image):
    # Chuyển ảnh sang ảnh grayscale nếu cần
    if len(image.shape) == 3:
        image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Cân bằng histogram
    enhanced_image = cv2.equalizeHist(image)
    
    return enhanced_image

# Hàm điều chỉnh độ sáng và độ tương phản
def adjust_brightness_contrast(image, alpha=1.5, beta=30):
    adjusted_image = cv2.convertScaleAbs(image, alpha=alpha, beta=beta)
    return adjusted_image

# Chọn ảnh từ bàn phím và tăng cường độ sáng
def enhance_image():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.png *.jpeg")])
    if file_path:
        original_image = cv2.imread(file_path)
        
        if original_image is not None:
            enhanced_image = enhance_brightness_contrast(original_image)
            adjusted_image = adjust_brightness_contrast(original_image)
            
            # Hiển thị ảnh gốc và ảnh đã tăng cường
            cv2.imshow("Original Image", original_image)
            cv2.imshow("Enhanced Image", enhanced_image)
            cv2.imshow("Adjusted Image", adjusted_image)
            
            cv2.waitKey(0)
            cv2.destroyAllWindows()
        else:
            messagebox.showerror("Error", "Không thể đọc ảnh. Vui lòng chọn một tập tin hình ảnh hợp lệ.")

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

root = tk.Tk()
root.title("Tăng cường ảnh chụp thiếu sáng")
root.geometry('450x450')
select_button = tk.Button(root, text="Chọn ảnh", command=enhance_image)
select_button.pack()

root.mainloop()
