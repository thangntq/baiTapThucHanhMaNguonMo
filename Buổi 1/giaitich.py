import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton

import sympy as sym

class Maytinh_daoham(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('May tinh dao ham')
        self.setGeometry(100, 100, 400, 200)

        self.layout = QVBoxLayout()

        self.expression_label = QLabel('Nhap bieu thuc:')
        self.expression_input = QLineEdit()
        self.degree_label = QLabel('Nhap bac cua dao ham:')
        self.degree_input = QLineEdit()
        self.result_label = QLabel()

        self.calculate_button = QPushButton('Tinh dao ham')
        self.calculate_button.clicked.connect(self.calculate_derivative)

        self.calculate_integral_button = QPushButton('Tinh tich phan')
        self.calculate_integral_button.clicked.connect(self.calculate_integral)

        self.layout.addWidget(self.expression_label)
        self.layout.addWidget(self.expression_input)
        self.layout.addWidget(self.degree_label)
        self.layout.addWidget(self.degree_input)
        self.layout.addWidget(self.calculate_button)
        self.layout.addWidget(self.calculate_integral_button)
        self.layout.addWidget(self.result_label)

        self.setLayout(self.layout)

    def calculate_derivative(self):
        expression = self.expression_input.text()
        degree = self.degree_input.text()

        x = sym.Symbol('x')
        try:
            degree = int(degree)
            derivative = sym.diff(expression, x, degree)
            self.result_label.setText(f" Dao ham bac {degree} la : {derivative}")
        except (ValueError, sympy.SympifyError):
            self.result_label.setText("Bac khong hop le")

    def calculate_integral(self):
        expression = self.expression_input.text()
        x = sym.Symbol('x')
        try:
            integral = sym.integrate(expression, x)
            self.result_label.setText(f" Tich phan cua bieu thuc la : {integral}")
        except (ValueError, sympy.SympifyError):
            self.result_label.setText("Gia tri hoac bieu thuc khong hop le")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Maytinh_daoham()
    window.show()
    sys.exit(app.exec_())
