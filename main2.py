import sys
import xlwings as xw
import pandas as pd
from PySide6.QtWidgets import (QApplication, QTableWidget, QTableWidgetItem, 
                               QVBoxLayout, QHBoxLayout, QWidget, QMenu, 
                               QHeaderView, QPushButton, QFileDialog, QScrollArea)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPalette

class ExcelTableWidget(QTableWidget):
    def __init__(self, rows=5, cols=6):
        super().__init__()
        self.setRowCount(rows)
        self.setColumnCount(cols)
        self.setup_headers()
        self.customize_header_style()
        self.center_align_all_cells()
        
        # Enable table to stretch
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def setup_headers(self):
        horizontal_header = self.horizontalHeader()
        vertical_header = self.verticalHeader()
        
        horizontal_header.setContextMenuPolicy(Qt.CustomContextMenu)
        vertical_header.setContextMenuPolicy(Qt.CustomContextMenu)
        
        horizontal_header.customContextMenuRequested.connect(self.show_column_menu)
        vertical_header.customContextMenuRequested.connect(self.show_row_menu)

    def show_column_menu(self, pos):
        column = self.horizontalHeader().logicalIndexAt(pos)
        menu = QMenu()
        delete_action = menu.addAction("Delete Column")
        action = menu.exec(self.horizontalHeader().mapToGlobal(pos))
        
        if action == delete_action:
            self.removeColumn(column)
            # Ensure headers remain stretched
            self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            self.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def show_row_menu(self, pos):
        row = self.verticalHeader().logicalIndexAt(pos)
        menu = QMenu()
        delete_action = menu.addAction("Delete Row")
        action = menu.exec(self.verticalHeader().mapToGlobal(pos))
        
        if action == delete_action:
            self.removeRow(row)
            # Ensure headers remain stretched
            self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            self.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def load_excel_data(self, file_path):
        # Open the Excel workbook
        wb = xw.Book(file_path)
        sheet = wb.sheets[0]

        # Read data into pandas DataFrame
        df = sheet.range('A1:U19').options(pd.DataFrame, header=1, index=False).value
        
        # Sort DataFrame by 'Subject' column
        df_sorted = df.sort_values(by='번호')

        # Prepare data for TableWidget (include headers)
        data = [df_sorted.columns.tolist()] + df_sorted.values.tolist()

        # Clear existing content
        self.clearContents()
        self.setRowCount(len(data)-1)
        self.setColumnCount(len(data[0]))

        # Populate table
        for row_idx, row in enumerate(data[1:], start=0):
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignCenter)
                self.setItem(row_idx, col_idx, item)
        
        # Set headers
        self.setHorizontalHeaderLabels(data[0])
        
        # Re-enable stretching
        self.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

    def customize_header_style(self):
        h_header = self.horizontalHeader()
        h_header.setStyleSheet("""
            QHeaderView::section {
                background-color: #3498db;
                color: white;
                padding: 4px;
                border: 1px solid #2980b9;
                font-weight: bold;
                text-align: center;
            }
        """)

        v_header = self.verticalHeader()
        v_header.setStyleSheet("""
            QHeaderView::section {
                background-color: #2ecc71;
                color: white;
                padding: 4px;
                border: 1px solid #27ae60;
                font-weight: bold;
                text-align: center;
            }
        """)

    def center_align_all_cells(self):
        h_header = self.horizontalHeader()
        v_header = self.verticalHeader()
        h_header.setDefaultAlignment(Qt.AlignCenter)
        v_header.setDefaultAlignment(Qt.AlignCenter)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # Main layout
        layout = QVBoxLayout()

        # Button layout
        button_layout = QHBoxLayout()
        self.file_select_btn = QPushButton("Excel 파일 선택")
        self.file_select_btn.clicked.connect(self.select_excel_file)
        button_layout.addWidget(self.file_select_btn)

        # Scroll Area
        scroll_area = QScrollArea()
        self.table = ExcelTableWidget()
        scroll_area.setWidget(self.table)
        scroll_area.setWidgetResizable(True)  # Important: allows the widget to resize
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

        # Add layouts
        layout.addLayout(button_layout)
        layout.addWidget(scroll_area)

        self.setLayout(layout)
        self.setWindowTitle('Excel Viewer')
        self.resize(800, 600)

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Excel 파일 선택", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                self.table.load_excel_data(file_path)
            except Exception as e:
                print(f"파일 로드 중 오류 발생: {e}")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()