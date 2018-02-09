import os
import xlrd
import sys
import xlsxwriter
from PyQt4 import QtGui, QtCore
from PyQt4.QtGui import *


class Form(QDialog):
    def __init__(self, parent=None):
        super(Form, self).__init__(parent)

        layout = QtGui.QGridLayout()
        layout.setSpacing(1)
        self.setLayout(layout)

        self.lbl1 = QLabel("1. Select the folder containing the to-be-archived files.")
        layout.addWidget(self.lbl1, 1, 1, 1, 3)
        self.btn1 = QPushButton("Select Folder")
        self.btn1.clicked.connect(self.setSource)
        layout.addWidget(self.btn1, 1, 4, 1, 2)

        self.fileBox = QLineEdit(self)
        self.fileBox.setReadOnly(True)
        layout.addWidget(self.fileBox, 2, 1, 1, 5)

        self.lbl2 = QLabel("2. Select the folder where you want the new excel file to go.")
        layout.addWidget(self.lbl2, 3, 1, 1, 3)
        self.btn2 = QPushButton("Select Folder")
        self.btn2.clicked.connect(self.setOutput)
        layout.addWidget(self.btn2, 3, 4, 1, 2)

        self.output = QLineEdit(self)
        self.output.setReadOnly(True)
        layout.addWidget(self.output, 4, 1, 1, 5)

        self.lbl4 = QLabel("Name your new template: ")
        layout.addWidget(self.lbl4, 5, 1, 1, 2)
        self.name = QLineEdit(self)
        layout.addWidget(self.name, 5, 3, 1, 2)
        self.lbl5 = QLabel(".xlsx")
        layout.addWidget(self.lbl5, 5, 5, 1, 1)

        self.btn3 = QPushButton("Create template")
        self.btn3.clicked.connect(self.create_template)
        layout.addWidget(self.btn3, 6, 2, 1, 3)

        self.setWindowTitle("Metadata Excel template starter")
        self.setGeometry(50, 50, 500, 220)

    def get_output(self):
        self.xlsFile = QFileDialog.getOpenFileName(
            self, 'Open Excel sheet', '', 'Excel files (*.xls *.xlsx)'
        )
        self.output.setText(self.xlsFile)
        self.set_sheets(self.xlsFile)

    def get_folder(self):
        fileDir = QFileDialog.getExistingDirectory()
        return fileDir

    def setSource(self):
        s = self.get_folder()
        self.fileBox.setText(s)

    def setOutput(self):
        x = self.get_folder()
        self.output.setText(x)

    def set_sheets(self, wkbk):
        self.workbook = xlrd.open_workbook(wkbk)
        self.sheetCmb.clear()
        self.sheetCmb.addItems(self.workbook.sheet_names())

    def create_template(self):
        try:

            wsName = self.name.text()+".xlsx"

            wb = xlsxwriter.Workbook(self.output.text() + "/" + wsName)
            ws = wb.add_worksheet('DC metadata')

            fileList = os.listdir(self.fileBox.text())

            ws.write(0, 0, "File name")
            ws.write(0, 1, "Title")
            ws.write(0, 2, "Creator")
            ws.write(0, 3, "Subject")
            ws.write(0, 4, "Description")
            ws.write(0, 5, "Publisher")
            ws.write(0, 6, "Contributor")
            ws.write(0, 7, "Date")
            ws.write(0, 8, "Type")
            ws.write(0, 9, "Format")
            ws.write(0, 10, "Identifier")
            ws.write(0, 11, "Source")
            ws.write(0, 12, "Language")
            ws.write(0, 13, "Relation")
            ws.write(0, 14, "Coverage")
            ws.write(0, 15, "Rights")

            row = 1

            for file in fileList:
                ws.write(row, 0, file)
                row = row+1

            wb.close()

            doneBox = QMessageBox()
            doneBox.setIcon(QMessageBox.Information)
            doneBox.setText("Template file " + wsName + " created.")
            done = doneBox.exec()

        except:
            failBox = QMessageBox()
            failBox.setIcon(QMessageBox.Error)
            failBox.setText("Something went wrong")

def main():
    app = QApplication(sys.argv)
    ex = Form()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
