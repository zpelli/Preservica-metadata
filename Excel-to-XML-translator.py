import xlrd
import sys
from PyQt4 import QtGui, QtCore
from PyQt4.QtGui import *


class Form(QDialog):
    def __init__(self, parent=None):
        super(Form, self).__init__(parent)

        layout = QtGui.QGridLayout()
        layout.setSpacing(1)
        self.setLayout(layout)

        self.l1 = QLabel("1. Select the excel sheet containing metadata.")
        layout.addWidget(self.l1, 1, 1, 1, 3)

        self.xlsFileBox = QLineEdit(self)
        self.xlsFileBox.setReadOnly(True)
        layout.addWidget(self.xlsFileBox, 2, 1, 1, 4)

        self.b1 = QPushButton("Select File")
        self.b1.clicked.connect(self.select_file)
        layout.addWidget(self.b1, 1, 4, 1, 1)

        self.l2 = QLabel("2. Select sheet: ")
        #self.l2.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        layout.addWidget(self.l2, 4, 1, 1, 2)

        self.sheetCmb = QtGui.QComboBox(self)
        layout.addWidget(self.sheetCmb, 4, 3, 1, 2)

        self.l3 = QLabel("3. Select the output directory for the metadata files. *This should be where the to-be-archived files are located*")
        #self.l3.setAlignment(QtCore.Qt.AlignVCenter)
        layout.addWidget(self.l3, 5, 1, 1, 3)

        self.output = QLineEdit(self)
        self.output.setReadOnly(True)
        layout.addWidget(self.output, 6, 1, 1, 4)

        self.b2 = QPushButton("Select folder")
        self.b2.clicked.connect(self.get_output)
        layout.addWidget(self.b2, 5, 4, 1, 1)

        self.b3 = QPushButton("Create .metadata files")
        self.b3.clicked.connect(self.create_metadata)
        layout.addWidget(self.b3, 7, 2, 1, 2)

        self.setWindowTitle("Metadata converter")
        self.setGeometry(50, 50, 500, 250)

    def select_file(self):
        self.xlsFile = QFileDialog.getOpenFileName(
            self, 'Open Excel sheet', '', 'Excel files (*.xls *.xlsx)'
        )
        self.xlsFileBox.setText(self.xlsFile)
        self.set_sheets(self.xlsFile)

    def get_output(self):
        self.outputDir = QFileDialog.getExistingDirectory()
        self.output.setText(self.outputDir)

    def set_sheets(self, wkbk):
        self.workbook = xlrd.open_workbook(wkbk)
        self.sheetCmb.clear()
        self.sheetCmb.addItems(self.workbook.sheet_names())

    def create_metadata(self):
        try:
            worksheet = self.workbook.sheet_by_name(self.sheetCmb.currentText())

            c = 0

            xmlStr = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<oai_dc:dc\n" \
            + "xsi:schemaLocation=\"http://www.openarchives.org/OAI/2.0/oai_dc/ oai_dc.xsd\""   \
            + "\nxmlns:dc=\"http://purl.org/dc/elements/1.1/\" "   \
            + "\nxmlns:oai_dc=\"http://www.openarchives.org/OAI/2.0/oai_dc/\"" \
            + "\nxmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"

            for x in range(1, worksheet.nrows):
                fileName = worksheet.cell(x, 0).value.replace(" ", "") + ".metadata"
                f = open(self.outputDir + '\\' + fileName, "w+")
                fStr = xmlStr \
                + "\n<dc:title>" + worksheet.cell(x, 1).value + "</dc:title>" \
                + "\n<dc:creator>" + worksheet.cell(x, 2).value + "</dc:creator>" \
                + "\n<dc:subject>" + worksheet.cell(x, 3).value + "</dc:subject>"   \
                + "\n<dc:description>" + worksheet.cell(x, 4).value + "</dc:description>" \
                + "\n<dc:publisher>" + worksheet.cell(x, 5).value + "</dc:publisher>" \
                + "\n<dc:contributor>" + worksheet.cell(x, 6).value + "</dc:contributor>" \
                + "\n<dc:date>" + str(worksheet.cell(x, 7).value) + "</dc:date>" \
                + "\n<dc:type>" + worksheet.cell(x, 8).value + "</dc:type>" \
                + "\n<dc:format>" + worksheet.cell(x, 9).value + "</dc:format>" \
                + "\n<dc:identifier>" + str(worksheet.cell(x, 10).value) + "</dc:identifier>" \
                + "\n<dc:source>" + worksheet.cell(x, 11).value + "</dc:source>" \
                + "\n<dc:language>" + worksheet.cell(x, 12).value + "</dc:language>\n</oai_dc:dc>"

                f.write(fStr)
                c = c + 1

            doneBox = QMessageBox()
            doneBox.setIcon(QMessageBox.Information)
            doneBox.setText(str(c) + " .metadata files created.")
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
