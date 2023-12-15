import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QHBoxLayout, QVBoxLayout, QTabWidget,
    QTableWidgetItem, QMenuBar, QFileDialog, QColorDialog, QLineEdit, QWidget
)
from PyQt5.QtCore import Qt, QSize
from PyQt5 import QtGui

from openpyxl import Workbook, load_workbook

class AnswersTable(QTableWidget):
    def __init__(self, parent, data = None):
        super().__init__(parent)

        if data == None:
            self.row_count = 1
        else:
            data = self.preprocessing(data)
            self.row_count = len(data)

        self.setColumnCount(2)
        self.setRowCount(self.row_count)

        self.setColumnWidth(0, 500)
        self.setColumnWidth(1, 100)

        row = 0
        for word in sorted(data, key = lambda elem: data[elem], reverse = True):
            self.setRowHeight(row, 50)

            wordItem = QTableWidgetItem(word)
            countItem = QTableWidgetItem(str(data[word]))
            self.setItem(row, 0,  wordItem)
            self.setItem(row, 1, countItem)

            row += 1
        

        self.horizontalHeader().hide()
        self.verticalHeader().hide()
    
    def preprocessing(self, wordList):
        wordSet = set(wordList)
        treatedWords = {}
        for word in wordSet:
            treatedWords[word] = wordList.count(word)
        
        return treatedWords

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle('hundred to one')
        #self.setWindowIcon(QIcon('./assets/usergroup.png'))
        self.setGeometry(100, 100, 1100, 800)

        self.panel = QMenuBar(self)
        self.setMenuBar(self.panel)
        importAct = self.panel.addAction("Импортировать ответы")
        importAct.triggered.connect(self.importFile)
        
        self.pageTape = QTabWidget(self)
        self.mainLay = QHBoxLayout()
        self.pageTape.setLayout(self.mainLay)
        #self.mainLay.addWidget(QLineEdit(self))
        #self.mainLay.addWidget(QLineEdit(self))
        self.setCentralWidget(self.pageTape)

        self.showMaximized()
        self.show()
    
    def importFile(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Import answers", ".", "Exel files (*.xlsx)")
        if filename:
            #try:
            wb = load_workbook(filename)
            exel = wb.active
                
            #pageInd = 0
            for column in "BCDEFG":
                row = 2
                data = []
                while exel['A{}'.format(row)].value != None:
                    data.append(exel['{}{}'.format(column, row)].value.strip())
                    row += 1
                self.pageTape.addTab(AnswersTable(self.pageTape, data), exel['{}1'.format(column)].value)
            #self.answerTable = AnswersTable(self, wb.active)
            wb.close()
            #except:
                #raise ValueError("Incorrect file values") 


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())