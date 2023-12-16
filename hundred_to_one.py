import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QHBoxLayout, QVBoxLayout, QTabWidget,
    QTableWidgetItem, QMenuBar, QFileDialog, QColorDialog, QLineEdit, QWidget, QPushButton
)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont, QColor

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

        self.setColumnWidth(0, 733)
        self.setColumnWidth(1, 100)

        row = 0
        for word in sorted(data, key = lambda elem: data[elem], reverse = True):
            self.setRowHeight(row, 50)

            wordItem = QTableWidgetItem(word)
            countItem = QTableWidgetItem(str(data[word]))

            wordItem.setFont(QFont("Times new roman", 20))
            wordItem.setTextAlignment(Qt.AlignCenter)
            wordItem.setFlags(Qt.ItemIsEnabled)
            wordItem.setBackground(QColor(255,255,255))
            
            countItem.setFont(QFont("Times new roman", 20))
            countItem.setTextAlignment(Qt.AlignCenter)
            countItem.setFlags(Qt.ItemIsEnabled)
            countItem.setBackground(QColor(255,255,255))

            self.setItem(row, 0,  wordItem)
            self.setItem(row, 1, countItem)

            row += 1
        

        self.horizontalHeader().hide()
        self.verticalHeader().hide()

        self.cellClicked.connect(self.chooseWord)
        self.chosenWordIndexes = []
    
    def preprocessing(self, wordList):
        wordSet = set(wordList)
        treatedWords = {}
        for word in wordSet:
            treatedWords[word] = wordList.count(word)
        
        return treatedWords
    
    def chooseWord(self):
        row = self.currentRow()
        if row not in self.chosenWordIndexes:
            self.item(row, 0).setBackground(QColor(255,100,0))
            self.item(row, 1).setBackground(QColor(255,100,0))
            self.chosenWordIndexes.append(row)
        else:
            self.item(row, 0).setBackground(QColor(255,255,255))
            self.item(row, 1).setBackground(QColor(255,255,255))
            self.chosenWordIndexes.remove(row)
        
        print(self.chosenWordIndexes)

class Page(QWidget):
    def __init__(self, parent, data):
        super().__init__(parent)
        self.pageLay = QHBoxLayout(self)
        self.setLayout(self.pageLay)
        self.answerTable = AnswersTable(self, data)
        self.pageLay.addWidget(self.answerTable, stretch = 2)
        
        self.wordForUnionLine = QLineEdit(self)
        self.wordForUnionLine.setFont(QFont("Times new roman", 20))
        self.btn = QPushButton(self)
        self.rightLayout = QVBoxLayout(self)
        self.rightLayout.addWidget(self.wordForUnionLine, stretch = 1)
        self.rightLayout.addWidget(self.btn)
        self.rightLayout.addWidget(QWidget(self))

        self.pageLay.addLayout(self.rightLayout, stretch = 1)



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
        self.setCentralWidget(self.pageTape)

        self.showMaximized()
        self.show()
    
    def preImportCheck(self, exel):
        for column in "BCDEFG":
            row = 1
            while exel['A{}'.format(row)].value != None:
                if type(exel['{}{}'.format(column, row)].value) != str:
                    raise ValueError(f"Incorrect value in cell {column}{row}")
                row += 1
    
    def importFile(self):
        filename, _ = QFileDialog.getOpenFileName(self, "Import answers", ".", "Exel files (*.xlsx)")
        if filename:
            #try:
            wb = load_workbook(filename)
            exel = wb.active
            self.preImportCheck(exel)
                
            #pageInd = 0
            for column in "BCDEFG":
                row = 2
                data = []
                while exel['A{}'.format(row)].value != None:
                    data.append(exel['{}{}'.format(column, row)].value.strip())
                    row += 1
                
                self.pageTape.addTab(Page(self.pageTape, data), exel['{}1'.format(column)].value)
            #self.answerTable = AnswersTable(self, wb.active)
            wb.close()
            #except:
                #raise ValueError("Incorrect file values") 


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())