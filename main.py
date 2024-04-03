print("Start")
from time import time
print("-"*30)
print("Starting Imports")
t1 = time()
from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow, QLabel, QApplication, QTreeWidgetItem, QTableWidgetItem, QAbstractItemView, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QFont, QColor
from PyQt5.QtPrintSupport import QPrintDialog, QPrinter
from PyPDF2 import PdfMerger
# from pdf2image import convert_from_path
import pandas as pd
from cryptography.fernet import Fernet
fernetKey = Fernet(b'NDeVZN2769bnfxgI51FzSuXEGjNpIl1SvxauVNtUGKs=')
import os, glob, shutil, subprocess
t2 = time()
print(f"Imports Runtime = {t2-t1:.2f}s")
print("-"*30)
print("Starting TempDir Setup")
t1 = time()
import tempfile
for filename in glob.glob(tempfile.gettempdir() + "\\" + "FBTemp%f%*"):
    shutil.rmtree(filename)
tempDir = tempfile.TemporaryDirectory(prefix="FBTemp%f%")
t2 = time()
print(f"TempDir Runtime = {t2-t1:.2f}s")
print("-"*30)
class finalBook(QMainWindow):
    def __init__(self):
        print("Starting GUI Init")
        t1 = time()
        super(finalBook, self).__init__()
        # load the ui file
        loadUi("GUI.ui", self)
        # config
        self.encState = True
        self.prefix = ".finalbookpak"
        self.resourceFolderPath = 'Resources'
        self.Errs = ["_", "ERR: Record Not Found", "ERR: Duplicated Record", "ERR: File Not Found", "ERR: Unknown"]
        t2 = time()
        print(f"GUI Init Runtime = {t2-t1:.2f}s")
        print("-" * 30)

        # load App
        self.loadApp()
        # Connect Signals
        print("Starting Signals Init")
        t1 = time()
        ## Actions
        self.refreshAction.triggered.connect(self.close)
        self.refreshAction.triggered.connect(self.__init__)
        ###### setting action ######
        self.manualAction.triggered.connect(lambda: self.openFile(f'Resources\\manual{self.prefix}'))
        self.closeAction.triggered.connect(self.close)

        ## Buttons
        self.refreshBtn.clicked.connect(self.loadApp)
        ###### setting button ######
        self.manualBtn.clicked.connect(lambda: self.openFile(f'Resources\\manual{self.prefix}'))
        self.exitBtn.clicked.connect(self.close)
        self.docsTreeWidget.itemDoubleClicked.connect(self.handler)
        self.docsTableWidget.itemDoubleClicked.connect(self.tableHandler)
        self.viewMergedBtn.clicked.connect(self.veiwMergedPdf)
        # self.printMergedBtn.clicked.connect(self.printMergedPdf)
        self.viewTableBtn.clicked.connect(self.veiwTableXlsx)
        # self.printTableBtn.clicked.connect(self.printTableXlsx)
        t2 = time()
        print(f"Signals Init Runtime = {t2 - t1:.2f}s")
        print("-" * 30)

        try:
            import pyi_splash
            pyi_splash.update_text('UI Loaded ...')
            pyi_splash.close()
        except:
            pass


    # Funcs
    def loadApp(self):
        # Widgets
        ## TreeWidget load
        print("Loading TreeWidget")
        t1 = time()
        self.docsTreeWidget.clear()
        self.loadFolder('Resources\\Files', self.docsTreeWidget)
        self.docsTreeWidget.setHeaderLabel("")
        t2 = time()
        print(f"TreeWidget Runtime = {t2 - t1:.2f}s")
        print("-" * 30)

        ## TableWidget load
        print("Loading TableWidget")
        t1 = time()
        self.docsTableWidget.clear()
        self.loadTable('Resources\\Tables', self.docsTableWidget)
        t2 = time()
        print(f"TableWidget Runtime = {t2 - t1:.2f}s")
        print("-" * 30)

        ## project info Labels load and set
        print("Loading Info")
        t1 = time()
        self.cleanLayput(self.projectInfoGridLayout)
        # self.cleanLayput(self.productInfoGridLayout)
        self.loadInfo()
        t2 = time()
        print(f"Info Runtime = {t2 - t1:.2f}s")
        print("-" * 30)

        self.show()
    def handler(self, item, column_no):
        par = item
        path = ""
        while par.parent() != None:
            path = par.parent().text(0)+ "\\" + path
            par = par.parent()

        if item.childCount() == 0:
            self.openFile(self.resourceFolderPath + "\\" + "Files" + "\\" + path + "\\" + item.text(0) + self.prefix)

    def tableHandler(self, item):
        if item.text() not in self.Errs:
            if item.column() == 1 and self.t == 2:
                self.openFile(self.resourceFolderPath + "\\" + "Tables" + "\\" + self.df.columns[item.column()] + "\\" + self.df.iloc[item.row(),item.column()] + self.prefix)
            else:
                self.openFile(self.resourceFolderPath + "\\" + "Tables" + "\\" + self.df.columns[item.column()] + "\\" + self.df.iloc[item.row(),0] + self.prefix)

        #print(self.df.iloc[item.row(),0])

    def openFile(self, path, prefix = ".pdf", encState = True):
        print(f"Openning File:{path}")
        t1 = time()
        if encState:
            self.decryptFile(path, prefix=prefix)
            subprocess.Popen(tempDir.name + '\\' + path.split(".")[0] + prefix, shell=True)
        else:
            subprocess.Popen(path, shell=True)
        t2 = time()
        print(f"Open Runtime = {t2 - t1:.2f}s")
        print("-" * 30)

    def cleanLayput(self, layout):
        for i in reversed(range(layout.count())):
            layout.itemAt(i).widget().deleteLater()

    def loadFolder(self, startpath, tree):
        for element in os.listdir(startpath):
            path_info = startpath + "\\" + element
            parent_itm = QTreeWidgetItem(tree, [os.path.basename(element).split(".")[0]])
            if os.path.isdir(path_info):
                self.loadFolder(path_info, parent_itm)
                parent_itm.setIcon(0, QIcon('Images\\folder.png'))
                parent_itm.setFont(0, QFont("B Nazanine", 10, QFont.Bold))
                parent_itm.setExpanded(True)
            else:
                parent_itm.setIcon(0, QIcon('Images\\file.png'))

    def loadTable(self, startpath, Table):
        self.df = pd.DataFrame(columns = os.listdir(startpath))
        self.df['1-ID'] = os.listdir(startpath + "\\" + "1-ID")

        self.t = 1
        if os.path.isdir("Resources\\Tables\\2-Coil Cert"):
            self.t = 2
            if self.encState:
                self.decryptFile(self.resourceFolderPath + "\\" + "Coil_Certificate.finalbookpak", prefix=".xlsx")
                coilCert = pd.read_excel(tempDir.name + "\\" + self.resourceFolderPath + "\\" + "Coil_Certificate.xlsx")
            else:
                coilCert = pd.read_excel(self.resourceFolderPath + "\\" + "Coil_Certificate.xlsx")

            coilCertDirList = os.listdir("Resources\\Tables\\2-Coil Cert")
            coilCertDirList = list(map(lambda x:x.split(".")[0], coilCertDirList))

            for n, id in enumerate(self.df['1-ID']):
                tempTuple = coilCert['Coil ID'].loc[coilCert['Pipe ID'] == int(id.split('.')[0])].values
                try:
                    if str(tempTuple.item()) not in coilCertDirList:
                        self.df.iloc[n, 1] = "ERR: File Not Found"
                    else:
                        self.df.iloc[n, 1] = str(tempTuple.item())
                except:
                    if tempTuple.shape == (0,):
                        self.df.iloc[n, 1] = "ERR: Record Not Found"
                    elif tempTuple.shape > (1,):
                        self.df.iloc[n, 1] = "ERR: Duplicated Record"
                    else:
                        self.df.iloc[n, 1] = "ERR: Unknown"


        for col in self.df.iloc[:,self.t:].columns:
            tempList = []
            for i in self.df['1-ID']:
                if i in os.listdir(startpath + "\\" + col):
                    tempList.append(u'\u2713')
                else:
                    tempList.append('_')
            self.df[col] = tempList.copy()
        self.df['1-ID'] = self.df['1-ID'].apply(lambda x: x.split('.')[0])

        Table.setColumnCount(self.df.shape[1])
        Table.setRowCount(self.df.shape[0])
        Table.setHorizontalHeaderLabels(map(lambda x:x.split("-")[-1], self.df.columns.tolist()))

        for r in range(self.df.shape[0]):
            for c in range(self.df.shape[1]):
                tableItem = QTableWidgetItem(str(self.df.iloc[r,c]))
                tableItem.setTextAlignment(Qt.AlignCenter)
                if self.df.iloc[r,c] in self.Errs[1:]:
                    tableItem.setBackground(QColor(255, 90, 90))
                Table.setItem(r, c, tableItem)

        Table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        #Table.resizeColumnsToContents()
        #print(self.df)

    def loadInfo(self):
        if self.encState:
            self.decryptFile(self.resourceFolderPath + "\\" + "info.finalbookpak", prefix=".xlsx")
            self.projectInfo = pd.read_excel(tempDir.name + "\\" + self.resourceFolderPath + "\\" + "info.xlsx", sheet_name=0)
            self.productInfo = pd.read_excel(tempDir.name + "\\" + self.resourceFolderPath + "\\" + "info.xlsx", sheet_name=1)

        else:
            self.projectInfo = pd.read_excel(self.resourceFolderPath + "\\" + "info.xlsx", sheet_name=0)
            self.productInfo = pd.read_excel(self.resourceFolderPath + "\\" + "info.xlsx", sheet_name=1)

        self.projectInfo = self.projectInfo.values.astype("str")
        self.productInfo = self.productInfo.values.astype("str")

        for i, row in enumerate(self.projectInfo):
            for j, v in enumerate(row):
                tempLabel = QLabel(v)
                if j == 0:
                    tempLabel.setText(tempLabel.text() + ":")
                    tempLabel.setFont(QFont("B Nazanine", 12, QFont.Bold))
                    tempLabel.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    tempLabel.setFont(QFont("B Nazanine", 10))
                    tempLabel.setAlignment(Qt.AlignCenter)
                self.projectInfoGridLayout.addWidget(tempLabel, i, 1 - j)

        for i, row in enumerate(self.productInfo):
            for j, v in enumerate(row):
                tempLabel = QLabel(v)
                if j == 0:
                    tempLabel.setText(tempLabel.text() + ":")
                    tempLabel.setFont(QFont("B Nazanine", 12, QFont.Bold))
                    tempLabel.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                else:
                    tempLabel.setFont(QFont("B Nazanine", 10))
                    tempLabel.setAlignment(Qt.AlignCenter)
                self.productInfoGridLayout.addWidget(tempLabel, i, 1 - j)

        self.infoVerticalLayout.setStretch(0, self.projectInfo.shape[0] + 1)
        self.infoVerticalLayout.setStretch(1, self.productInfo.shape[0] + 1)
        self.projectInfoGridLayout.setColumnStretch(0, 100)
        self.projectInfoGridLayout.setColumnStretch(1, 41)
        self.productInfoGridLayout.setColumnStretch(0, 100)
        self.productInfoGridLayout.setColumnStretch(1, 41)


    def getSelected(self):
        itemDirs = []
        selected = list(map(lambda x:(x.row(),x.column()), self.docsTableWidget.selectedItems()))
        if len(selected) != 0:
            for cord in selected:
                if self.df.iloc[cord] not in self.Errs:
                    if cord[1] == 1 and self.t == 2:
                        itemDirs.append(f'Resources\\Tables\\{self.df.columns[cord[1]]}\\{self.df.iloc[cord]}{self.prefix}')
                    else:
                        itemDirs.append(f'Resources\\Tables\\{self.df.columns[cord[1]]}\\{self.df.iloc[cord[0], 0]}{self.prefix}')
        return itemDirs

    def veiwMergedPdf(self):
        self.progressBar.setValue(0)
        merger = PdfMerger()
        gotSelected = self.getSelected()
        if len(gotSelected) != 0:
            for i, pdf in enumerate(gotSelected):
                self.decryptFile(pdf)
                merger.append(tempDir.name + "\\" + pdf.split(".")[0] + ".pdf")
                self.progressBar.setValue(int((i+1)*100/len(gotSelected)))

            merger.write(tempDir.name + '\\merged.pdf')
            self.openFile(tempDir.name + '\\merged.pdf', encState=False)
        else:
            self.popBox("حداقل یک مورد را انتخاب کنید!")
        merger.close()
        self.progressBar.setValue(0)

    def veiwTableXlsx(self):
        self.df.to_excel(tempDir.name + "\\docsTable.xlsx", sheet_name="جدول مدارک")
        self.openFile(tempDir.name + "\\docsTable.xlsx", encState=False)


    def popBox(self, msg):
        pop = QMessageBox()
        pop.setWindowTitle(" ")
        pop.setText(msg)
        pop.setIcon(QMessageBox.Warning)
        pop.exec_()

    def decryptFile(self, path, prefix = ".pdf", isFolder = False):
        if isFolder:
            for element in os.listdir(path):
                path_info = path + "\\" + element
                if os.path.isdir(path_info):
                    self.decryptFile(path_info)
                else:
                    with open(path_info, "rb") as enc_file:
                        enc_bytes = enc_file.read()
                        dec_bytes = fernetKey.decrypt(enc_bytes)
                        if not os.path.isdir(tempDir.name + '\\' + path):
                            os.makedirs(tempDir.name + '\\' + path)
                        with open(tempDir.name + '\\' + path_info.split(".")[0] + prefix, "wb") as dec_file:
                            dec_file.write(dec_bytes)

        else:
            with open(path, "rb") as enc_file:
                enc_bytes = enc_file.read()
                dec_bytes = fernetKey.decrypt(enc_bytes)
                indx = path.rfind("\\")
                if not os.path.isdir(tempDir.name + '\\' + path[:indx]):
                    os.makedirs(tempDir.name + '\\' + path[:indx])
                with open(tempDir.name + '\\' + path.split(".")[0] + prefix, "wb") as dec_file:
                    dec_file.write(dec_bytes)

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    MainWindow = finalBook()
    ret = app.exec_()

    ##### Remove Temp #####
    tempDir.cleanup()
    sys.exit(ret)