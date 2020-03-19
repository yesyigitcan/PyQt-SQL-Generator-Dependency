import sys
sys.path.append("..\\Lib\\site-packages")
from PyQt5 import QtWidgets, uic, QtGui
from PyQt5.QtWidgets import QFileDialog, QTableWidget, QTableWidgetItem, QTableView, QMessageBox, QProgressDialog, QCheckBox, QLineEdit, QMainWindow, QTextBrowser
from PyQt5.Qt import QApplication
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import PyQt5.QtCore as QtCore
import sys
import xlrd
import cx_Oracle
import random
import webbrowser
import csv
import io
import getpass
import time
class MainFrame(QtWidgets.QMainWindow):
    def __init__(self, ui_name):
        QtWidgets.QMainWindow.__init__(self)
        self.win = uic.loadUi(ui_name)  # specify the location of your .ui file
        # self.win.show()
        self.win.pushButton.clicked.connect(self.getFileExplorer)
        self.win.pushButton_2.clicked.connect(self.acceptButtonClicked)
        self.win.pushButton_3.clicked.connect(self.fillTable)
        self.win.pushButton_4.clicked.connect(self.deleteSelectedRows)
        self.win.pushButton_5.clicked.connect(self.autoFillClicked)
        self.win.pushButton_6.clicked.connect(self.deleteNullRows)
        self.win.actionOpen_File.triggered.connect(self.getFileExplorer)
        self.fileisprevious = 0
        self.win.actionOpen_Previous_File.triggered.connect(self.openPreviousFile)
        self.win.actionSend_Mail_for_Help.triggered.connect(self.sendMailForHelp)
        self.win.actionSend_Bug_Notice.triggered.connect(self.sendBugNotice)
        #self.win.actionShare_via_Email.triggered.connect(self.shareViaEmail)
        self.win.closeEvent = self.exitEvent
        self.win.tableWidget.itemChanged.connect(self.editTrigger)  # When any item in table is changed
        self.win.tableWidget.itemSelectionChanged.connect(self.selectionTrigger)
        self.win.tableWidget.customContextMenuRequested.connect(self.rightClickComboBoxMenu) # Right button clicked
        self.win.tableWidget.keyPressEvent = self.keyPressEvent
        self.autoGrayDeleteStatus = 0
        self.previousFilePath = "NONE"
        self.processedFilePath = ""
        try:
            with open('Initial.txt','r') as file:
                if int(file.readline().split(',')[1].strip()):
                    self.win.autoGrayDeleteCheckBox.setChecked(True)
                    self.autoGrayDeleteStatus = 1
                try:
                    self.previousFilePath = file.readline().split(',')[1].strip()
                except Exception as e:
                    print(str(e))
        except Exception as e:
            print(str(e))
        self.win.autoGrayDeleteCheckBox.stateChanged.connect(self.autoGrayDeleteCheckBoxTrigger)
        pixmap = QPixmap("D:\\Users\\VB85788\\PycharmProjects\\pyqt_gui\\Icon\\cursor.png")
        scaled_pixmap = pixmap.scaled(QSize(15,15), Qt.KeepAspectRatio)
        cursor = QCursor(scaled_pixmap, -1, 0)
        self.win.tableWidget.setCursor(cursor)
        self.table_header = self.win.tableWidget.horizontalHeader()
        self.table_header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.table_header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        self.table_header.setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        self.table_header.setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        self.table_header.setSectionResizeMode(4, QtWidgets.QHeaderView.Stretch)
        self.table_header.setSectionResizeMode(5, QtWidgets.QHeaderView.Stretch)
        self.table_header.setSectionResizeMode(6, QtWidgets.QHeaderView.Stretch)
        self.table_header.setSectionResizeMode(7, QtWidgets.QHeaderView.Stretch)
        self.hints = [
            "Ipucu: Excel Listenizde Olmayan Fakat Tabloya Gelen Herhangi Bir Satırı Update veya Insert Basmak için Önce Bir Hücresini Boş Yapın (Satır Sarı Olacaktır). Daha Sonra Doldurun (Update için Satır Yeşil Olacaktır, Insert için Satır Mavi Olacaktır).",
            "Ipucu: Excel Listenizden Gelen Satırlar DB'den Gelen Satırların Altında Yer Alır.",
            "Renkler: Açık Gri - Excelde Yok, DBde Var, Activelik E\tKoyu Gri - Excelde Yok, DBde Var, Activelik H\tSarı - Excelde Var, DBde Yok\tMavi - Üzerinde Düzenleme Yapılmış Satır, Insert Olacak\tYeşil - Excelde Var, DBde Var, Update Olacak",
            "Ipucu: Tablodan Tabular Formatta Kopyalama Yapmak İçin İlgili Sütunları Seçerek Ctrl+C Yapmanız Yeterlidir.",
            "Ipucu: Dilediğiniz Satırı Seçerek Ctrl+U ile update satırı (Satır Yeşil Olacaktır) yapabilir veya Ctrl+I ile insert satırı (Satır Mavi Olacaktır) yapabilirsiniz. Satırları pasif yapmak için direk silebilirsiniz.",
            "Ipucu: Herhangi bir bağımlılık işlemi için talep sahibinden prosedür ismi, şema adı ve tablo adı nı istemeyi unutmayın. Active ve Period sütunları default değer üretmektedir."
        ]
        self.win.hintText.setText("Ipucu: ' (u) ' olan sütunlardaki değişikler update olur, ' (i) ' olan sütunlardaki değişiklikler insert olur.")
        self.previousHintIndex = -1
        self.win.showMaximized()
    def popupMessage(self, title, message):
        msgBox = QMessageBox()
        msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
        msgBox.setIcon(QMessageBox.Warning)
        msgBox.setText(message)
        msgBox.setWindowTitle(title)
        msgBox.exec()
    def openPreviousFile(self):
        if self.previousFilePath != "NONE":
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Question)
            msg.setText("Are you sure to open previous file?\nPrevious file path: " + self.previousFilePath)
            msg.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
            msg.setWindowTitle("Open Previous File")
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            msg.buttonClicked.connect(self.openPreviousFileClickTrigger)
            msg.exec()
    def openPreviousFileClickTrigger(self,e):
        if e.text() == "&Yes":
            self.fileisprevious = 1
            self.getFileExplorer()
    def sendMailForHelp(self):
        webbrowser.open("mailto:<<<Mail Address>>>&subject=Dependency Matcher Uygulaması Yardım")
    def sendBugNotice(self):
        webbrowser.open("mailto:<<<Mail Address>>>&subject=Dependency Matcher Uygulaması Bug Bildirimi")
    def autoGrayDeleteCheckBoxTrigger(self):
        if self.win.autoGrayDeleteCheckBox.isChecked():
            self.autoGrayDeleteStatus = 1
        else:
            self.autoGrayDeleteStatus = 0
    def rightClickComboBoxMenu(self, pos):
        menu = QtWidgets.QMenu()
        copy_action = menu.addAction("Copy (Ctrl+C)")
        copy_all_action = menu.addAction("Copy All (Ctrl+Alt+C)")
        convert_to_update_action = menu.addAction("Convert to Update (Ctrl+U)")
        convert_to_insert_action = menu.addAction("Convert to Insert (Ctrl+I)")
        delete_selected_action = menu.addAction("Delete Selected Rows (Delete)")
        delete_null_action = menu.addAction("Delete Null Rows")
        delete_gray_action = menu.addAction("Delete Gray (Unfocused) Rows")
        action = menu.exec_(self.win.tableWidget.mapToGlobal(pos))
        if action == copy_action:
            self.copySelectedRows()
            self.win.hintText.setText("Seçilen Hücreler Kopyalandı!")
        elif action == copy_all_action:
            self.copyAllRows()
            self.win.hintText.setText("Bütün Hücreler Kopyalandı!")
        elif action == convert_to_update_action:
            self.convertToUpdate()
        elif action == convert_to_insert_action:
            self.convertToInsert()
        elif action == delete_selected_action:
            self.deleteSelectedRows()
        elif action == delete_null_action:
            self.deleteNullRows()
        elif action == delete_gray_action:
            self.deleteGrayRows()
    def copySelectedRows(self):
        selection = self.win.tableWidget.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            columns = sorted(index.column() for index in selection)
            rowcount = rows[-1] - rows[0] + 1
            colcount = columns[-1] - columns[0] + 1
            table = [[''] * colcount for _ in range(rowcount)]
            for index in selection:
                row = index.row() - rows[0]
                column = index.column() - columns[0]
                table[row][column] = index.data()
            stream = io.StringIO()
            csv.writer(stream, delimiter='\t').writerows(table)
            QtWidgets.qApp.clipboard().setText(stream.getvalue())
    def copyAllRows(self):
        self.win.tableWidget.selectAll()
        self.copySelectedRows()
    def convertToUpdate(self):
        selection = self.win.tableWidget.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            for row in rows:
                is_empty = 0
                for col in range(8):
                    try:
                        if not self.win.tableWidget.item(row, col).text():
                            is_empty = 1
                            break
                    except:
                        is_empty = 1
                if not is_empty:
                    for col in range(8):
                        self.win.tableWidget.item(row, col).setBackground(QColor(200, 250, 200))
                else:
                    self.win.hintText.setText("Uyarı: Convert işleminin yapılabilmesi için boş hücre olmaması lazım!")
    def convertToInsert(self):
        selection = self.win.tableWidget.selectedIndexes()
        if selection:
            rows = sorted(index.row() for index in selection)
            for row in rows:
                is_empty = 0
                for col in range(8):
                    try:
                        if not self.win.tableWidget.item(row, col).text():
                            is_empty = 1
                            break
                    except:
                        is_empty = 1
                if not is_empty:
                    for col in range(8):
                        self.win.tableWidget.item(row, col).setBackground(QColor(135, 206, 235))
                else:
                    self.win.hintText.setText("Uyarı: Convert işleminin yapılabilmesi için boş hücre olmaması lazım!")
    def keyPressEvent(self, e):
        if QKeySequence(e.key() + int(e.modifiers())) == QKeySequence("Ctrl+C"):
            self.copySelectedRows()
            self.win.hintText.setText("Seçilen Hücreler Kopyalandı!")
        elif QKeySequence(e.key() + int(e.modifiers())) == QKeySequence("Ctrl+Alt+C"):
            self.copyAllRows()
            self.win.hintText.setText("Bütün Hücreler Kopyalandı!")
        elif QKeySequence(e.key() + int(e.modifiers())) == QKeySequence("Ctrl+U"):
            self.convertToUpdate()
        elif QKeySequence(e.key() + int(e.modifiers())) == QKeySequence("Ctrl+I"):
            self.convertToInsert()
        elif QKeySequence(e.key() + int(e.modifiers())) == QKeySequence("Shift+Up"):
            self.selectWithArrowKeys(0, 0)
        elif QKeySequence(e.key() + int(e.modifiers())) == QKeySequence("Shift+Down"):
            self.selectWithArrowKeys(1, 0)
        elif e.key() == 16777223: # if delete key pressed
            self.deleteSelectedRows()
        elif e.key() == 16777235: # if up arrow key pressed
            self.selectWithArrowKeys(0, 1)
        elif e.key() == 16777237: # if down arrow key pressed
            self.selectWithArrowKeys(1, 1)
    def selectWithArrowKeys(self, way = 1, reset = 0):
        selection = self.win.tableWidget.selectedIndexes()
        rowPosition = self.win.tableWidget.rowCount()
        if reset:
            self.win.tableWidget.clearSelection()
        if len(selection) == 0:
            if rowPosition > 0:
                if way == 1: # 0 for up, 1 for down
                    for i in range(8):
                        try:
                            self.win.tableWidget.selectRow(0)
                        except:
                            pass
                elif way == 0:
                    for i in range(8):
                        self.win.tableWidget.selectRow(rowPosition-1)
                else:
                    raise Exception("Selection Way Has to Be Either 0 for up or 1 for down")
            else:
                return
        else:
            if rowPosition > 0:
                if way == 1: # 0 for up, 1 for down
                    if selection[-1].row() + 1 > rowPosition:
                        self.win.tableWidget.selectRow(rowPosition - 1)
                    else:
                        for i in range(8):
                            try:
                                self.win.tableWidget.selectRow(selection[-1].row() + 1)
                            except:
                                pass
                elif way == 0:
                    if selection[0].row() - 1 < 0:
                        self.win.tableWidget.selectRow(0)
                    else:
                        self.win.tableWidget.selectRow(selection[0].row() - 1)
                else:
                    raise Exception("Selection Way Has to Be Either 0 for up or 1 for down")
            else:
                return
    def selectionTrigger(self):
        selection = self.win.tableWidget.selectedIndexes()
        if len(selection) > 0:
            self.win.hintText.setText("Seçilen Hücre Sayısı: " + str(len(selection)))
        else:
            self.win.hintText.setText("Uyarı: Accept Butonuna Basmadan Önce Tablodaki Bilgilerin Doğruluğundan Emin Olun.")
        return
    def editTrigger(self, item):
        try:
            self.win.tableWidget.itemChanged.disconnect()
        except Exception as e:
            print(str(e))
        is_empty = 0
        for i in range(8):
            try:
                if not self.win.tableWidget.item(item.row(), i).text():
                    is_empty = 1
                    break
            except:
                is_empty = 1
        if is_empty:
            for i in range(8):
                try:
                    self.win.tableWidget.item(item.row(), i).setBackground(QColor(250, 250, 1))
                except:
                    pass
        else:
            color = self.win.tableWidget.item(item.row(), (item.column() + 1) // 7).background()
            if color == QColor(250, 250, 1) or color == QColor(255, 0,0):  # if color yellow or red and no empty cell, make'em all green
                # blue QColor(135, 206, 235)
                # green QColor(200, 250, 200)
                if item.column() in (2,3,4): # if you change procedure name, schema or table it becomes insert, else it becomes update
                    for i in range(8):
                        self.win.tableWidget.item(item.row(), i).setBackground(QColor(135, 206, 235))
                else:
                    for i in range(8):
                        self.win.tableWidget.item(item.row(), i).setBackground(QColor(200, 250, 200))
            else:
                for i in range(8):
                    self.win.tableWidget.item(item.row(), i).setBackground(color)
        try:
            self.win.tableWidget.itemChanged.connect(self.editTrigger)
        except Exception as e:
            print(str(e))
    def updateTableWithDB(self, procedure_name_list):
        con = cx_Oracle.connect('<<<MASKED>>>')
        cur = con.cursor()
        row_index = 0
        for procedure_name in procedure_name_list:
            # VITDWH
            query = "select ACTIVE, PERIOD, PROSEDUR_ISMI, OWNER, TABLE_NAME, BAGLI_WORKFLOW, BAGLI_SESSION, FOLDER_ADI from <<<MASKED>>> where PROSEDUR_ISMI = '" + procedure_name + "'"
            cur.execute(query)
            data = cur.fetchone()
            rowPosition = self.win.tableWidget.rowCount()
            while data:
                if str(data[2]) not in self.global_procedure_list:
                    self.global_procedure_list.append(str(data[2]))
                self.win.tableWidget.insertRow(rowPosition)
                if str(data[0]) == "E":
                    color = QColor(150, 150, 150)
                elif str(data[0]) == "H":
                    color = QColor(100, 100, 100)
                else:
                    color = QColor(100, 100, 100)
                for i in range(8):
                    self.win.tableWidget.setItem(row_index, i, QTableWidgetItem(str(data[i])))
                    self.win.tableWidget.item(row_index, i).setBackground(color)
                rowPosition += 1
                row_index += 1
                data = cur.fetchone()
            con.commit()
        cur.close()
        con.close()
    def getFileExplorer(self):
        self.win.tableWidget.clearSelection()
        randomIndex = random.randint(0, len(self.hints) - 1)
        while randomIndex == self.previousHintIndex and len(self.hints) > 1:
            randomIndex = random.randint(0, len(self.hints) - 1)
        self.previousHintIndex = randomIndex
        self.win.hintText.setText(self.hints[randomIndex])
        try:
            self.win.tableWidget.itemChanged.disconnect()
        except Exception as e:
            print(str(e))
        # label.setText()
        self.global_procedure_list = []
        if self.fileisprevious:
            if self.previousFilePath == "NONE":
                self.win.hintText.setText("Uyarı: Önceki Dosya Konumu Bulunamadı!")
                self.fileisprevious = 0
                return
            else:
                filename = self.previousFilePath
        else:
            try:
                filename = QFileDialog.getOpenFileName(caption="Choose an Excel File", directory="",
                                                       filter="Excel File(*.xlsx);;Excel File(*.xls)")[0]
            except Exception as e:
                print(str(e), " :: Function Get File Explorer (1)")
                try:
                    self.win.tableWidget.itemChanged.connect(self.editTrigger)
                except Exception as e:
                    print(str(e))
                return
            if not filename:
                try:
                    self.win.tableWidget.itemChanged.connect(self.editTrigger)
                except Exception as e:
                    print(str(e))
                return
        self.previousFilePath = filename
        self.processedFilePath = filename
        if self.fileisprevious:
            self.win.hintText.setText("Bilgilendirme: Önceki Dosyanın Bulunduğu Dizin: " + filename)
        else:
            self.win.hintText.setText("Uyarı: Accept Butonuna Basmadan Önce Tablodaki Bilgilerin Doğruluğundan Emin Olun.")
        # Delete Rows If They Already Exist. It Means a New Start
        while self.win.tableWidget.rowCount() != 0:
            self.win.tableWidget.removeRow(0)
        self.win.tableWidget.setSortingEnabled(False)
        wb = xlrd.open_workbook(filename)
        sheet = wb.sheet_by_index(0)
        if sheet.row_len(0) > 5 or sheet.row_len(0) < 3:
            msgBox = QMessageBox()
            msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
            msgBox.setIcon(QMessageBox.Warning)
            msgBox.setText("You need to provide file content constraints. Procedure name, schema and table name are must. Active and period are optional.") # bak buraya
            msgBox.setWindowTitle("Invalid File Content")
            msgBox.exec()
            #self.progress.hide()
            self.fileisprevious = 0
            print("Column Size:", sheet.row_len(0))
            return
        try:
            if sheet.cell_value(0, 3) and sheet.cell_value(0,4):
                pass
        except:
            msgBox = QMessageBox()
            msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
            msgBox.setIcon(QMessageBox.Warning)
            msgBox.setText(
                "File does not have active or period column. So some cells will be filled automatically with E or 0")  # bak buraya
            msgBox.setWindowTitle("Auto Fill in Some Columns")
            msgBox.exec()
        procedure_names = []
        for i in range(1, sheet.nrows):
            try:
                procedure_name = str(sheet.cell_value(i, 0)).strip()
                if procedure_name not in procedure_names:
                    procedure_names.append(procedure_name)
            except Exception as e:
                print(str(e), " :: Function Get File Explorer (2)")
        self.updateTableWithDB(procedure_names)
        is_col_miss = 0
        for i in range(1, sheet.nrows):
            try:
                if not sheet.cell_value(i, 0) or not sheet.cell_value(i, 1) or not sheet.cell_value(i, 2):
                    is_col_miss = 1
                    continue
            except:
                is_col_miss = 1
                continue
            try:
                exist_flag = 0
                for j in range(0, self.win.tableWidget.rowCount()):
                    if self.win.tableWidget.item(j, 4).text() == str(
                            sheet.cell_value(i, 2)).strip().upper() and self.win.tableWidget.item(j, 3).text() == str(
                            sheet.cell_value(i, 1)).strip().upper() and self.win.tableWidget.item(j, 2).text() == str(
                            sheet.cell_value(i, 0)).strip().upper():
                        exist_flag = 1
                        over_index = j
                        break
                rowPosition = self.win.tableWidget.rowCount()
                if exist_flag == 0:
                    self.win.tableWidget.insertRow(rowPosition)
                    try:
                        if str(sheet.cell_value(0, 3)).upper() == "ACTIVE":
                            try:
                                if str(sheet.cell_value(0, 4)).upper() == "PERIOD":
                                    if str(sheet.cell_value(i,3)):
                                        self.win.tableWidget.setItem(rowPosition, 0,
                                                                 QTableWidgetItem(str(sheet.cell_value(i, 3))))
                                    else:
                                        self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem("E"))
                                    if str(sheet.cell_value(i,4)):
                                        self.win.tableWidget.setItem(rowPosition, 1,
                                                                 QTableWidgetItem(str(sheet.cell_value(i, 4)).strip().split('.')[0]))
                                    else:
                                        self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem("0"))
                                else:
                                    if str(sheet.cell_value(i,3)):
                                        self.win.tableWidget.setItem(rowPosition, 0,
                                                                 QTableWidgetItem(str(sheet.cell_value(i, 3))))
                                    else:
                                        self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem("E"))
                                    self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem("0"))
                            except:
                                if str(sheet.cell_value(i, 3)):
                                    self.win.tableWidget.setItem(rowPosition, 0,
                                                                 QTableWidgetItem(str(sheet.cell_value(i, 3))))
                                else:
                                    self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem("E"))
                                self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem("0"))
                        else:
                            try:
                                if str(sheet.cell_value(0, 3)).upper() == "PERIOD":
                                    self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem("E"))
                                    if str(sheet.cell_value(0, 4)):
                                        self.win.tableWidget.setItem(rowPosition, 1,QTableWidgetItem(str(sheet.cell_value(i, 4)).strip().split('.')[0]))
                                    else:
                                        self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem("0"))
                                else:
                                    self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem("E"))
                                    self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem("0"))
                            except:
                                self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem("E"))
                                self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem("0"))
                    except Exception as e:
                        print("Error Exist Flag 0 ", e)
                        self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem("E"))
                        self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem("0"))
                    self.win.tableWidget.item(rowPosition, 0).setBackground(QColor(250, 250, 1))
                    self.win.tableWidget.item(rowPosition, 1).setBackground(QColor(250, 250, 1))
                    self.win.tableWidget.setItem(rowPosition, 2, QTableWidgetItem(str(sheet.cell_value(i, 0)).upper().strip()))
                    self.win.tableWidget.item(rowPosition, 2).setBackground(QColor(250, 250, 1))
                    self.win.tableWidget.setItem(rowPosition, 3, QTableWidgetItem(str(sheet.cell_value(i, 1)).upper().strip()))
                    self.win.tableWidget.item(rowPosition, 3).setBackground(QColor(250, 250, 1))
                    self.win.tableWidget.setItem(rowPosition, 4, QTableWidgetItem(str(sheet.cell_value(i, 2)).upper().strip()))
                    self.win.tableWidget.item(rowPosition, 4).setBackground(QColor(250, 250, 1))
                else:
                    row_data = []
                    for col in range(8):
                        row_data.append(self.win.tableWidget.item(over_index, col).text())
                    # update olanları silmek yerine önceki halini de üste bırakmaya karar verdim self.win.tableWidget.removeRow(over_index)
                    # yukarıdaki olayı kaldırdım 11/03/2020
                    self.win.tableWidget.removeRow(over_index)
                    rowPosition = self.win.tableWidget.rowCount()
                    self.win.tableWidget.insertRow(rowPosition)
                    try:
                        if str(sheet.cell_value(0, 3)).upper() == "ACTIVE":
                            try:
                                if str(sheet.cell_value(0, 4)).upper() == "PERIOD":
                                    self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem(str(sheet.cell_value(i, 3))))
                                    self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem(str(sheet.cell_value(i, 4)).strip().split('.')[0]))
                                else:
                                    self.win.tableWidget.setItem(rowPosition, 0,
                                                                 QTableWidgetItem(str(sheet.cell_value(i, 3))))
                                    self.win.tableWidget.setItem(rowPosition, 1,
                                                                 QTableWidgetItem("0"))
                            except Exception as e:
                                self.win.tableWidget.setItem(rowPosition, 0,
                                                             QTableWidgetItem(str(sheet.cell_value(i, 3))))
                                self.win.tableWidget.setItem(rowPosition, 1,
                                                             QTableWidgetItem("0"))
                        else:
                            try:
                                if str(sheet.cell_value(0, 3)).upper() == "PERIOD":
                                    self.win.tableWidget.setItem(rowPosition, 1,
                                                                 QTableWidgetItem("E"))
                                    self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem(
                                        str(sheet.cell_value(i, 4)).strip().split('.')[0]))
                                else:
                                    self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem("E"))
                                    self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem("0"))
                            except:
                                self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem("E"))
                                self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem("0"))
                    except Exception as e:
                        print("Error Exist Flag 1 ", e)
                        self.win.tableWidget.setItem(rowPosition, 0, QTableWidgetItem("E"))
                        self.win.tableWidget.setItem(rowPosition, 1, QTableWidgetItem("0"))
                    for col in range(2, 8):
                        self.win.tableWidget.setItem(rowPosition, col, QTableWidgetItem(row_data[col]))
                    for k in range(8):
                        self.win.tableWidget.item(rowPosition, k).setBackground(QColor(200, 250, 200))
            except Exception as e:
                print("Error at Row ", i - 1, " ", str(e), " :: Function Get File Explorer (3)")
        self.table_header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        self.table_header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        self.table_header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        self.table_header.setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)
        self.table_header.setSectionResizeMode(6, QtWidgets.QHeaderView.ResizeToContents)
        self.table_header.setSectionResizeMode(7, QtWidgets.QHeaderView.ResizeToContents)
        if is_col_miss == 1:
            msgBox = QMessageBox()
            msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
            msgBox.setIcon(QMessageBox.Warning)
            msgBox.setText("There are missing columns that cannot be null in your excel file!")
            msgBox.setWindowTitle("Warning")
            msgBox.exec()
        if self.autoGrayDeleteStatus:
            self.deleteGrayRows()
        self.fileisprevious = 0
        self.win.tableWidget.setSortingEnabled(True)
        try:
            self.win.tableWidget.itemChanged.connect(self.editTrigger)
        except Exception as e:
            print(str(e))
        #self.progress.hide()
    def deleteSelectedRows(self):
        try:
            self.win.tableWidget.itemChanged.disconnect()
        except Exception as e:
            print(str(e))
        try:
            selected = sorted(self.win.tableWidget.selectedIndexes())
            rows = set()
            for index in selected:
                rows.add(index.row())
            for row in sorted(rows, reverse=True):
                self.win.tableWidget.removeRow(row)
        except Exception as e:
            print(str(e), " :: Function Delete Selected Rows")
        try:
            self.win.tableWidget.itemChanged.connect(self.editTrigger)
        except Exception as e:
            print(str(e))
    def fillTable(self):
        if self.win.tableWidget.rowCount() == 0:
            return
        try:
            self.win.tableWidget.itemChanged.disconnect()
        except Exception as e:
            print(str(e))
        error_message = ""
        try:
            con = cx_Oracle.connect('<<<MASKED>>>')
            cur = con.cursor()
            for i in range(self.win.tableWidget.rowCount()):
                try:
                    if self.win.tableWidget.item(i, 5).text() and self.win.tableWidget.item(i,
                                                                                            6).text() and self.win.tableWidget.item(
                            i, 7).text():
                        continue
                except:
                    pass
                try:
                    # ,USERNAME,TABLENAME
                    query = "select WORKFLOWNAME, SESSIONNAME, FOLDER from <<<MASKED>>> where tablename='" + self.win.tableWidget.item(
                        i, 4).text() + "' and USERNAME = '" + self.win.tableWidget.item(i, 3).text() + "'"
                    cur.execute(query)
                    data = cur.fetchone()
                    if not data:
                        continue
                    self.win.tableWidget.item(i, 0).setBackground(QColor(135, 206, 235))
                    self.win.tableWidget.item(i, 1).setBackground(QColor(135, 206, 235))
                    self.win.tableWidget.item(i, 2).setBackground(QColor(135, 206, 235))
                    self.win.tableWidget.item(i, 3).setBackground(QColor(135, 206, 235))
                    self.win.tableWidget.item(i, 4).setBackground(QColor(135, 206, 235))
                    self.win.tableWidget.setItem(i, 5, QTableWidgetItem(str(data[0])))
                    self.win.tableWidget.item(i, 5).setBackground(QColor(135, 206, 235))  # sky blue
                    self.win.tableWidget.setItem(i, 6, QTableWidgetItem(str(data[1])))
                    self.win.tableWidget.item(i, 6).setBackground(QColor(135, 206, 235))
                    self.win.tableWidget.setItem(i, 7, QTableWidgetItem(str(data[2])))
                    self.win.tableWidget.item(i, 7).setBackground(QColor(135, 206, 235))
                except Exception as e:
                    error_message += str(e) + " at row " + str(i)
                    print(str(e), " :: Function Fill Table (1)")
            if error_message:
                msgBox = QMessageBox()
                msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
                msgBox.setIcon(QMessageBox.Critical)
                msgBox.setText(error_message)
                msgBox.setWindowTitle("Error Occured")
                msgBox.exec()
            cur.close()
            con.close()
        except Exception as e:
            print(str(e), " :: Function Fill Table (2)")
        try:
            self.win.tableWidget.itemChanged.connect(self.editTrigger)
        except Exception as e:
            print(str(e))
    def autoFillClicked(self):
        for i in range(self.win.tableWidget.rowCount()):
            try:
                if not self.win.tableWidget.item(i, 2) or not self.win.tableWidget.item(i, 2).text():
                    self.win.tableWidget.setItem(i, 2, QTableWidgetItem(self.win.tableWidget.item(i - 1, 2).text()))
                    self.win.tableWidget.item(i, 2).setBackground(self.win.tableWidget.item(i, 1).background())
            except Exception as e:
                print(str(e), " :: Function Auto Fill Clicked")
    def deleteNullRows(self):
        try:
            self.win.tableWidget.itemChanged.disconnect()
        except Exception as e:
            print(str(e))
        nullRows = []
        for i in range(self.win.tableWidget.rowCount()):
            try:
                # buraya bak
                if not self.win.tableWidget.item(i, 7) or not self.win.tableWidget.item(i,
                                                                                        6) or not self.win.tableWidget.item(
                        i, 5) or not self.win.tableWidget.item(i, 4) or not self.win.tableWidget.item(i,
                                                                                                      3) or not self.win.tableWidget.item(
                        i, 2) or not self.win.tableWidget.item(i, 1) or not self.win.tableWidget.item(i, 0):
                    nullRows.append(i)
                elif not self.win.tableWidget.item(i, 7).text().strip() or not self.win.tableWidget.item(i,
                                                                                                 6).text().strip() or not self.win.tableWidget.item(
                        i, 5).text() or not self.win.tableWidget.item(i, 4).text().strip() or not self.win.tableWidget.item(i,
                                                                                                                    3).text().strip() or not self.win.tableWidget.item(
                        i, 2).text() or not self.win.tableWidget.item(i, 1).text().strip() or not self.win.tableWidget.item(i,
                                                                                                                    0).text().strip():
                    nullRows.append(i)
            except Exception as e:
                print(str(e), " :: Function Delete Null Rows (1)")
        for row in sorted(nullRows, reverse=True):
            try:
                self.win.tableWidget.removeRow(row)
            except Exception as e:
                print(str(e), " :: Function Delete Null Rows (2)")
        try:
            self.win.tableWidget.itemChanged.connect(self.editTrigger)
        except Exception as e:
            print(str(e))
    def deleteGrayRows(self):
        deleted_row_indexes = []
        for i in range(self.win.tableWidget.rowCount()):
            # QColor(150, 150, 150) dark gray
            # QColor(100, 100, 100) light gray
            if self.win.tableWidget.item(i, 0).background() == QColor(150, 150, 150) or self.win.tableWidget.item(i,0).background() == QColor(100, 100, 100):
                deleted_row_indexes.append(i)
        for i in range(len(deleted_row_indexes) - 1, -1, -1):
            self.win.tableWidget.removeRow(deleted_row_indexes[i])
    def acceptButtonClicked(self):
        con = cx_Oracle.connect('<<<MASKED>>>')
        cur = con.cursor()
        self.queries_procedure_call = []
        self.queries_procedure_desc = []
        try:
            self.win.tableWidget.itemChanged.disconnect()
        except Exception as e:
            print(str(e))
        # self.editTrigger()
        procedure_target_table_command_count = 0
        procedure_description_command_count = 0
        procedure_call_command_count = 0
        previous_sp_list = []
        is_missing_cell = 0
        is_active_sense = 0
        try:
            text = ""
            text2 = ""
            for i in range(0, self.win.tableWidget.rowCount()):
                try:
                    ACTIVENESS = self.win.tableWidget.item(i, 0).text().upper()
                    PERIOD = self.win.tableWidget.item(i, 1).text().upper()
                    procedure_name = self.win.tableWidget.item(i, 2).text().upper()
                    USERNAME = self.win.tableWidget.item(i, 3).text().upper()
                    TABLENAME = self.win.tableWidget.item(i, 4).text().upper()
                    WORKFLOWNAME = self.win.tableWidget.item(i, 5).text()
                    SESSIONNAME = self.win.tableWidget.item(i, 6).text()
                    FOLDER = self.win.tableWidget.item(i, 7).text()
                    if not ACTIVENESS or not PERIOD or not procedure_name or not USERNAME or not TABLENAME or not WORKFLOWNAME or not SESSIONNAME or not FOLDER:
                        is_missing_cell = 1
                        for j in range(7):
                            try:
                                self.win.tableWidget.item(i, j).setBackground(QColor(255, 0, 0))
                            except Exception as e:
                                print(str(e))
                        continue
                    elif ACTIVENESS != "E" and ACTIVENESS != "H":
                        is_active_sense = 1
                        for j in range(8):
                            try:
                                self.win.tableWidget.item(i, j).setBackground(QColor(255, 0, 0))
                            except Exception as e:
                                print(str(e))
                        continue
                    session_name = "s_m_" + procedure_name
                    # QColor(135,206,235) blue
                    # QColor(200, 250, 200) green
                    if self.win.tableWidget.item(i, 0).background() == QColor(135, 206, 235):  # if color blue insert
                        temp = "insert into <<<MASKED>>> values (\'" + procedure_name + "\',\'" + WORKFLOWNAME + "\',\'" + SESSIONNAME + "\',sysdate,\'" + FOLDER + "\',\'" + USERNAME + "\',\'" + TABLENAME + "\',\'" + ACTIVENESS + "\'," + PERIOD + ")"
                        self.queries_procedure_call.append(temp)
                        text += temp
                        text += ";\n\n"
                    elif self.win.tableWidget.item(i, 0).background() == QColor(200, 250, 200):  # if color green update
                        temp = "update <<<MASKED>>> set BAGLI_WORKFLOW = '" + WORKFLOWNAME + "', BAGLI_SESSION = '" + SESSIONNAME + "', FOLDER_ADI = '" + FOLDER + "', ACTIVE = '" + ACTIVENESS + "', ETLDATE = sysdate, PERIOD = " + PERIOD + " where PROSEDUR_ISMI = '" + procedure_name + "' and OWNER = '" + USERNAME + "' and TABLE_NAME = '" + TABLENAME + "'"
                        self.queries_procedure_call.append(temp)
                        text += temp
                        text += ";\n\n"
                    else:
                        continue
                    procedure_call_command_count += 1
                    temp = "select * from <<<MASKED>>> where prosedur_ismi = '" + procedure_name + "'"

                    if procedure_name in previous_sp_list or procedure_name in self.global_procedure_list:
                        continue
                    cur.execute(temp)
                    if cur.fetchone():
                        continue
                    temp = "INSERT INTO <<<MASKED>>> (PROSEDUR_ISMI, INFORMATICA_FOLDER_ISMI, WORKFLOW_ISMI, SESSION_ISMI, CALISMA_PERIYODU,HAFTASONU_CALISACAKMI, CALISMA_SAATI,CALISMA_DAKIKASI, HAFTANIN_CALISMA_GUNU, AYIN_CALISMA_GUNU, BLOKE_DURUMU, ETLDATE, EXPIRE_TIME) VALUES ( '" + procedure_name + "','ATOMIC','ATOMIC','" + session_name + "' ,1,'E',0,0,0,0,'A',sysdate,'23:30')"
                    self.queries_procedure_desc.append(temp)
                    text2 += temp
                    text2 += "\n\n"
                    procedure_description_command_count += 1
                    previous_sp_list.append(procedure_name)
                except Exception as e:
                    is_missing_cell = 1
                    for j in range(7):
                        try:
                            self.win.tableWidget.item(i, j).setBackground(QColor(255, 0, 0))
                        except:
                            pass
                    print(str(e), " :: Function Accept Button Clicked (1)")
            cur.close()
            con.close()
            if is_active_sense:
                msgBox = QMessageBox()
                msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
                msgBox.setIcon(QMessageBox.Warning)
                msgBox.setText("Activeness has to be E or H!")
                msgBox.setWindowTitle("Warning")
                msgBox.exec()
            if is_missing_cell:
                msgBox = QMessageBox()
                msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
                msgBox.setIcon(QMessageBox.Warning)
                msgBox.setText("There are missing cells in your table!")
                msgBox.setWindowTitle("Warning")
                msgBox.exec()
            # procedure target table text edit fill
            try:
                query_all = ""
                # sheet 2 read
                wb = xlrd.open_workbook(self.previousFilePath)

                if wb.nsheets >= 2:
                    sheet2 = wb.sheet_by_index(1)
                    if sheet2.row_len(0) != 11:
                        msgBox = QMessageBox()
                        msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
                        msgBox.setIcon(QMessageBox.Critical)
                        msgBox.setText("Excel File Must Have All Columns at Sheet 2")
                        msgBox.setWindowTitle("Error")
                        msgBox.exec()
                    else:
                        con = cx_Oracle.connect('<<<MASKED>>>')
                        cur = con.cursor()
                        for i in range(1, sheet2.nrows):
                            isUpdate = 0
                            if str(sheet2.cell_value(i, 0)) and str(sheet2.cell_value(i, 1)) and str(sheet2.cell_value(i, 2)) and str(sheet2.cell_value(i, 3)):
                                cntrlquery = "select * from <<<MASKED>>> where TABLEOWNER = '" + str(sheet2.cell_value(i, 0)) + "' and TABLENAME = '" + str(sheet2.cell_value(i, 1)) + "' and SPOWNER = '" + str(sheet2.cell_value(i, 2)) + "' and SPNAME = '" + str(sheet2.cell_value(i, 3)) + "'"
                                cur.execute(cntrlquery)
                                if cur.fetchone():
                                    isUpdate = 1
                            if isUpdate == 0:
                                query = "insert into VITDM.DM_MIS_PROCEDURETARGETTABLE values(sysdate,"
                                for j in range(11):
                                    if str(sheet2.cell_value(i, j)):
                                        if j == 7 or j == 9:
                                            query += "'" + str(sheet2.cell_value(i, j)).upper().strip().split('.')[
                                                0] + "',"
                                        else:
                                            query += "'" + str(sheet2.cell_value(i, j)).upper().strip() + "',"
                                    else:
                                        query += "'',"
                                query = query[:-1] + ");"
                                query_all += query + "\n"
                                procedure_target_table_command_count += 1
                            else:
                                query = "update <<<MASKED>>> SET "
                                query += "ETLDATE = sysdate,"
                                query += "MASKEDTABLENAME = '" + str(sheet2.cell_value(i, 4)).upper().strip() + "',"
                                query += "VIEWNAME = '" + str(sheet2.cell_value(i, 5)).upper().strip() + "',"
                                if str(sheet2.cell_value(i, 6)):
                                    query += "RELATEDTM = '" + str(sheet2.cell_value(i, 6)).upper().strip() + "',"
                                if str(sheet2.cell_value(i, 7)):
                                    query += "AUTOMICSCHEDULED = " + str(sheet2.cell_value(i, 7)).upper().strip().split('.')[0] + ","
                                if str(sheet2.cell_value(i, 8)):
                                    query += "DETAIL1 = '" + str(sheet2.cell_value(i, 8)).upper().strip() + "',"
                                if str(sheet2.cell_value(i, 9)):
                                    query += "DAYSONMONTH = '" + str(sheet2.cell_value(i, 9)).upper().strip().split('.')[0] + "',"
                                if str(sheet2.cell_value(i, 10)):
                                    query += "OWNER = '" + str(sheet2.cell_value(i, 10)).upper().strip() + "',"
                                query = query[:-1]
                                query += " WHERE TABLEOWNER = '" + str(sheet2.cell_value(i, 0)).upper().strip() + "' and TABLENAME = '" + str(sheet2.cell_value(i, 1)).upper().strip() + "' and SPOWNER = '" + str(sheet2.cell_value(i, 2)).upper().strip() + "' and SPNAME = '" + str(sheet2.cell_value(i, 3)).upper().strip() + "'"
                                query += ";"
                                query_all += query + "\n"
                                procedure_target_table_command_count += 1
                        cur.close()
                        con.close()
                self.win2 = uic.loadUi(".\\GUI\\final.ui")  # specify the location of your .ui file
                if wb.nsheets < 2:
                    self.win2.label_4.setText("No Procedure Target Table Sheet Exists")
                else:
                    self.win2.label_4.setText("Procedure Target Table " + str(procedure_target_table_command_count) + " rows")
                self.win2.textEdit_3.setText(query_all)
                self.win2.textEdit_2.setText(text)
                self.win2.textEdit.setText(text2)
                # self.win2.show()
                self.win2.showMaximized()
                self.win2.label_2.setText(
                    self.win2.label_2.text() + " " + str(procedure_description_command_count) + " rows")
                self.win2.label_3.setText(
                    self.win2.label_3.text() + " " + str(procedure_call_command_count) + " rows")
                self.win2.pushButton.clicked.connect(self.copyText1)
                self.win2.pushButton_2.clicked.connect(self.copyText2)
                self.win2.pushButton_5.clicked.connect(self.copyText3)
                self.win2.pushButton_3.clicked.connect(self.copyTextAll)
                self.win2.pushButton_4.clicked.connect(self.writeToDB)
            except Exception as e:
                print(str(e))
        except Exception as e:
            print(str(e), " :: Function Accept Button Clicked (2)")
        try:
            self.win.tableWidget.itemChanged.connect(self.editTrigger)
        except Exception as e:
            print(str(e))
    def copyText1(self):
        try:
            cb = QApplication.clipboard()
            cb.clear(mode=cb.Clipboard)
            cb.setText(self.win2.textEdit.toPlainText(), mode=cb.Clipboard)
            print("Procedure Description Text Copied to Clipboard")
            self.win2.label.setText("Procedure Description Copied to Clipboard")
        except Exception as e:
            print(str(e))
    def copyText2(self):
        try:
            cb = QApplication.clipboard()
            cb.clear(mode=cb.Clipboard)
            cb.setText(self.win2.textEdit_2.toPlainText(), mode=cb.Clipboard)
            print("Procedure Call Text Copied to Clipboard")
            self.win2.label.setText("Procedure Call Copied to Clipboard")
        except Exception as e:
            print(str(e))
    # target table
    def copyText3(self):
        try:
            cb = QApplication.clipboard()
            cb.clear(mode=cb.Clipboard)
            cb.setText(self.win2.textEdit_3.toPlainText(), mode=cb.Clipboard)
            print("Procedure Target Table Text Copied to Clipboard")
            self.win2.label.setText("Procedure Target Table Copied to Clipboard")
        except Exception as e:
            print(str(e))
    def copyTextAll(self):
        try:
            cb = QApplication.clipboard()
            cb.clear(mode=cb.Clipboard)
            cb.setText(self.win2.textEdit_3.toPlainText() + "\n\n\n\n\n\n" + self.win2.textEdit.toPlainText() + "\n\n\n\n\n\n" + self.win2.textEdit_2.toPlainText(),
                       mode=cb.Clipboard)
            print("All Text Copied to Clipboard")
            self.win2.label.setText("All Contents Copied to Clipboard")
        except Exception as e:
            print(str(e))
    def writeToDB(self):
        con = cx_Oracle.connect('HADOOPUSER/e4R5ghuc_234@exa-scan:1954/EXAORCX')
        cur = con.cursor()
        error_message = ""
        error_count = 0
        success_count = 0
        is_error = 0
        targettablequeries = self.win2.textEdit_3.toPlainText().split("\n")
        for query in targettablequeries:
            try:
                if query:
                    cur.execute(query[:-1]) # noktalı virgülü siliyor
                    success_count += 1
            except Exception as e:
                is_error = 1
                error_count += 1
                error_message += "Oracle Error\n" + str(e) + "\n\nQuery\n" + query + "\n\n\n"
                print("Error:Procedure Target Table \n", str(e))
                print(query)
        for query in self.queries_procedure_desc:
            try:
                cur.execute(query)
                success_count += 1
            except Exception as e:
                is_error = 1
                error_count += 1
                error_message += str(e) + "\n"
                print("Error:Procedure Description ", str(e))
                print(query)
        for query in self.queries_procedure_call:
            try:
                cur.execute(query)
                success_count += 1
            except Exception as e:
                is_error = 1
                error_count += 1
                error_message += str(e) + "\n"
                print("Error:Procedure Call ", str(e))
                print(query)
        if is_error:
            msgBox = QMessageBox()
            msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
            msgBox.setIcon(QMessageBox.Warning)
            msgBox.setText("There are some failed rows" + " " * 150) # Message Box Resize edemediğim için böyle bir yöntem kullanmak durumunda kaldım
            msgBox.setDetailedText(error_message)
            msgBox.setWindowTitle("Failed Rows Exist")
            msgBox.exec()
        try:
            con.commit()
        except Exception as e:
            msgBox = QMessageBox()
            msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
            msgBox.setIcon(QMessageBox.Warning)
            msgBox.setText(str(e))
            msgBox.setWindowTitle("Commit Failed")
            msgBox.exec()
        msgBox = QMessageBox()
        msgBox.setWindowIcon(QtGui.QIcon('.\\Icon\\chain.png'))
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setText("Successfull Count:\t" + str(success_count) + "\n" + "Failed Count:\t\t" + str(error_count))
        msgBox.setWindowTitle("Write Process Done")
        msgBox.exec()
        cur.close()
        con.close()
    def exitEvent(self, e):
        try:
            with open('Initial.txt', 'w') as file:
                if self.win.autoGrayDeleteCheckBox.isChecked():
                    file.write("AUTO_GRAY_DELETE,1")
                else:
                    file.write("AUTO_GRAY_DELETE,0")
                file.write("\n")
                file.write("PREVIOUS_PATH_FILE," + self.previousFilePath)
        except Exception as e:
            print(str(e))
        elapsed_time = int(time.time() - start_time)
        try:
            con = cx_Oracle.connect('<<<MASKED>>>')
            cur = con.cursor()
            if not self.previousFilePath:
                self.previousFilePath = ""
            cur.execute("insert into <<<MASKED>>>(ACTIVITYDATE, ACTIVITYTYPE, USERNAME, DURATION, FILENAME) values (sysdate,'LOGOUT','" + str(getpass.getuser()) + "','" + str(elapsed_time) + "','"+ self.processedFilePath +"')")
            cur.close()
            con.commit()
            con.close()
        except Exception as e:
            print("WARNING Connection May Cause Failure")
        print("bye bye!")
if __name__ == '__main__':
    start_time = time.time()
    try:
        con = cx_Oracle.connect('<<<MASKED>>>')
        cur = con.cursor()
        cur.execute("insert into <<<MASKED>>>(ACTIVITYDATE, ACTIVITYTYPE, USERNAME) values (sysdate,'LOGIN','" + str(getpass.getuser()) + "')")
        cur.close()
        con.commit()
        con.close()
    except Exception as e:
        print("WARNING Connection May Cause Failure")
    app = QtWidgets.QApplication([])
    window = MainFrame(".\\GUI\\main.ui")
    sys.exit(app.exec_())
