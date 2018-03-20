# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'second.ui'
#
# Created by: PyQt5 UI code generator 5.8.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from time import time
import rc_rc
import pandas as pd
import logging
from book import Book
from analyzer import Analyzer

class Ui_Main_frame(object):

    book = None

    def setupUi(self, Main_frame):
        Main_frame.setObjectName("Main_frame")
        Main_frame.setWindowModality(QtCore.Qt.NonModal)
        Main_frame.setEnabled(True)
        Main_frame.resize(757, 785)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Ignored, QtWidgets.QSizePolicy.Ignored)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Main_frame.sizePolicy().hasHeightForWidth())
        Main_frame.setSizePolicy(sizePolicy)
        Main_frame.setMinimumSize(QtCore.QSize(600, 453))
        Main_frame.setMaximumSize(QtCore.QSize(1200, 900))
        Main_frame.setFocusPolicy(QtCore.Qt.ClickFocus)
        Main_frame.setContextMenuPolicy(QtCore.Qt.NoContextMenu)
        Main_frame.setAcceptDrops(False)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/dna.jpg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        Main_frame.setWindowIcon(icon)
        Main_frame.setToolTipDuration(0)
        Main_frame.setWhatsThis("")
        Main_frame.setAutoFillBackground(False)
        Main_frame.setSizeGripEnabled(False)
        Main_frame.setModal(False)
        self.Btn_Box = QtWidgets.QDialogButtonBox(Main_frame)
        self.Btn_Box.setGeometry(QtCore.QRect(560, 30, 121, 81))
        self.Btn_Box.setOrientation(QtCore.Qt.Vertical)
        self.Btn_Box.setStandardButtons(QtWidgets.QDialogButtonBox.Close|QtWidgets.QDialogButtonBox.Ok)
        self.Btn_Box.setObjectName("Btn_Box")

        self.Btn_Box.accepted.connect(self.Analyzing)

        self.groupBox = QtWidgets.QGroupBox(Main_frame)
        self.groupBox.setEnabled(True)
        self.groupBox.setGeometry(QtCore.QRect(20, 20, 481, 521))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.groupBox.setFont(font)
        self.groupBox.setFlat(False)
        self.groupBox.setObjectName("groupBox")
        self.sheetlist_Widget = QtWidgets.QListWidget(self.groupBox)
        self.sheetlist_Widget.setGeometry(QtCore.QRect(20, 70, 441, 431))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(1)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.sheetlist_Widget.sizePolicy().hasHeightForWidth())
        self.sheetlist_Widget.setSizePolicy(sizePolicy)
        self.sheetlist_Widget.setMinimumSize(QtCore.QSize(0, 0))
        self.sheetlist_Widget.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.sheetlist_Widget.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.sheetlist_Widget.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.sheetlist_Widget.setLineWidth(2)
        self.sheetlist_Widget.setAlternatingRowColors(True)
        self.sheetlist_Widget.setFlow(QtWidgets.QListView.TopToBottom)
        self.sheetlist_Widget.setObjectName("sheetlist_Widget")

        self.sheetlist_Widget.itemClicked.connect(self.on_change)

        item = QtWidgets.QListWidgetItem()
        item.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsUserCheckable|QtCore.Qt.ItemIsEnabled)
        item.setCheckState(QtCore.Qt.Checked)
        self.sheetlist_Widget.addItem(item)
        item = QtWidgets.QListWidgetItem()
        item.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsUserCheckable|QtCore.Qt.ItemIsEnabled)
        item.setCheckState(QtCore.Qt.Checked)
        self.sheetlist_Widget.addItem(item)
        self.filename_label = QtWidgets.QLabel(self.groupBox)
        self.filename_label.setGeometry(QtCore.QRect(20, 40, 191, 31))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.filename_label.setFont(font)
        self.filename_label.setObjectName("filename_label")
        self.Load_btn = QtWidgets.QPushButton(self.groupBox)
        self.Load_btn.setGeometry(QtCore.QRect(310, 40, 151, 31))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.Load_btn.setFont(font)
        self.Load_btn.setText("파일 불러오기")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/load2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.Load_btn.setIcon(icon1)
        self.Load_btn.setAutoDefault(False)
        self.Load_btn.setDefault(False)
        self.Load_btn.setFlat(True)
        self.Load_btn.setObjectName("Load_btn")

        self.Load_btn.clicked.connect(self.getfile)

        self.groupBox_2 = QtWidgets.QGroupBox(Main_frame)
        self.groupBox_2.setGeometry(QtCore.QRect(520, 120, 211, 421))
        self.groupBox_2.setObjectName("groupBox_2")
        self.Dumaseq_Combobox = QtWidgets.QComboBox(self.groupBox_2)
        self.Dumaseq_Combobox.setGeometry(QtCore.QRect(20, 90, 171, 21))
        self.Dumaseq_Combobox.setObjectName("Dumaseq_Combobox")
        self.label = QtWidgets.QLabel(self.groupBox_2)
        self.label.setGeometry(QtCore.QRect(20, 70, 141, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.groupBox_2)
        self.label_2.setGeometry(QtCore.QRect(20, 170, 151, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.GenomeST_combobox = QtWidgets.QComboBox(self.groupBox_2)
        self.GenomeST_combobox.setGeometry(QtCore.QRect(20, 190, 171, 24))
        self.GenomeST_combobox.setObjectName("GenomeST_combobox")
        self.label_3 = QtWidgets.QLabel(self.groupBox_2)
        self.label_3.setGeometry(QtCore.QRect(20, 220, 121, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.RepeatReG_combobox = QtWidgets.QComboBox(self.groupBox_2)
        self.RepeatReG_combobox.setGeometry(QtCore.QRect(20, 240, 171, 24))
        self.RepeatReG_combobox.setObjectName("RepeatReG_combobox")
        self.label_4 = QtWidgets.QLabel(self.groupBox_2)
        self.label_4.setGeometry(QtCore.QRect(20, 270, 77, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.ORF_combobox = QtWidgets.QComboBox(self.groupBox_2)
        self.ORF_combobox.setGeometry(QtCore.QRect(20, 290, 171, 24))
        self.ORF_combobox.setObjectName("ORF_combobox")
        self.label_5 = QtWidgets.QLabel(self.groupBox_2)
        self.label_5.setGeometry(QtCore.QRect(20, 330, 181, 16))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(0, 0, 0))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        self.label_5.setPalette(palette)
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.label_5.setFont(font)
        self.label_5.setAutoFillBackground(False)
        self.label_5.setObjectName("label_5")
        self.DumaPos_Combobox = QtWidgets.QComboBox(self.groupBox_2)
        self.DumaPos_Combobox.setGeometry(QtCore.QRect(20, 40, 171, 24))
        self.DumaPos_Combobox.setObjectName("DumaPos_Combobox")
        self.label_9 = QtWidgets.QLabel(self.groupBox_2)
        self.label_9.setGeometry(QtCore.QRect(20, 20, 151, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label_9.setFont(font)
        self.label_9.setObjectName("label_9")
        self.seq_combobox = QtWidgets.QComboBox(self.groupBox_2)
        self.seq_combobox.setGeometry(QtCore.QRect(20, 140, 171, 24))
        self.seq_combobox.setObjectName("seq_combobox")
        self.label_10 = QtWidgets.QLabel(self.groupBox_2)
        self.label_10.setGeometry(QtCore.QRect(20, 120, 121, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label_10.setFont(font)
        self.label_10.setObjectName("label_10")
        self.radio_full = QtWidgets.QRadioButton(self.groupBox_2)
        self.radio_full.setGeometry(QtCore.QRect(30, 390, 161, 22))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.radio_full.setFont(font)
        self.radio_full.setChecked(True)
        self.radio_full.setObjectName("radio_full")
        self.radio_minor = QtWidgets.QRadioButton(self.groupBox_2)
        self.radio_minor.setGeometry(QtCore.QRect(30, 360, 151, 22))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.radio_minor.setFont(font)
        self.radio_minor.setChecked(False)
        self.radio_minor.setObjectName("radio_minor")
        self.groupBox_3 = QtWidgets.QGroupBox(Main_frame)
        self.groupBox_3.setGeometry(QtCore.QRect(20, 550, 711, 221))
        self.groupBox_3.setObjectName("groupBox_3")
        self.Genome_edit = QtWidgets.QTextEdit(self.groupBox_3)
        self.Genome_edit.setGeometry(QtCore.QRect(20, 50, 151, 161))
        self.Genome_edit.setObjectName("Genome_edit")
        self.label_11 = QtWidgets.QLabel(self.groupBox_3)
        self.label_11.setGeometry(QtCore.QRect(30, 30, 151, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label_11.setFont(font)
        self.label_11.setObjectName("label_11")
        self.label_12 = QtWidgets.QLabel(self.groupBox_3)
        self.label_12.setGeometry(QtCore.QRect(210, 30, 111, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label_12.setFont(font)
        self.label_12.setObjectName("label_12")
        self.Repeat_edit = QtWidgets.QTextEdit(self.groupBox_3)
        self.Repeat_edit.setGeometry(QtCore.QRect(200, 50, 141, 161))
        self.Repeat_edit.setObjectName("Repeat_edit")
        self.label_13 = QtWidgets.QLabel(self.groupBox_3)
        self.label_13.setGeometry(QtCore.QRect(390, 30, 77, 18))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.label_13.setFont(font)
        self.label_13.setObjectName("label_13")
        self.label_14 = QtWidgets.QLabel(self.groupBox_3)
        self.label_14.setGeometry(QtCore.QRect(560, 30, 77, 18))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.label_14.setFont(font)
        self.label_14.setObjectName("label_14")
        self.ORF_combobox_edit = QtWidgets.QTextEdit(self.groupBox_3)
        self.ORF_combobox_edit.setGeometry(QtCore.QRect(380, 50, 141, 161))
        self.ORF_combobox_edit.setObjectName("ORF_combobox_edit")
        self.NCR_edit = QtWidgets.QTextEdit(self.groupBox_3)
        self.NCR_edit.setGeometry(QtCore.QRect(550, 50, 141, 161))
        self.NCR_edit.setObjectName("NCR_edit")
        self.groupBox_2.raise_()
        self.Btn_Box.raise_()
        self.groupBox.raise_()
        self.groupBox_3.raise_()

        self.xls = None
        self.retranslateUi(Main_frame)
        self.sheetlist_Widget.setCurrentRow(-1)
        self.Btn_Box.accepted.connect(Main_frame.accept)
        self.Btn_Box.rejected.connect(Main_frame.reject)
        QtCore.QMetaObject.connectSlotsByName(Main_frame)

        

    def retranslateUi(self, Main_frame):
        _translate = QtCore.QCoreApplication.translate
        Main_frame.setWindowTitle(_translate("Main_frame", "Analyzing Sequence"))
        self.groupBox.setTitle(_translate("Main_frame", "입력 데이터"))
        self.sheetlist_Widget.setSortingEnabled(False)
        __sortingEnabled = self.sheetlist_Widget.isSortingEnabled()
        self.sheetlist_Widget.setSortingEnabled(False)
        item = self.sheetlist_Widget.item(0)
        item.setText(_translate("Main_frame", "sheet1"))
        item = self.sheetlist_Widget.item(1)
        item.setText(_translate("Main_frame", "sheet2"))
        self.sheetlist_Widget.setSortingEnabled(__sortingEnabled)
        self.filename_label.setText(_translate("Main_frame", "sample.xlsx"))
        self.groupBox_2.setTitle(_translate("Main_frame", "Column 명 설정"))
        self.label.setText(_translate("Main_frame", "Duma Seqeunce"))
        self.label_2.setText(_translate("Main_frame", "Genome Structure"))
        self.label_3.setText(_translate("Main_frame", "Repeat Region"))
        self.label_4.setText(_translate("Main_frame", "ORF"))
        self.label_5.setText(_translate("Main_frame", " ※전처리 선택"))
        self.label_9.setText(_translate("Main_frame", "Duma Position"))
        self.label_10.setText(_translate("Main_frame", "Sequence"))
        self.radio_minor.setText(_translate("Main_frame", "5% 이상 변화 추출"))
        self.radio_full.setText(_translate("Main_frame", "Full Sequence"))
        self.groupBox_3.setTitle(_translate("Main_frame", "Column 값 설정"))
        self.Genome_edit.setHtml(_translate("Main_frame", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'Gulim\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">TRS,</p>\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">UL,</p>\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">...</p></body></html>"))
        self.label_11.setText(_translate("Main_frame", "Genome Structure"))
        self.label_12.setText(_translate("Main_frame", "Repeat Region"))
        self.Repeat_edit.setHtml(_translate("Main_frame", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'Gulim\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p></body></html>"))
        self.label_13.setText(_translate("Main_frame", "ORF"))
        self.label_14.setText(_translate("Main_frame", "NCR"))

    def getfile(self):
        # get full path of selected file
        # filename = (filepath, filetype)
        from os import getenv
        filename = QtWidgets.QFileDialog.getOpenFileName(None, 'Open File', 
            getenv('HOME'), "Excel files (*.xlsx)")[0]

        if filename is '':
            # logging
            return
        self.book = Book()
        # extract only file name from full path
        self.book.filename = filename.split('/')[-1]
        self.filename_label.setText(QtCore.QCoreApplication.translate("Main_frame", self.book.filename))

        self.sheets = []
        self.sheetlist_Widget.clear()

        t1 = time()
        # read file
        with pd.ExcelFile(filename) as xls:
            self.book.xls = xls
            for name in xls.sheet_names:
                self.sheets.append(xls.parse(name)) # parse는 분석 시작 후로 미루어야 할 듯

                # add item(sheet) into Widgetlist
                item = QtWidgets.QListWidgetItem()
                item.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsUserCheckable|QtCore.Qt.ItemIsEnabled)
                item.setCheckState(QtCore.Qt.Unchecked)
                self.sheetlist_Widget.addItem(item)
                item.setText(QtCore.QCoreApplication.translate("Main_frame", str(name)))
            
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)   
        msg.setWindowTitle("Info")
        msg.setText("Load Success !")
        msg.setInformativeText("Time : {0}".format(int(time()-t1)))
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        msg.exec_()            


    def replace_(self, L):
        L = L.replace('\n',',').replace(',,',',').replace(',,',',').replace(',,',',').replace(',,',',').strip(',').split(',')
        if L[-1] == '':
            L.pop()
        if L[0] == '':
            L.pop(0)
        return L

    def Analyzing(self):
        if self.book.xls is not None:
        # try:
            logging.basicConfig(filename='./GPA_log.log',level=logging.DEBUG)
            
            for index in range(self.sheetlist_Widget.count()):
                if self.sheetlist_Widget.item(index).checkState() == QtCore.Qt.Checked:
                    self.book.sheet_list.append(self.sheetlist_Widget.item(index).text())
                    
            self.book.col_DumaPosition = self.DumaPos_Combobox.currentText()
            self.book.col_DumaSeq = self.Dumaseq_Combobox.currentText()
            self.book.col_Sequence = self.seq_combobox.currentText()
            self.book.col_GenomeStructure = self.GenomeST_combobox.currentText()
            self.book.col_RepeatRegion = self.RepeatReG_combobox.currentText()
            self.book.col_ORF = self.ORF_combobox.currentText()

            self.book.GenomeStructure = self.replace_(str(self.Genome_edit.toPlainText()))
            self.book.RepeatRegion = self.replace_(str(self.Repeat_edit.toPlainText()))
            self.book.ORF = self.replace_(str(self.ORF_combobox_edit.toPlainText()))
            self.book.NCR = self.replace_(str(self.NCR_edit.toPlainText()))
            
            start_time = time()
            
            logging.info("{0} Start Initialization of Data".format(time()))
            excel = Analyzer(self.book)
            if self.radio_full.isChecked():
                Analyze_type = "Full"
                logging.info("{0} Start Analyzing of full sequence ".format(time()))
            else:
                Analyze_type = "Difference_of_Minor"
                logging.info("{0} Start analyzing of difference_of_Minor Analyzation".format(time()))

            # Start Analyzing
            logging.info("{0} Start Analyzation".format(time()))
            self.Analyze_Dialog(excel.Analyze(Analyze_type, excel.book), start_time)

        # except Exception as e:
        #     logging.error("{0} {1}".format(time(), e))
                
    def Analyze_Dialog(self, Result, start_time):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)   
        msg.setWindowTitle("Info")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        # if Result == "Success":
        #     msg.setText("Success !")
        #     msg.setInformativeText("Running Time : {0}".format(int(time()-start_time)))
        #     logging.info("{0} Success to finish the analyzation".format(time()))
        # elif Result == "Error":
        #     msg.setText("Error!\n컬럼명 일치 여부, 입력 누락, 컬럼 선택 등 기타 주의사항 확인")
        #     msg.setInformativeText("ls123kr@naver.com 문의")
        #     logging.error("{0} Etc Error".format(time()))
        # elif Result == "Permission":
        #     msg.setText("Permission Error!\n생성하려는 파일과 동일한 이름의 파일이 열려있습니다.")
        #     logging.error("{0} File Permission Error".format(time()))
        msg.exec_()

    def on_change(self):
        curitem = self.sheetlist_Widget.currentItem()
        idx = self.sheetlist_Widget.currentRow()
        if curitem == None:
            return
        isselect = None        
        if curitem.checkState() == QtCore.Qt.Checked:
            isselect = QtCore.Qt.Unchecked
        else:
            isselect = QtCore.Qt.Checked
            if self.book.xls is not None:
                cursheet = self.sheets[idx]
                if self.DumaPos_Combobox.count() == 0:
                    self.DumaPos_Combobox.addItem('선택안함')
                    self.Dumaseq_Combobox.addItem('선택안함')
                    self.GenomeST_combobox.addItem('선택안함')
                    self.RepeatReG_combobox.addItem('선택안함')
                    self.ORF_combobox.addItem('선택안함')
                    self.seq_combobox.addItem('선택안함')
                    for col in cursheet.columns:
                        self.DumaPos_Combobox.addItem(col)
                        self.Dumaseq_Combobox.addItem(col)
                        self.GenomeST_combobox.addItem(col)
                        self.RepeatReG_combobox.addItem(col)
                        self.ORF_combobox.addItem(col)
                        self.seq_combobox.addItem(col)
        curitem.setCheckState(isselect)
        curitem.setSelected(False)

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Main_frame = QtWidgets.QDialog()
    ui = Ui_Main_frame()
    ui.setupUi(Main_frame)
    Main_frame.show()
    sys.exit(app.exec_())

