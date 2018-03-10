# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'second.ui'
#
# Created by: PyQt5 UI code generator 5.8.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import rc_rc
import pandas as pd
import time
import Excel
import logging

class Ui_Main_frame(object):
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
        self.listWidget_2 = QtWidgets.QListWidget(self.groupBox)
        self.listWidget_2.setGeometry(QtCore.QRect(20, 70, 441, 431))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(1)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.listWidget_2.sizePolicy().hasHeightForWidth())
        self.listWidget_2.setSizePolicy(sizePolicy)
        self.listWidget_2.setMinimumSize(QtCore.QSize(0, 0))
        self.listWidget_2.setMaximumSize(QtCore.QSize(16777215, 16777215))
        self.listWidget_2.setFrameShape(QtWidgets.QFrame.WinPanel)
        self.listWidget_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.listWidget_2.setLineWidth(2)
        self.listWidget_2.setAlternatingRowColors(True)
        self.listWidget_2.setFlow(QtWidgets.QListView.TopToBottom)
        self.listWidget_2.setObjectName("listWidget_2")

        self.listWidget_2.itemClicked.connect(self.on_change)

        item = QtWidgets.QListWidgetItem()
        item.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsUserCheckable|QtCore.Qt.ItemIsEnabled)
        item.setCheckState(QtCore.Qt.Checked)
        self.listWidget_2.addItem(item)
        item = QtWidgets.QListWidgetItem()
        item.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsUserCheckable|QtCore.Qt.ItemIsEnabled)
        item.setCheckState(QtCore.Qt.Checked)
        self.listWidget_2.addItem(item)
        self.file_name = QtWidgets.QLabel(self.groupBox)
        self.file_name.setGeometry(QtCore.QRect(20, 40, 191, 31))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.file_name.setFont(font)
        self.file_name.setObjectName("file_name")
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
        self.Dm_seq = QtWidgets.QComboBox(self.groupBox_2)
        self.Dm_seq.setGeometry(QtCore.QRect(20, 90, 171, 21))
        self.Dm_seq.setObjectName("Dm_seq")
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
        self.Genome_st = QtWidgets.QComboBox(self.groupBox_2)
        self.Genome_st.setGeometry(QtCore.QRect(20, 190, 171, 24))
        self.Genome_st.setObjectName("Genome_st")
        self.label_3 = QtWidgets.QLabel(self.groupBox_2)
        self.label_3.setGeometry(QtCore.QRect(20, 220, 121, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.Repeat_re = QtWidgets.QComboBox(self.groupBox_2)
        self.Repeat_re.setGeometry(QtCore.QRect(20, 240, 171, 24))
        self.Repeat_re.setObjectName("Repeat_re")
        self.label_4 = QtWidgets.QLabel(self.groupBox_2)
        self.label_4.setGeometry(QtCore.QRect(20, 270, 77, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.ORF = QtWidgets.QComboBox(self.groupBox_2)
        self.ORF.setGeometry(QtCore.QRect(20, 290, 171, 24))
        self.ORF.setObjectName("ORF")
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
        self.Dm_position = QtWidgets.QComboBox(self.groupBox_2)
        self.Dm_position.setGeometry(QtCore.QRect(20, 40, 171, 24))
        self.Dm_position.setObjectName("Dm_position")
        self.label_9 = QtWidgets.QLabel(self.groupBox_2)
        self.label_9.setGeometry(QtCore.QRect(20, 20, 151, 18))
        font = QtGui.QFont()
        font.setPointSize(7)
        self.label_9.setFont(font)
        self.label_9.setObjectName("label_9")
        self.SEQ = QtWidgets.QComboBox(self.groupBox_2)
        self.SEQ.setGeometry(QtCore.QRect(20, 140, 171, 24))
        self.SEQ.setObjectName("SEQ")
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
        self.ORF_edit = QtWidgets.QTextEdit(self.groupBox_3)
        self.ORF_edit.setGeometry(QtCore.QRect(380, 50, 141, 161))
        self.ORF_edit.setObjectName("ORF_edit")
        self.NCR_edit = QtWidgets.QTextEdit(self.groupBox_3)
        self.NCR_edit.setGeometry(QtCore.QRect(550, 50, 141, 161))
        self.NCR_edit.setObjectName("NCR_edit")
        self.groupBox_2.raise_()
        self.Btn_Box.raise_()
        self.groupBox.raise_()
        self.groupBox_3.raise_()

        self.retranslateUi(Main_frame)
        self.listWidget_2.setCurrentRow(-1)
        self.Btn_Box.accepted.connect(Main_frame.accept)
        self.Btn_Box.rejected.connect(Main_frame.reject)
        QtCore.QMetaObject.connectSlotsByName(Main_frame)

        self.xls = None

    def retranslateUi(self, Main_frame):
        _translate = QtCore.QCoreApplication.translate
        Main_frame.setWindowTitle(_translate("Main_frame", "Analyzing Sequence"))
        self.groupBox.setTitle(_translate("Main_frame", "입력 데이터"))
        self.listWidget_2.setSortingEnabled(False)
        __sortingEnabled = self.listWidget_2.isSortingEnabled()
        self.listWidget_2.setSortingEnabled(False)
        item = self.listWidget_2.item(0)
        item.setText(_translate("Main_frame", "sheet1"))
        item = self.listWidget_2.item(1)
        item.setText(_translate("Main_frame", "sheet2"))
        self.listWidget_2.setSortingEnabled(__sortingEnabled)
        self.file_name.setText(_translate("Main_frame", "sample.xlsx"))
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
        from os import getenv
        # get full path about selected file
        filename = QtWidgets.QFileDialog.getOpenFileName(None, 'Open File', 
            getenv('HOME'), "Excel files (*.xlsx)")
        if filename[0] != '':
            # extract only file name from full path
            name_label = filename[0].split('/')[-1]
            self.file_name.setText(QtCore.QCoreApplication.translate("Main_frame", name_label))
            item_num = 0
            
            filename = filename[0]
            self.filename_ = name_label
            self.sheets = list()

            if filename is not None:
                start_time = time.time()
                xls = pd.ExcelFile(filename)
                self.listWidget_2.setSortingEnabled(False)
                self.listWidget_2.clear()
                with xls:
                    for n in xls.sheet_names:
                        self.sheets.append(xls.parse(n))
                        item = self.listWidget_2.item(item_num)
                        if item is None:
                            item = QtWidgets.QListWidgetItem()
                            item.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsUserCheckable|QtCore.Qt.ItemIsEnabled)
                            item.setCheckState(QtCore.Qt.Unchecked)
                            self.listWidget_2.addItem(item)
                        else:
                            item.setFlags(QtCore.Qt.ItemIsSelectable|QtCore.Qt.ItemIsUserCheckable|QtCore.Qt.ItemIsEnabled)
                            item.setCheckState(QtCore.Qt.Unchecked)
                        item.setText(QtCore.QCoreApplication.translate("Main_frame", str(n)))
                        item_num = item_num + 1
                msg = QtWidgets.QMessageBox()
                msg.setIcon(QtWidgets.QMessageBox.Information)   
                msg.setWindowTitle("Info")
                msg.setText("Load Success !")
                msg.setInformativeText("Running Time : {0}".format(int(time.time()-start_time)))
                msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
                msg.exec_()

            else:
                xls = None
            self.set_xlsx(xls)

    def replace_(self, L):
        L = L.replace('\n',',').replace(',,',',').replace(',,',',').replace(',,',',').replace(',,',',').strip(',').split(',')
        if L[-1] == '':
            L.pop()
        if L[0] == '':
            L.pop(0)
        return L

    def Analyzing(self):
        if self.xls is not None:
            try:
                logging.basicConfig(filename='./GPA_log.log',level=logging.DEBUG)
                selected_sheets = list()
                for index in range(self.listWidget_2.count()):
                    if self.listWidget_2.item(index).checkState() == QtCore.Qt.Checked:
                        selected_sheets.append(self.listWidget_2.item(index).text())
                        Dplist = self.Dm_position.currentText()
                        Dslist = self.Dm_seq.currentText()
                        Dgelist = self.Genome_st.currentText()
                        DRelist = self.Repeat_re.currentText()
                        DORF = self.ORF.currentText()
                        Dseq = self.SEQ.currentText()

                Repeat_edit_ = str(self.Repeat_edit.toPlainText())
                Repeat_edit_ = self.replace_(Repeat_edit_)

                ORF_edit_ = str(self.ORF_edit.toPlainText())
                ORF_edit_ = self.replace_(ORF_edit_)

                NCR_edit_ = str(self.NCR_edit.toPlainText())
                NCR_edit_ = self.replace_(NCR_edit_)

                Genome_edit_ = str(self.Genome_edit.toPlainText())
                Genome_edit_ = self.replace_(Genome_edit_)

                start_time = time.time()
                logging.info("{0} Start Initialization of Data".format(time.time()))
                excel = Excel.Excel(self.xls, self.filename_, selected_sheets, 
                    Dplist, Dslist, Dgelist, DRelist, DORF, Dseq, 
                    Genome_edit_, Repeat_edit_, ORF_edit_, NCR_edit_)
                
                if self.radio_full.isChecked():
                    Analyze_type = "Full"
                    logging.info("{0} Full Sequence Analyzation".format(time.time()))
                else:
                    Analyze_type = "Difference_of_Minor"
                    logging.info("{0} Difference_of_Minor Analyzation".format(time.time()))

                # Start Ananlyzation
                logging.info("{0} Start Analyzation".format(time.time()))
                self.Analyze_Dialog(excel.Analyze(Analyze_type), start_time)

            except Exception as e:
                logging.error("{0} {1}".format(time.time(), e))
                
    def Analyze_Dialog(self, Result, start_time):
        msg = QtWidgets.QMessageBox()
        msg.setIcon(QtWidgets.QMessageBox.Information)   
        msg.setWindowTitle("Info")
        msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        if Result == "Success":
            msg.setText("Success !")
            msg.setInformativeText("Running Time : {0}".format(int(time.time()-start_time)))
            logging.error("{0} Success to finish the analyzation".format(time.time()))
        elif Result == "Error":
            msg.setText("Error!\n컬럼명 일치 여부, 입력 누락, 컬럼 선택 등 기타 주의사항 확인")
            msg.setInformativeText("ls123kr@naver.com 문의")
            logging.error("{0} Etc Error".format(time.time()))
        elif Result == "Permission":
            msg.setText("Permission Error!\n생성하려는 파일과 동일한 이름의 파일이 열려있습니다.")
            logging.error("{0} File Permission Error".format(time.time()))
        msg.exec_()

    def on_change(self):
        curitem = self.listWidget_2.currentItem()
        idx = self.listWidget_2.currentRow()
        if curitem == None:
            return
        isselect = QtCore.Qt.Unchecked
        if curitem.checkState() == QtCore.Qt.Checked:
            isselect = QtCore.Qt.Unchecked 
        else:
            isselect = QtCore.Qt.Checked
            if self.xls is not None:
                cursheet = self.sheets[idx]
                if self.Dm_position.count() == 0:
                    self.Dm_position.addItem('선택안함')
                    self.Dm_seq.addItem('선택안함')
                    self.Genome_st.addItem('선택안함')
                    self.Repeat_re.addItem('선택안함')
                    self.ORF.addItem('선택안함')
                    self.SEQ.addItem('선택안함')
                    for col in cursheet.columns:
                        self.Dm_position.addItem(col)
                        self.Dm_seq.addItem(col)
                        self.Genome_st.addItem(col)
                        self.Repeat_re.addItem(col)
                        self.ORF.addItem(col)
                        self.SEQ.addItem(col)
        curitem.setCheckState(isselect)
        curitem.setSelected(False)

    def set_xlsx(self, xls):
        self.xls = xls

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Main_frame = QtWidgets.QDialog()
    ui = Ui_Main_frame()
    ui.setupUi(Main_frame)
    Main_frame.show()
    sys.exit(app.exec_())

