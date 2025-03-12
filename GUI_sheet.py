# -*- coding: utf-8 -*-
"""
Created on Sun Feb 28 11:21:14 2021

@authors: 童丹佛, danfo_tong@hotmail.com
@投资人：童伟
"""

from PyQt5.QtWidgets import QWidget,QPushButton,QApplication,QLabel,QFileDialog,QLineEdit,QComboBox,QRadioButton,QGroupBox,QProgressBar,QMessageBox,QSpinBox
from PyQt5 import QtCore
import sys
from sheet_wos_for_GUI import *
class Screen(QWidget):
    def __init__(self):
        super(Screen,self).__init__()
        # self.setObjectName("爸爸是SB")
        self.move(300, 400)
        self.length = 500
        self.height = 300
    def setup(self):
        # self.setWindowTitle('爸爸是SB')
        self.setWindowTitle('文章信息自动采集和填写 1.0')
        self.setFixedSize(self.length,self.height)
        #self.setStyleSheet("#爸爸是SB{background-color:blue}")
                                
        text1 = '''  这个程序是可以根据excel表中文章的标题（C列，2行开始）自动填写相关信息。
  1.点击“start wos”从webofscience获取信息。（下方的数值框表示从第几个标题开始）
  2.点击“恢复”，自动从上次中断处开始。
  3.点击“start fengqu”从分区表网站获取信息。（下方的数值框表示从第几个标题开始）
  4.运行时不要点击excel表中的cell。
    注：“start fengqu”开始以前，请先完成wos的填写。
    contact:  danfo_tong@hotmail.com  Tong Danfo
          or  weitong@hmfl.ac.cn      Tong Wei'''
        self.what = QLabel(self)
        self.what.setGeometry(0, 0, self.length, self.height)
        self.what.setWordWrap(True)
        self.what.setAlignment(QtCore.Qt.AlignTop)
        self.what.setText(text1)

        self.button_start_wos = QPushButton('开始\nwos',self)
        self.button_start_wos.setGeometry(50, self.height-70, 50, 50)
        self.button_start_wos.clicked.connect(self.start_wos)

        self.button_start_fengqu = QPushButton('开始\nfengqu',self)
        self.button_start_fengqu.setGeometry(self.length-100, self.height-70, 50, 50)
        self.button_start_fengqu.clicked.connect(self.start_fengqu)

        self.location_save = QLineEdit('',self)
        self.location_save.setGeometry(10, self.height-175, 360, 30)
        self.location_save.setReadOnly(True)

        self.sheet = QComboBox(self)
        self.sheet.setGeometry(self.length-130,self.height-175,80,30)
        self.sheet.activated.connect(self.sheet_set)

        self.button_location = QPushButton('浏览', self)
        self.button_location.setGeometry(self.length-50, self.height-180, 40, 40)
        self.button_location.clicked.connect(self.open_)

        self.button_reset = QPushButton('恢复',self)
        self.button_reset.setGeometry(self.length-275, self.height-70, 50, 50)
        self.button_reset.clicked.connect(self.start_reset)

        self.pbar_lable = QLabel(self)
        self.pbar_lable.setObjectName("pbar_lable")
        self.pbar_lable.setGeometry(10, self.height-134, 150, 18)
        self.pbar_lable.setStyleSheet("#pbar_lable{background-color: white;border: 1px solid black}")
        self.pbar_lable.setVisible(False)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(10, self.height-115, 480, 15)
        self.progress_bar.setVisible(False)

        self.sbox_wos = QSpinBox(self)
        self.sbox_wos.setGeometry(50, self.height-20, 50, 20)

        self.sbox_fengqubiao = QSpinBox(self)
        self.sbox_fengqubiao.setGeometry(self.length-100, self.height-20, 50, 20)

        #self.button_test = QPushButton('test',self)
        #self.button_test.clicked.connect(self.test)

        # self.button.setStyleSheet("QPushButton{border-image: url(img/exit.png)}")
        # self.button.setStyleSheet("QPushButton:hover{border-image: url(img/exit_hover.png)}\QPushButton{border-image: url(img/exit.png)}")
        
    #def test(self):
        #self.what.setText('11111')
    def open_(self):
        self.location = QFileDialog.getOpenFileName(self,"请选择文件...",'./',('Excel (*.xlsx)')) 
        if self.location == ():
            return None
        self.location_save.setText(self.location[0])
        app_ = xlwings.App(visible=True,add_book=False)
        book = app_.books.open(r'%s'%(self.location[0]))
        self.sheet.clear()
        self.sheet.addItems([sht.name for sht in book.sheets])
        book.app.quit()
        self.sheet_set()
    def sheet_set(self):
        self.sbox_wos.setMinimum(1)
        self.sbox_fengqubiao.setMinimum(1)
        app_ = xlwings.App(visible=True,add_book=False)
        book = app_.books.open(r'%s'%(self.location_save.text()))
        sht = book.sheets[self.sheet.currentText()]
        self.sbox_wos.setMaximum(len(sht.range(2,3).expand('down').value))
        self.sbox_fengqubiao.setMaximum(len(sht.range(2,3).expand('down').value))
        book.app.quit()
    def start_wos(self):
        if self.location_save.text() == '':
            QMessageBox.information(self,'提示','请先选择地址。')
            return 
        self.button_start_fengqu.setEnabled(False)
        self.button_reset.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.pbar_lable.setVisible(True)
        #filling_sheet(self.location_save.text(),self.sheet.currentText())
        for x in filling_sheet(self.location_save.text(),self.sheet.currentText(),restart='wos',start_at=self.sbox_wos.value()):
            self.progress_bar.setValue(int(x[0]/x[1]*100))
            self.pbar_lable.setText(x[2]+':'+str(x[0])+'/'+str(x[1]))
            QApplication.processEvents()
        self.button_reset.setEnabled(True)
        self.button_start_fengqu.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.pbar_lable.setVisible(False)
        #print(self.location_save.text(),self.sheet.currentText())
    def start_fengqu(self):
        if self.location_save.text() == '':
            QMessageBox.information(self,'提示','请先选择地址。')
            return 
        self.button_start_wos.setEnabled(False)
        self.button_reset.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.pbar_lable.setVisible(True)
        #filling_sheet(self.location_save.text(),self.sheet.currentText())
        for x in filling_sheet(self.location_save.text(),self.sheet.currentText(),restart='fengqubiao',start_at=self.sbox_fengqubiao.value()):
            self.progress_bar.setValue(int(x[0]/x[1]*100))
            self.pbar_lable.setText(x[2]+':'+str(x[0])+'/'+str(x[1]))
            QApplication.processEvents()
        self.button_reset.setEnabled(True)
        self.button_start_wos.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.pbar_lable.setVisible(False)
        #print(self.location_save.text(),self.sheet.currentText())
    def start_reset(self):
        if self.location_save.text() == '':
            QMessageBox.information(self,'提示','请先选择地址。')
            return 
        self.button_start_fengqu.setEnabled(False)
        self.button_start_wos.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.pbar_lable.setVisible(True)
        #filling_sheet(self.location_save.text(),self.sheet.currentText())
        for x in filling_sheet(self.location_save.text(),self.sheet.currentText()):
            self.progress_bar.setValue(int(x[0]/x[1]*100))
            self.pbar_lable.setText(x[2]+':'+str(x[0])+'/'+str(x[1]))
            QApplication.processEvents()
        self.button_start_fengqu.setEnabled(True)
        self.button_start_wos.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.pbar_lable.setVisible(False)
        #print(self.location_save.text(),self.sheet.currentText())


# a = Screen()
# a.setup()
# a.show()
if __name__ == '__main__':
    app = QApplication(sys.argv)
    a = Screen()
    a.setup()
    a.show()
    sys.exit(app.exec_())
    
    

