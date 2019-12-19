import sys
import os
import numpy as np
import pandas as pd

from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QAction, QLineEdit, QMessageBox, QLabel,  QCheckBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot, QSize

from OBFMain import OBFMain

class App(QMainWindow):

    def __init__(self):
        super().__init__()
        self.title = 'AutOP v4.0'
        self.left = 10
        self.top = 10
        self.width = 400
        self.height = 250

        self.data = 0
        self.file_check = np.zeros((10,), dtype=int)
        self.to_check = np.zeros((5,), dtype=int)

        self.initUI()


    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # Create label #1
        self.label_to = QLabel(self)
        self.label_to.setText("TO:")
        self.label_to.move(20,10)

        # Create label #2
        self.label_to_alt = QLabel(self)
        self.label_to_alt.setText("alt TO:")
        self.label_to_alt.move(20,40)

        self.label_lz_alt = QLabel(self)
        self.label_lz_alt.setText("Laufzeit:")
        self.label_lz_alt.move(20,70)

        self.label_fb = QLabel(self)
        self.label_fb.move(20,100)

        self.label_fr = QLabel(self)
        self.label_fr.move(20,130)

        self.label_info = QLabel(self)
        self.label_info.setText("need to check")
        self.label_info.resize(200,20)
        self.label_info.move(20,160)

        # Create textbox #1
        self.textbox1 = QLineEdit(self)
        self.textbox1.move(80, 10)
        self.textbox1.resize(80,20)

        # Create textbox #2
        self.textbox2 = QLineEdit(self)
        self.textbox2.move(80, 40)
        self.textbox2.resize(80,20)

        # Create textbox #2
        self.textbox3 = QLineEdit(self)
        self.textbox3.move(80, 70)
        self.textbox3.resize(80,20)

#dat_fuc = True
#dat_es = True
#dat_es_hs = True
#dat_zv_ze = True
#dat_kima = False
#dat_wpplan = False
#dat_natur = True
        self.label1 = QLabel(self)
        self.label2 = QLabel(self)
        self.label3 = QLabel(self)
        self.label4 = QLabel(self)
        self.label5 = QLabel(self)
        self.label6 = QLabel(self)
        self.label7 = QLabel(self)
        self.label8 = QLabel(self)
        self.label9 = QLabel(self)
        self.label10 = QLabel(self)

        checkbes = [self.label1, self.label2, self.label3, self.label4, self.label5, self.label6, self.label7, self.label8, self.label9, self.label10]
        # loop though the checkboxses
        for i, checkb in enumerate(checkbes):
            checkb.move(180,10 + i*20)

        self.checkbox1 = QCheckBox("SAP roh",self)
        self.checkbox2 = QCheckBox("SAP natur",self)
        self.checkbox3 = QCheckBox("HS-Bilanz neu",self)
        self.checkbox4 = QCheckBox("HS-Bilanz alt",self)
        self.checkbox5 = QCheckBox("BW alter",self)
        self.checkbox6 = QCheckBox("BW neig",self)
        self.checkbox7 = QCheckBox("BW seeh",self)
        self.checkbox8 = QCheckBox("BW uz",self)
        self.checkbox9 = QCheckBox("BW zufaellige",self)
        self.checkbox10 = QCheckBox("SPI",self)

        checkbes = [self.checkbox1, self.checkbox2, self.checkbox3, self.checkbox4, self.checkbox5, self.checkbox6, self.checkbox7, self.checkbox8, self.checkbox9, self.checkbox10]
        # loop though the checkboxses
        for i, checkb in enumerate(checkbes):
            checkb.move(220,10 + i*20)
            checkb.resize(320,40)

        # Create a button in the window
        self.button_check = QPushButton('check', self)
        self.button_check.move(20,190)

        # Create a button in the window
        self.button_run = QPushButton('run', self)
        self.button_run.move(20,210)

        # connect button to function on_click
        self.button_check.clicked.connect(self.on_click_check)
        self.button_run.clicked.connect(self.on_click_run)
        self.show()


    def check_files(self, to_now):
        ### check if all files are avalible
        checkbes = [[self.checkbox1, '_sap.XLS'], [self.checkbox2, '_sap_natur.XLS'], [self.checkbox3, '_hs_bilanz.XLS'], [self.checkbox4, '_hs_bilanz_old.XLS'], [self.checkbox5, '_bw_alter.xlsx'], [self.checkbox6, '_bw_neig.xlsx'], [self.checkbox7, '_bw_seeh.xlsx'], [self.checkbox8, '_bw_uz.xlsx'], [self.checkbox9, '_bw_zufaellige.xlsx'], [self.checkbox10, 'SPI_2018.txt']]
        # loop though the checkboxses
        for i, checkb in enumerate(checkbes):
            file = os.path.isfile(os.path.join(os.getcwd(), 'data', 'TO' + to_now, 'to_' + to_now + checkb[1]))
            if file == True:
                checkb[0].setCheckState( QtCore.Qt.Checked )
                self.file_check[i] = 1
            else:
                checkb[0].setCheckState( QtCore.Qt.Unchecked )
                self.file_check[i] = 0
        # check SPI
        file = os.path.isfile(os.path.join(os.getcwd(), 'data', 'stichprobe', 'SPI_' + self.year + '.txt'))
        if file == True:
            self.checkbox10.setCheckState( QtCore.Qt.Checked )
            self.file_check[9] = 1
        else:
            self.checkbox10.setCheckState( QtCore.Qt.Unchecked )
            self.file_check[9] = 0


    def check_to_sap_roh(self, to_now):
        ### check if all files are avalible
        data = pd.read_csv(os.path.join(os.getcwd(), 'data', 'TO' + to_now, 'to_' + to_now + '_sap.XLS'), sep='\t', encoding = "ISO-8859-1", decimal=',', error_bad_lines=False)

        fb = data['Forstbetrieb'].unique()[0]
        self.label_fb.setText("FB: " + str(fb))

        fre = data['Forstrevier'].unique()
        fre_str = ''
        for fr in fre:
            fre_str = fre_str + str(fr) + ', '
        self.label_fr.setText("FR: " + str(fre_str))

        to = data['Teiloperats-ID'].unique()[0]

        # check if input-to and XLS-to are the same
        if int(to_now) == to:
            self.label1.setText("ok")
            self.to_check[0] = 1
        else:
            self.label1.setText("no")
            self.to_check[0] = 0
        #self.label4.setText("FR: " + str(self.data['Forstrevier'].unique()))

    def check_to_sap_natur(self, to_now):
        ### check if all files are avalible
        data = pd.read_csv(os.path.join(os.getcwd(), 'data', 'TO' + to_now, 'to_' + to_now + '_sap_natur.XLS'), sep='\t', encoding = "ISO-8859-1", decimal=',', error_bad_lines=False)
        to = data['Teiloperats-ID'].unique()[0]
        if int(to_now) == to:
            self.label2.setText("ok")
            self.to_check[1] = 1
        else:
            self.label2.setText("no")
            self.to_check[1] = 0

    def check_to_hs_bilanz(self, to_now):
        ### check if all files are avalible
        data = pd.read_csv(os.path.join(os.getcwd(), 'data', 'TO' + to_now, 'to_' + to_now + '_hs_bilanz.XLS'), sep='\t', encoding = "ISO-8859-1", decimal=',', error_bad_lines=False)
        to = data['Teiloperats-ID'].unique()[0]
        if int(to_now) == to:
            self.label3.setText("ok")
            self.to_check[2] = 1
        else:
            self.label3.setText("no")
            self.to_check[2] = 0

    def check_to_hs_bilanz_old(self, to_now, to_old):
        ### check if all files are avalible
        data = pd.read_csv(os.path.join(os.getcwd(), 'data', 'TO' + to_now, 'to_' + to_now + '_hs_bilanz_old.XLS'), sep='\t', encoding = "ISO-8859-1", decimal=',', error_bad_lines=False)
        to = data['Teiloperats-ID'].unique()
        if int(to_old) in to:
            self.label4.setText("ok")
            self.to_check[3] = 1
        else:
            self.label4.setText("no")
            self.to_check[3] = 0

    def check_to_spi(self, to_now):
        ### check if all files are avalible
        data = pd.read_csv(os.path.join(os.getcwd(), 'data', 'stichprobe', 'SPI_' + self.year + '.txt'), sep='\t', encoding = "ISO-8859-1", decimal=',', thousands='.', skipinitialspace=True, skiprows=3)
        to = data['TOID'].unique()
        if int(to_now) in to:
            self.label10.setText("ok")
            self.to_check[4] = 1
        else:
            self.label10.setText("no")
            self.to_check[4] = 0


    @pyqtSlot()
    def on_click_check(self):
        to_now = self.textbox1.text()
        to_old = self.textbox2.text()
        self.year = self.textbox3.text()[-4:]

        if to_now == '':
            to_now = '0'
        if to_old == '':
            to_old = '0'

        self.check_files(to_now)

        ### get FB & FR for TO
        if self.checkbox1.isChecked():
            self.check_to_sap_roh(to_now)
        else:
            self.label_fb.setText("")
            self.label_fr.setText("")
            self.label1.setText("")
            self.to_check[1] = 0

        if self.checkbox2.isChecked():
            self.check_to_sap_natur(to_now)
        else:
            self.label2.setText("")
            self.to_check[1] = 0

        if self.checkbox3.isChecked():
            self.check_to_hs_bilanz(to_now)
        else:
            self.label3.setText("")
            self.to_check[2] = 0

        if self.checkbox4.isChecked():
            self.check_to_hs_bilanz_old(to_now, to_old)
        else:
            self.label4.setText("")
            self.to_check[3] = 0

        if self.checkbox10.isChecked():
            self.check_to_spi(to_now)
        else:
            self.label10.setText("")
            self.to_check[4] = 0

        if np.array_equal(self.to_check, np.ones(5)):
            self.label_info.setText("checked")
        else:
            self.label_info.setText("check unsuccessful")


    ### run the machine!!!
    def on_click_run(self):

        print(self.to_check)

        if np.array_equal(self.to_check, np.ones(5)):

            self.label_info.setText('making Operat...')
            to_now = self.textbox1.text()
            to_old = self.textbox2.text() #[1069] # [1031, 1041]
            to_old = [int(to_old)]
            laufzeit_old = self.textbox3.text() # '2008-2017'
            obf_main = OBFMain(to_now, to_old, laufzeit_old, self.file_check)
            info = obf_main.run()
            self.label_info.setText(info)

        else:
            self.label_info.setText('not able to run')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
