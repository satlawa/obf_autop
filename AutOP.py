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
        self.width = 700
        self.height = 450

        self.height_titles = 60
        self.height_main = 100
        self.height_buttons = 400

        self.data = 0
        self.file_check = np.zeros((10,), dtype=int)
        self.to_check = np.zeros((5,), dtype=int)

        self.initUI()


    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        #--------------#
        #--   path   --#
        #--------------#

        x_path = 40
        # Create label Path
        self.label_path = QLabel(self)
        self.label_path.setText("Pfad:")
        self.label_path.move(x_path, self.height_main - 60)
        # Create textbox Path
        self.textbox_path = QLineEdit(self)
        self.textbox_path.move(x_path+60, self.height_main - 55)
        self.textbox_path.resize(500, 20)
        self.textbox_path.setText(os.path.join(os.getcwd(), 'data'))

        #----------------#
        #--   labels   --#
        #----------------#

        # set positon on x-axis
        x_label = 40
        # title
        self.label_dir = QLabel(self)
        self.label_dir.setText("Informationen")
        self.label_dir.move(x_label, self.height_main - 30)

        # Create label TO
        self.label_to = QLabel(self)
        self.label_to.setText("TO:")
        self.label_to.move(x_label,self.height_main)

        # Create label old TO
        self.label_to_alt = QLabel(self)
        self.label_to_alt.setText("altes TO:")
        self.label_to_alt.move(x_label,self.height_main + 30)

        self.label_ks1 = QLabel(self)
        self.label_ks1.setText("Klimastation:")
        self.label_ks1.move(x_label,self.height_main + 30*2)

        self.label_ks2 = QLabel(self)
        self.label_ks2.setText("Klimastation:")
        self.label_ks2.move(x_label,self.height_main + 30*3)

        self.label_ks3 = QLabel(self)
        self.label_ks3.setText("Klimastation:")
        self.label_ks3.move(x_label,self.height_main + 30*4)

        self.label_fb = QLabel(self)
        self.label_fb.move(x_label,self.height_main + 30*5)

        self.label_fr = QLabel(self)
        self.label_fr.move(x_label,self.height_main + 30*6)

        self.label_ks3 = QLabel(self)
        self.label_ks3.setText("Status:")
        self.label_ks3.move(x_label,self.height_main + 30*7)

        self.label_info = QLabel(self)
        self.label_info.setText("need to check")
        self.label_info.resize(200,20)
        self.label_info.move(x_label,self.height_main + 30*8)

        #-------------------#
        #--   textboxes   --#
        #-------------------#

        x_textbox = 130
        # Create textbox TO_new
        self.textbox_to_new = QLineEdit(self)
        self.textbox_to_new.move(x_textbox, self.height_main + 5)
        self.textbox_to_new.resize(80, 20)

        # Create textbox TO_old
        self.textbox_to_old = QLineEdit(self)
        self.textbox_to_old.move(x_textbox, self.height_main + 35)
        self.textbox_to_old.resize(80, 20)

        # Create textbox TO_Klimastation
        self.textbox_klimastation1 = QLineEdit(self)
        self.textbox_klimastation1.move(x_textbox, self.height_main + 65)
        self.textbox_klimastation1.resize(80, 20)

        # Create textbox TO_Klimastation
        self.textbox_klimastation2 = QLineEdit(self)
        self.textbox_klimastation2.move(x_textbox, self.height_main + 95)
        self.textbox_klimastation2.resize(80, 20)

        # Create textbox TO_Klimastation
        self.textbox_klimastation3 = QLineEdit(self)
        self.textbox_klimastation3.move(x_textbox, self.height_main + 125)
        self.textbox_klimastation3.resize(80, 20)

        #----------------------------#
        #--   checkboxes - files   --#
        #----------------------------#

        # set positon on x-axis
        x_chkbox_files = 200
        # title
        self.label_dir = QLabel(self)
        self.label_dir.setText("Dateien")
        self.label_dir.move(x_chkbox_files+40, self.height_main - 30)

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

        checkbes_f = [self.label1, self.label2, self.label3, self.label4, self.label5, self.label6, self.label7, self.label8, self.label9, self.label10]
        # loop though the checkboxses
        for i, checkb in enumerate(checkbes_f):
            checkb.move(x_chkbox_files+150, self.height_main + i*20)

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

        checkbes_f = [self.checkbox1, self.checkbox2, self.checkbox3, self.checkbox4, self.checkbox5, self.checkbox6, self.checkbox7, self.checkbox8, self.checkbox9, self.checkbox10]
        # loop though the checkboxses
        for i, checkb in enumerate(checkbes_f):
            checkb.move(x_chkbox_files+40, self.height_main + i*20 - 5)
            checkb.resize(320,40)

        #-------------------------------#
        #--   checkboxes - chapters   --#
        #-------------------------------#


        # set positon on x-axis
        x_chkbox_chp = 400
        # title
        self.label_dir = QLabel(self)
        self.label_dir.setText("Kapitel")
        self.label_dir.move(x_chkbox_chp+40, self.height_main - 30)

        self.labelc1 = QLabel(self)
        self.labelc2 = QLabel(self)
        self.labelc3 = QLabel(self)
        self.labelc4 = QLabel(self)
        self.labelc5 = QLabel(self)
        self.labelc6 = QLabel(self)
        self.labelc7 = QLabel(self)
        self.labelc8 = QLabel(self)
        self.labelc9 = QLabel(self)
        self.labelc10 = QLabel(self)
        self.labelc11 = QLabel(self)
        self.labelc12 = QLabel(self)

        checkbes_c = [self.labelc1, self.labelc2, self.labelc3, self.labelc4, self.labelc5, self.labelc6, self.labelc7, self.labelc8, self.labelc9, self.labelc10, self.labelc11, self.labelc12]
        # loop though the checkboxses
        for i, checkb in enumerate(checkbes_c):
            checkb.move(x_chkbox_chp, self.height_main + i*20)

        self.checkboxc1 = QCheckBox("01 - Allgemeines",self)
        self.checkboxc2 = QCheckBox("02 - Hauptergebnisse",self)
        self.checkboxc3 = QCheckBox("03 - Natuerliche Grundlagen",self)
        self.checkboxc4 = QCheckBox("04 - Besitzstand",self)
        self.checkboxc5 = QCheckBox("05 - Wald",self)
        self.checkboxc6 = QCheckBox("06 - Holzernte",self)
        self.checkboxc7 = QCheckBox("07 - Waldpflege",self)
        self.checkboxc8 = QCheckBox("08 - Einforstung",self)
        self.checkboxc9 = QCheckBox("09 - Wirtschaftsbeschraenkungen",self)
        self.checkboxc10 = QCheckBox("10 - Naturschutz",self)
        self.checkboxc11 = QCheckBox("11 - Vormerkungen",self)
        self.checkboxc12 = QCheckBox("12 - Anhang",self)

        self.checkbes_c = [self.checkboxc1, self.checkboxc2, self.checkboxc3, self.checkboxc4, self.checkboxc5, self.checkboxc6, self.checkboxc7, self.checkboxc8, self.checkboxc9, self.checkboxc10, self.checkboxc11, self.checkboxc12]
        # loop though the checkboxses
        for i, checkb in enumerate(self.checkbes_c):
            checkb.move(x_chkbox_chp+40, self.height_main + i*20 - 5)
            checkb.setCheckState(QtCore.Qt.Checked)
            checkb.resize(320,40)

        #-----------------#
        #--   buttons   --#
        #-----------------#

        # Create a button in the window
        self.button_check = QPushButton('check', self)
        self.button_check.move(40,self.height_buttons)

        # Create a button in the window
        self.button_run = QPushButton('run', self)
        self.button_run.move(550,self.height_buttons)

        # connect button to function on_click
        self.button_check.clicked.connect(self.on_click_check)
        self.button_run.clicked.connect(self.on_click_run)
        self.show()


    ##################
    ###   checks   ###
    ##################

    def check_files(self, to_now):
        ### check if all files are avalible
        checkbes = [[self.checkbox1, '_sap.XLS'], [self.checkbox2, '_sap_natur.XLS'], [self.checkbox3, '_hs_bilanz.XLS'], [self.checkbox4, '_hs_bilanz_old.XLS'], [self.checkbox5, '_bw_alter.xlsx'], [self.checkbox6, '_bw_neig.xlsx'], [self.checkbox7, '_bw_seeh.xlsx'], [self.checkbox8, '_bw_uz.xlsx'], [self.checkbox9, '_bw_zufaellige.xlsx'], [self.checkbox10, 'SPI_2018.txt']]
        # loop though the checkboxses
        for i, checkb in enumerate(checkbes):
            file = os.path.isfile(os.path.join(self.path_dir, 'TO' + to_now, 'to_' + to_now + checkb[1]))
            if file == True:
                checkb[0].setCheckState( QtCore.Qt.Checked )
                self.file_check[i] = 1
            else:
                checkb[0].setCheckState( QtCore.Qt.Unchecked )
                self.file_check[i] = 0
        # check SPI
        file = os.path.isfile(os.path.join(self.path_dir, 'stichprobe', 'SPI_' + self.year + '.txt'))
        if file == True:
            self.checkbox10.setCheckState( QtCore.Qt.Checked )
            self.file_check[9] = 1
        else:
            self.checkbox10.setCheckState( QtCore.Qt.Unchecked )
            self.file_check[9] = 0


    def check_to_sap_roh(self, to_now):

        path = os.path.join(self.path_dir, 'TO' + to_now, 'to_' + to_now + '_sap.XLS')
        file = os.path.isfile(path)

        if file == True:

            ### check if all files are avalible
            data = pd.read_csv(os.path.join(path), sep='\t', encoding = "ISO-8859-1", decimal=',', error_bad_lines=False)

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
                self.label1.setText("fail")
                self.to_check[0] = 0
            #self.label4.setText("FR: " + str(self.data['Forstrevier'].unique()))

            # set year
            self.year = str(int(data.loc[0,'Beg. Laufzeit'][-4:])-1)

            print(self.year)

            return True

        else:
            return False

    def check_to_sap_natur(self, to_now):
        ### check if all files are avalible
        data = pd.read_csv(os.path.join(self.path_dir, 'TO' + to_now, 'to_' + to_now + '_sap_natur.XLS'), sep='\t', encoding = "ISO-8859-1", decimal=',', error_bad_lines=False)
        to = data['Teiloperats-ID'].unique()[0]
        if int(to_now) == to:
            self.label2.setText("ok")
            self.to_check[1] = 1
        else:
            self.label2.setText("fail")
            self.to_check[1] = 0

    def check_to_hs_bilanz(self, to_now):
        ### check if all files are avalible
        data = pd.read_csv(os.path.join(self.path_dir, 'TO' + to_now, 'to_' + to_now + '_hs_bilanz.XLS'), sep='\t', encoding = "ISO-8859-1", decimal=',', error_bad_lines=False)
        to = data['Teiloperats-ID'].unique()[0]
        if int(to_now) == to:
            self.label3.setText("ok")
            self.to_check[2] = 1
        else:
            self.label3.setText("fail")
            self.to_check[2] = 0

    def check_to_hs_bilanz_old(self, to_now, to_old):
        ### check if all files are avalible
        data = pd.read_csv(os.path.join(self.path_dir, 'TO' + to_now, 'to_' + to_now + '_hs_bilanz_old.XLS'), sep='\t', encoding = "ISO-8859-1", decimal=',', error_bad_lines=False)
        to = data['Teiloperats-ID'].unique()
        if int(to_old) in to:
            self.label4.setText("ok")
            self.to_check[3] = 1
        else:
            self.label4.setText("fail")
            self.to_check[3] = 0

    def check_to_spi(self, to_now):
        ### check if all files are avalible
        data = pd.read_csv(os.path.join(self.path_dir, 'stichprobe', 'SPI_' + self.year + '.txt'), sep='\t', encoding = "ISO-8859-1", decimal=',', thousands='.', skipinitialspace=True, skiprows=3)
        to = data['TOID'].unique()
        if int(to_now) in to:
            self.label10.setText("ok")
            self.to_check[4] = 1
        else:
            self.label10.setText("fail")
            self.to_check[4] = 0

    def parse_to(self, to_text):
        # parse teiloperats-ids from text to int
        to_old = to_text.strip().split(',')
        #
        try:
            return list(np.array(to_old, dtype=np.int16))
        except Exception as e:
            raise


    @pyqtSlot()
    def on_click_check(self):
        self.path_dir = self.textbox_path.text()
        to_now = self.textbox_to_new.text()
        to_old = self.textbox_to_old.text()
        #self.year = self.textbox_laufzeit.text()[-4:]

        print(self.path_dir)
        print(to_now)
        print(to_old)

        if to_now == '':
            to_now = '0'
        if to_old == '':
            to_old = '0'

        main_check = self.check_to_sap_roh(to_now)

        if main_check:
            self.check_files(to_now)

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

        else:
            self.label_fb.setText("")
            self.label_fr.setText("")
            self.label1.setText("")
            self.to_check[1] = 0


    ### run the machine!!!
    def on_click_run(self):

        print(self.to_check)

        if np.array_equal(self.to_check, np.ones(5)):

            #-----------------------------------------------------#
            # read data
            self.label_info.setText('making Operat...')
            to_now = self.textbox_to_new.text()
            to_old = self.textbox_to_old.text() #[1069] # [1031, 1041]
            to_old = self.parse_to(to_old)

            #-----------------------------------------------------#
            # add klima stations
            klimast1 = self.textbox_klimastation1.text()
            klimast2 = self.textbox_klimastation2.text()
            klimast3 = self.textbox_klimastation3.text()
            klima_stationen = list()
            for klimast in [klimast1, klimast2, klimast3]:
                if klimast != "":
                    klima_stationen.append(klimast)

            #-----------------------------------------------------#
            # check which sections should to be created
            do_sections = dict()
            sections = [
                "1_Allgemeines",
                "2_Hauptergebnisse",
                "3_Natuerliche_Grundlagen",
                "4_Besitzstand",
                "5_Wald",
                "6_Holzernte",
                "7_Waldpflege",
                "8_Einforstung",
                "9_Wirtschaftsbeschraenkungen",
                "10_Naturschutz",
                "11_Vormerkungen",
                "12_Anhang",
            ]

            for i, checkb in enumerate(self.checkbes_c):
                do = 1 if checkb.isChecked() else 0
                do_sections[sections[i]] = do

            #-----------------------------------------------------#
            # run script
            obf_main = OBFMain(self.path_dir, to_now, to_old, klima_stationen, do_sections)
            info = obf_main.run()
            self.label_info.setText(info)

        else:
            self.label_info.setText('not able to run')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
