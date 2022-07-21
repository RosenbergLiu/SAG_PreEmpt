from cProfile import label
from lib2to3.pgen2 import driver
from typing import OrderedDict
from unittest import result
from PyQt6.QtWidgets import QApplication, QHBoxLayout, QProgressBar, QVBoxLayout, QWidget,QLabel, QLineEdit, QPushButton,QListWidget, QMessageBox
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt
import os
import time
from PyQt6.QtCore import QThread, pyqtSignal
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from os.path import exists
import json
import chrome_version 
import psycopg2 as db_connect

class Thread(QThread):
    _signal = pyqtSignal(int)
    def __init__(self):
        super(Thread,self).__init__()
    def __del__(self):
        self.wait()
    def run(self):
        for i in range(100):
            time.sleep(0.1)
            self._signal.emit(i)

class Window(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pre-empt")
        
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)


        # Job Number
        self.job_layout = QVBoxLayout()
        self.label1 = QLabel("Job Numer")
        self.job_layout.addWidget(self.label1)
        self.input1 = QLineEdit()
        self.job_layout.addWidget(self.input1)
        self.layout.addLayout(self.job_layout)


        # Part Number
        label2 = QLabel("Part Numer & Quantity")
        self.layout.addWidget(label2)

        self.part_layout = QHBoxLayout()
        self.input2 = QLineEdit("1")
        self.input2.setFixedWidth(30)
        self.part_layout.addWidget(self.input2)
        label_x = QLabel("x")
        self.part_layout.addWidget(label_x)
        self.input3 = QLineEdit()
        self.part_layout.addWidget(self.input3)

        self.layout.addLayout(self.part_layout)


        button_add = QPushButton("Add")
        button_add.setFixedWidth(40)
        button_del = QPushButton("Delete")
        button_del.setFixedWidth(50)
        self.layout_button = QHBoxLayout()
        self.layout_button.addWidget(button_add)
        self.layout_button.addWidget(button_del)
        self.layout_button.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.layout.addLayout(self.layout_button)

        self.part_list = QListWidget()

        self.layout.addWidget(self.part_list)


        button_add.clicked.connect(self.onAddClicked)
        button_del.clicked.connect(self.onDelClicked)


        self.button_run = QPushButton("Run")
        self.layout.addWidget(self.button_run)

        self.button_run.clicked.connect(self.onRunClicked)


        #Result
        self.result = QLineEdit()
        self.layout.addWidget(self.result)


        #Progress Bar
        self.PgBar = QProgressBar(self)
        self.layout.addWidget(self.PgBar)


    def onAddClicked(self):
        part_number = self.input3.text().replace(" ","")
        quantity = self.input2.text().replace(" ","")
        if part_number=="":
            dlg = QMessageBox(self)
            dlg.setWindowTitle("Error")
            dlg.setText("No part number entered")
            dlg.setIcon(QMessageBox.Icon.Warning)
            dlg.exec()
        elif len(part_number)!=6 or part_number.isnumeric() == False:
            dlg = QMessageBox(self)
            dlg.setWindowTitle("Error")
            dlg.setText("Invalid number entered")
            dlg.setIcon(QMessageBox.Icon.Warning)
            dlg.exec()
            self.input3.clear()
            self.input3.setFocus()
        else:
            record = part_number + "   X   " + quantity
            self.part_list.addItem(record)
            self.input3.clear()
            self.input2.setText("1")
            self.input3.setFocus()

    def onDelClicked(self):
        selected = self.part_list.currentRow()
        self.part_list.takeItem(selected)


    def onRunClicked(self):
        self.thread = Thread()
        self.thread._signal.connect(self.signal_accept)
        self.thread.start()
        if self.part_list.count() == 0:
            dlg = QMessageBox(self)
            dlg.setWindowTitle("Error")
            dlg.setText("No part added")
            dlg.setIcon(QMessageBox.Icon.Warning)
            dlg.exec()
        elif self.input1.text()=="":
            dlg = QMessageBox(self)
            dlg.setWindowTitle("Error")
            dlg.setText("Invalid job number")
            dlg.setIcon(QMessageBox.Icon.Warning)
            dlg.exec()
        else:
            with open('config.json','r') as j:
                config= json.load(j)
                usr= config['usr']
                pwd = config['pwd']
            jobNum = self.input1.text()
            driver = webdriver.Chrome(executable_path="C:/Users/roshan.liu/Scripts/SAG_PreEmpt/chromedriver.exe")
            driver.get("https://partners.gorenje.com/sagcc/Sredina.aspx")
            wait = WebDriverWait(driver, 10)
            driver.find_element(By.ID, "usr").send_keys(usr)
            driver.find_element(By.ID, "pwd").send_keys(pwd)
            driver.find_element(By.ID, "btnPrijava").click()
            original_window = driver.current_window_handle
            assert len(driver.window_handles) == 1
            driver.find_element(By.ID, "ctl00_tbOss").send_keys(jobNum)
            driver.find_element(By.ID, "ctl00_btnOss").click()
            wait.until(EC.number_of_windows_to_be(2))
            for window_handle in driver.window_handles:
                if window_handle != original_window:
                    driver.switch_to.window(window_handle)
                    break
            wait.until(EC.presence_of_element_located((By.ID,"ctl00_ContentPlaceHolder1_btnAUNarocilo")))
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnAUNarocilo").click()
            driver.switch_to.frame('ctl00_ContentPlaceHolder1_ASPxPopupControl1_CIF-1') 
            preemptParts = {}
            for row in range(self.part_list.count()):

                rowitem=self.part_list.item(row).text().split("   X   ")
                part=rowitem[0]
                qty=rowitem[1]
                driver.find_element(By.ID, "dd_izd_sifra_I").clear()
                driver.find_element(By.ID, "dd_izd_sifra_I").send_keys(part)
                driver.find_element(By.ID, "dd_izd_sifra_I").send_keys(Keys.ENTER)
                wait.until(EC.presence_of_element_located((By.ID,"dd_izd_sifra_DDD_L_LBI0T0")))
                driver.find_element(By.ID, "dd_izd_sifra_I").send_keys(Keys.DOWN)

                driver.find_element(By.ID, "dd_izd_sifra_I").send_keys(Keys.ENTER)

                #driver.find_element(By.CSS_SELECTOR, '#dd_izd_sifra_DDD_L_LBI0T0 > em').click()
                #driver.find_element(By.ID, "dd_izd_sifra_DDD_PWC-1").send_keys(Keys.ENTER)



                wait.until(EC.presence_of_element_located((By.ID,"spKolicinaMatNarocilo_I")))
                driver.find_element(By.ID, "spKolicinaMatNarocilo_I").clear()
                driver.find_element(By.ID, "spKolicinaMatNarocilo_I").send_keys(qty)
                driver.find_element(By.ID, "btnDodajNaNarocilo").click()
            driver.find_element(By.ID, "btnNarociloZapriInZakljuci_CD").click()
            alert = wait.until(EC.alert_is_present())
            ALE=alert.text
            alert.accept()
            
            SAP = ALE.split(": ")[1]
            self.result.setText(SAP)
            self.result.selectAll()
            self.part_list.clear()
            self.input1.clear()



            os.system("Scripts\\SAP_PreEmpt\\ReleaseOneTransfer.vbs {}".format(SAP))





    def signal_accept(self, msg):
        self.PgBar.setValue(int(msg))
        if self.PgBar.value() == 99:
            self.PgBar.setValue(0)


    #def onCredClicked(self):
        #cred_window = credWindow()
        #cred_window.show()

"""
class credWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pre-empt")
        
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        with open('config.json','r') as j:
            config = json.load(j)
            usr = config['usr']
            pwd = config['pwd']
        self.labelUsr = QLabel('User Name:')
        self.inputUsr = QLineEdit(usr)
        self.labelPwd = QLabel('Password')
        self.inputPwd = QLineEdit(pwd)
        self.button_sav = QPushButton
        self.button_sav.clicked.connect(self.onSavClicked)


        self.layout.addWidget(self.labelUsr)
        self.layout.addWidget(self.inputUsr)
        self.layout.addWidget(self.labelPwd)
        self.layout.addWidget(self.inputPwd)
        self.layout.addWidget(self.button_sav)

    def onSavClicked(self):
        usr = self.inputUsr.text()
        pwd = self.inputPwd.text()
        json_obj={
            'usr':usr,
            'pwd':pwd
        }
        with open('config.json','w') as j:
            json.dump(json_obj,j)

"""











print(chrome_version.get_chrome_version())

exist=exists('config.json')
if not exist:
    json_obj={
            'usr':'',
            'pwd':''
        }
    with open('config.json','w') as j:
        json.dump(json_obj,j)

app=QApplication([])
window=Window()
window.show()
app.exec()

