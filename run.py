# -*- coding: utf-8 -*-
from PyQt5 import *
import configparser
from guimain import Ui_MainWindow
from PyQt5 import QtWidgets
from PyQt5 import QtCore
import sys


class UI(object):
    def __init__(self):
        QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
        self.app = QtWidgets.QApplication(sys.argv)
        self.window = QtWidgets.QMainWindow()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self.window)
        self.win_adjust()
        self.window.show()
        self.others()
        sys.exit(self.app.exec_())

    def others(self):
        self.component_connect()
        self.load_config()

    def win_adjust(self):
        import win32api, win32con
        # self.window.resize(win32api.GetSystemMetrics(win32con.SM_CXSCREEN) / 2,
        #                    win32api.GetSystemMetrics(win32con.SM_CYSCREEN) / 2)

    def component_connect(self):
        self.ui.GopushButton.clicked.connect(self.login)

    def load_config(self):
        # 配置用户
        userconfig = configparser.ConfigParser()
        userconfig.read('配置文件\\user.ini')
        userconfig_dict = userconfig.defaults()
        self.user_name = userconfig_dict['user_name']
        self.ui.namelineEdit.setText(self.user_name)
        if userconfig_dict['remember'] == 'True':
            self.password = userconfig_dict['password']
            self.ui.pwdlineEdit.setText(self.password)
            self.ui.usercheckBox.setChecked(True)
        else:
            self.ui.usercheckBox.setChecked(False)

        # 配置筛选行业关键词
        istryconfig = configparser.ConfigParser()
        istryconfig.read('配置文件\\Istry.ini')
        istryconfig_dict = istryconfig.defaults()
        if istryconfig_dict['remember'] == 'True':
            self.istry = istryconfig_dict['istry']
            self.ui.industrylineEdit.setPlainText(self.istry)
            self.ui.industrycheckBox.setChecked(True)
        else:
            self.ui.industrycheckBox.setChecked(False)

        # 配置筛选关键词
        keyconfig = configparser.ConfigParser()
        keyconfig.read('配置文件\\Key.ini')
        keyconfig_dict = keyconfig.defaults()
        if keyconfig_dict['remember'] == 'True':
            self.key = keyconfig_dict['key']
            self.ui.keylineEdit.setPlainText(self.key)
            self.ui.keycheckBox.setChecked(True)
        else:
            self.ui.keycheckBox.setChecked(False)

        # 配置清除关键词
        clear_keyconfig = configparser.ConfigParser()
        clear_keyconfig.read('配置文件\\clear_Key.ini')
        clear_keyconfig_dict = clear_keyconfig.defaults()
        if clear_keyconfig_dict['remember'] == 'True':
            self.clear_key = clear_keyconfig_dict['clear_key']
            self.ui.clear_listslineEdit.setPlainText(self.clear_key)
            self.ui.clear_listscheckBox.setChecked(True)
        else:
            self.ui.clear_listscheckBox.setChecked(False)
    def login(self):
        # 配置用户
        self.user_name = self.ui.namelineEdit.text()
        self.password = self.ui.pwdlineEdit.text()
        userconfig = configparser.ConfigParser()
        if self.ui.usercheckBox.isChecked():
            userconfig["DEFAULT"] = {
                "user_name": self.user_name,
                "password": self.password,
                "remember": self.ui.usercheckBox.isChecked()
            }
        else:
            userconfig["DEFAULT"] = {
                "user_name": self.user_name,
                "password": "",
                "remember": self.ui.usercheckBox.isChecked()
            }
        with open('配置文件\\user.ini', 'w') as userconfigfile:
            userconfig.write(userconfigfile)
        # 配置筛选行业关键词
        self.istry = self.ui.industrylineEdit.toPlainText()
        istryconfig = configparser.ConfigParser()
        if self.ui.industrycheckBox.isChecked():
            istryconfig["DEFAULT"] = {
                "istry": self.istry,
                "remember": self.ui.industrycheckBox.isChecked()
            }
        else:
            istryconfig["DEFAULT"] = {
                "istry": "",
                "remember": self.ui.industrycheckBox.isChecked()
            }
        with open('配置文件\\Istry.ini', 'w') as istryconfigfile:
            istryconfig.write(istryconfigfile)

        # 配置筛选关键词
        self.key = self.ui.keylineEdit.toPlainText()
        keyconfig = configparser.ConfigParser()
        if self.ui.keycheckBox.isChecked():
            keyconfig["DEFAULT"] = {
                "key": self.key,
                "remember": self.ui.keycheckBox.isChecked()
            }
        else:
            keyconfig["DEFAULT"] = {
                "key": "",
                "remember": self.ui.keycheckBox.isChecked()
            }
        with open('配置文件\\Key.ini', 'w') as keyconfigfile:
            keyconfig.write(keyconfigfile)

        # 配置清除关键词
        self.clear_key = self.ui.clear_listslineEdit.toPlainText()
        clear_keyconfig = configparser.ConfigParser()
        if self.ui.clear_listscheckBox.isChecked():
            clear_keyconfig["DEFAULT"] = {
                "clear_key": self.clear_key,
                "remember": self.ui.clear_listscheckBox.isChecked()
            }
        else:
            clear_keyconfig["DEFAULT"] = {
                "clear_key": "",
                "remember": self.ui.clear_listscheckBox.isChecked()
            }
        with open('配置文件\\clear_Key.ini', 'w') as clear_keyconfigfile:
            clear_keyconfig.write(clear_keyconfigfile)
if __name__ == '__main__':
    ui = UI()