# trying to establish logic understanding of PyQt
from PyQt5 import QtWidgets
from ally_go_at_GUI import *
import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QAction, QLineEdit, QMessageBox, QLabel
from gui_ally_mk2_acrolite import *
import time

def establish_static_components(win):

    plug = QLabel(win)
    plug.setText("Product of SpaceCows™ est. 2020")
    plug.setGeometry(160, 10, 300, 30)
    version = QLabel(win)
    version.setText("Acrolite Version 2.1")
    version.setGeometry(230, 50, 300, 30)

    # create file name prompt
    filenameprompt = QLabel(win)
    filenameprompt.setText("File Name:")
    filenameprompt.setGeometry(15, 90, 100, 50)
    # Create file name textbox
    file_name_textbox = QLineEdit(win)
    file_name_textbox.move(100, 100)
    file_name_textbox.resize(200, 30)
    # create path name example
    sfilenameprompt = QLabel(win)
    filenameprompt.setText("format: Word.docx")
    filenameprompt.setGeometry(250, 85, 150, 50)

    # create path name prompt
    filenameprompt = QLabel(win)
    filenameprompt.setText("Path Name:")
    filenameprompt.setGeometry(15, 140, 100, 50)
    # Create path textbox
    file_path_textbox = QLineEdit(win)
    file_path_textbox.move(100, 150)
    file_path_textbox.resize(400, 30)

    find_doc_button = QPushButton('Find and Scan', win)
    find_doc_button.setGeometry(250, 200, 115, 30)

    # connect button to function get_file, change file_path global variable to return file name
    find_doc_button.clicked.connect(button1_clicked)

def button1_clicked():
    print("Button1 was clicked")
