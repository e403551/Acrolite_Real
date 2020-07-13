import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QAction, QLineEdit, QMessageBox, QLabel
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot
from gui_ally_mk2_acrolite import *
import time


class App(QMainWindow): # self is an instance, here the instance of a window
# instead of producing data from a button, you want to *change* existing *global* data with a button
# TODO: I'm just daisy chaining these bad boys...instead of calling sequentially, you sequence them by nesting within each other, instead of producing variables, you change global variables--init only goes thru once but the buttons are always hot, looping
    def __init__(self):
        super().__init__()
        self.title = 'Acrolite'
        self.left = 600
        self.top = 300
        self.width = 600
        self.height = 600
        self.initUI()

        # define glossary and follow up matches
        self.gloss_matches = None
        self.follow_ups = None
        self.count = 0

        # define file_path variable
        self.file_path = None
        self.file_retrieval()
        self.show()




    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        # self.setStyleSheet("background-color: gray;")

        # create the plugs
        # TODO: SEND THESE TO THE BOTTOM LEFT AND RIGHT CORNER OF THE WINDOW
        self.plug = QLabel(self)
        self.plug.setText("Product of SpaceCows™ est. 2020")
        self.plug.setGeometry(160, 10, 300, 30)
        self.version = QLabel(self)
        self.version.setText("Acrolite Version 2.1")
        self.version.setGeometry(230, 50, 300, 30)
        self.version.setAlignment

    def file_retrieval(self):

        # create file name prompt
        self.filenameprompt = QLabel(self)
        self.filenameprompt.setText("File Name:")
        self.filenameprompt.setGeometry(15,90,100,50)
        # Create file name textbox
        self.file_name_textbox = QLineEdit(self)
        self.file_name_textbox.move(100, 100)
        self.file_name_textbox.resize(200, 30)
        # create path name example
        self.filenameprompt = QLabel(self)
        self.filenameprompt.setText("format: Word.docx")
        self.filenameprompt.setGeometry(250, 85, 150, 50)

        # create path name prompt
        self.filenameprompt = QLabel(self)
        self.filenameprompt.setText("Path Name:")
        self.filenameprompt.setGeometry(15, 140, 100, 50)
        # Create path textbox
        self.file_path_textbox = QLineEdit(self)
        self.file_path_textbox.move(100, 150)
        self.file_path_textbox.resize(400, 30)

        # Create a button to find doc
        self.find_doc_button = QPushButton('Find and Scan', self)
        self.find_doc_button.setGeometry(250, 200,115,30)

        # connect button to function get_file, change file_path global variable to return file name
        self.find_doc_button.clicked.connect(self.get_file_and_read_back)

    def processing(self):

        print("at processing:")
        doc2bscanned = self.file_path

        # pull every word from document
        cdw_products = collect_doc_words(doc2bscanned)
        doc_words = cdw_products[0]
        scan_performance = cdw_products[1]

        print(scan_performance)
        # TODO: you must explicitly show these labels
        # append performance stats
        self.performance = QLabel(self)
        self.performance.setText(str(scan_performance))
        self.performance.setGeometry(300, 250,400,30)
        self.performance.show()

        print("doc words are " + str(doc_words))

        # pick out acronyms in word doc and possible definitions
        definitions = detect_pairs_above(doc_words)

        print("definitions are " + str(definitions))

        # intake definitions and determine glossary matches and blacklist deletions
        data_processing_products = data_processing(definitions)
        gloss_matches = data_processing_products[0]
        follow_ups = data_processing_products[1]

        self.gloss_matches = gloss_matches
        self.follow_ups = follow_ups
        self.user_affirmation()

    def user_affirmation(self):


        # initialize customer and glossary dictionary
        to_customer = list_to_dict(self.gloss_matches)
        to_glossary = list_to_dict(self.gloss_matches)

        # initialize blacklist for appending
        blacklist = csv_to_blacklist()

        print("seeme")



        # initiate user conversation
        count = self.count
        print(count)
        for follow_up in self.follow_ups:
            breakout = False
            if follow_up[1] == "" and count == 0:
                print("--No definition for '" + str(follow_up[0]) + "' was found--")
                while not breakout:
                    self.follow = QLabel(self)
                    self.follow.setText(str("--No definition for '" + str(follow_up[0]) + "' was found--"))
                    self.follow.setGeometry(50, 300, 500, 50)
                    self.follow.show()
                    count += 1
                    self.next = QPushButton('Next', self)
                    self.next.setGeometry(300, 200, 115, 30)

                    self.next.show()

                    # connect button to function get_file, change file_path global variable to return file name
                    self.next.clicked.connect(self.next_term)
                    breakout = True






        ################# new up

        # send blacklist back
        blacklist_to_csv(blacklist)

        # print("TO GLOSS IS " + str(to_glossary))

        write_to_csv(to_glossary, "lite_gloss.csv")

        # print("gloss matches are " + str(self.gloss_matches))
        # print("follow ups are " + str(self.follow_ups))



    def next_term(self):
        print("clicked next")
        self.count = 0


    def get_file_and_read_back(self):
        file_name = self.file_name_textbox.text()



        print("file path is " + str(self.file_path))

        if "." in file_name:
            self.find_doc_button.setText("Scanning...")
            self.find_doc_button.setEnabled(False)
            file_path = self.file_path_textbox.text()
            doc2bscanned = r"" + str(file_path) + "\\" + str(file_name)
            print(str(doc2bscanned))
            self.file_path = doc2bscanned
            self.processing()

            # append glossary matches, ask user if follow-ups are right pairs, ask for definitions to missing acronyms
            # to_customer_dictionary = user_affirmation(gloss_matches, follow_ups)
        else:
            print("Include suffix")






if __name__ == '__main__':
    app = QApplication(sys.argv)
    sun = App()
    sys.exit(app.exec_())

