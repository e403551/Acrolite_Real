# Name: Acrobat-Light [Mk III]
# Author(s): Nic Benedetto
# Date: June 29 [Mk I], June 30 [Mk II], July 9 [Mk III -- GUI]

from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QLabel,QLineEdit,QPushButton,QApplication,QMainWindow
from ally_go_at_GUI import *
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap, QFont
import sys
import os

# establish global variables
gloss_matches = []
follow_ups = []
index_iterable = iter(range(0,10000000000000000))
index_iterable_babysitter = []
to_customer_dict = {}
out_of_terms = []

# establish foundational pane
app = QApplication(sys.argv)
app_icon = QIcon("topleft_logo.png")
app.setWindowIcon(app_icon)

win = QMainWindow()
win.setGeometry(600,300,600,600)
win.setFixedSize(600,600)
win.setWindowTitle("Acrolite")
win.setAttribute(Qt.WA_StyledBackground,True)
win.setStyleSheet("QMainWindow { background-color: brown; border-style: solid; border-width: 2px; border-color: lightblue; padding: 16px; }")

# create file name prompt
filenameprompt = QLabel(win)
filenameprompt.setText("File Name")
filenameprompt.setGeometry(30,100,84,30)
filenameprompt.setStyleSheet('color: darkblue; background: lightgray;border-style: solid; border-width: 1px; border-color: black; }"')

# Create file name textbox
file_name_textbox = QLineEdit(win)
file_name_textbox.move(120, 100)
file_name_textbox.resize(350, 30)
file_name_textbox.setStyleSheet('background: lightblue;border-style: solid; border-width: 1px; border-radius: 5px; border-color: black; }"')
# create path name example
format = QLabel(win)
format.setText("format: Document.docx")
format.setGeometry(120, 65, 170, 50)
format.setStyleSheet('color: darkblue;')

# append acrolite logo
image = QLabel(win)
image.setGeometry(100, -12, 300, 100)
pixmap = QPixmap("logo2.png")
image.setPixmap(pixmap)
image.show()

# create title holding cell
title = QLabel(win)
title.setText("")
title.setGeometry(205, 12, 190, 50)
title.setStyleSheet('color: black; background: darkgray;border-style: solid; border-width: 1px; border-radius: 5px; border-color: black; }"')
title.lower()

# create path name prompt
filepathprompt = QLabel(win)
filepathprompt.setText("Path Name")
filepathprompt.setGeometry(30, 150, 84, 30)
filepathprompt.setStyleSheet('color: darkblue; background: lightgray;border-style: solid; border-width: 1px; border-color: black; }"')

# create top holding cell
top = QLabel(win)
top.setText("")
top.setGeometry(15, 70, 570, 175)
top.setStyleSheet('color: black; background: darkgray;border-style: solid; border-width: 1px; border-radius: 5px; border-color: black; }"')
top.lower()

# Create path textbox
file_path_textbox = QLineEdit(win)
file_path_textbox.move(120, 150)
file_path_textbox.resize(450, 30)
file_path_textbox.setStyleSheet(
    'background: lightblue;border-style: solid; border-width: 1px; border-radius: 5px; border-color: black; }"')

# create performance display
performance_display = QPushButton(win)
performance_display.setText("")
myFont=QFont()
myFont.setBold(True)
performance_display.setFont(myFont)
performance_display.setGeometry(50, 200, 500, 30)
performance_display.setStyleSheet("QPushButton { color: yellow;background-color: darkgray; border-style: solid }"
                               "QPushButton:pressed { background-color: darkgray }" )

# define button for collection
collect_Button = QPushButton('Find and Scan', win)
collect_Button.setGeometry(240, 200,115,30)
collect_Button.setStyleSheet("QPushButton { background-color: lightgray; border-style: outset; border-width: 2px; border-radius: 5px; border-color: lightgray; padding: 4px; }"
                               "QPushButton:pressed { background-color: darkblue }" )

# define button for 'blacklist'
blacklist_Button = QPushButton('', win)
blacklist_Button.setGeometry(175, 400,115,30)
blacklist_Button.setStyleSheet("QPushButton { color: black; background-color: pink; border-style: outset; border-width: 2px; border-radius: 15px; border-color: lightgray; padding: 4px; }"
                               "QPushButton:pressed { background-color: darkblue }" )
blacklist_Button.setEnabled(False)


# define button for 'idk'
unknown_Button = QPushButton('', win)
unknown_Button.setGeometry(300, 400,115,30)
unknown_Button.setStyleSheet("QPushButton { color: black; background-color: blue; border-style: outset; border-width: 2px; border-radius: 15px; border-color: lightgray; padding: 4px; }"
                               "QPushButton:pressed { background-color: darkblue }" )
unknown_Button.setEnabled(False)


# define button for 'correct pair'
correct_Button = QPushButton('', win)
correct_Button.setGeometry(243, 355,115,30)
correct_Button.setStyleSheet("QPushButton { background-color: lightgreen; border-style: outset; border-width: 2px; border-radius: 15px; border-color: lightgray; padding: 4px; }"
                               "QPushButton:pressed { background-color: darkblue }" )
correct_Button.setEnabled(False)


# define button for 'ok'
ok_Button = QPushButton('Define', win)
ok_Button.setGeometry(177, 490,60,30)
ok_Button.setStyleSheet("QPushButton { background-color: lightgray; border-style: outset; border-width: 2px; border-radius: 5px; border-color: lightgray; padding: 4px; }"
                               "QPushButton:pressed { background-color: darkblue }" )
ok_Button.setEnabled(False)

# create bottom holding cell
bottom = QLabel(win)
bottom.setText("")
bottom.setGeometry(15, 255, 570, 330)
bottom.setStyleSheet('background: darkgray;border-style: solid; border-width: 1px; border-radius: 5px; border-color: black; }"')
bottom.lower()

# define display window for list elements (ie the pair itself)
display_window = QPushButton(win)
display_window.setText("")
myFont=QFont()
myFont.setBold(True)
display_window.setFont(myFont)
display_window.setGeometry(20, 303,560,38)
display_window.setStyleSheet("QPushButton { color: black; background-color: lightgray; border-style: outset; border-width: 2px; border-radius: 5px; border-color: black; padding: 4px; }"
                               "QPushButton:pressed { background-color: lightgray }" )

# define display title for list elements (ie stating a pair was found)
display_title = QLabel(win)
display_title.setText("")
myFont=QFont()
myFont.setBold(True)
display_title.setFont(myFont)
display_title.setGeometry(250, 260,95,30)
display_title.setStyleSheet('color: darkblue; background: lightgray;border-style: solid; border-width: 1px; border-color: black; }"')

# define correction title
correction_title = QLabel(win)
correction_title.setText("")
correction_title.setGeometry(177, 450,240,30)
correction_title.setStyleSheet('color: darkblue; background: lightgray;border-style: solid; border-width: 1px; border-color: black; }"')

# define 'no suffix'/'no entry' warning messages
warning = QLabel(win)
warning.setText("")
warning.setGeometry(120, 125,500,30)
warning.setStyleSheet('color: red;')

# define 'no definition' warning message
nodef = QLabel(win)
nodef.setText("")
nodef.setGeometry(205, 512,450,30)
nodef.setStyleSheet('color: red;')

# define button for 'Customer List'
customer_Button = QPushButton('Print Customer List', win)
customer_Button.setGeometry(220, 535,150,30)
customer_Button.setStyleSheet("QPushButton { background-color: lightgray; border-style: outset; border-width: 2px; border-radius: 5px; border-color: lightgray; padding: 4px; }"
                               "QPushButton:pressed { background-color: darkblue }" )
customer_Button.setEnabled(False)


def collectButton(): # determine file path, display options, export initialized list of gloss_matches and follow_ups
    file_name = file_name_textbox.text()
    file_path = file_path_textbox.text()
    nodef.setText("")

    try:
        if "." in file_name and ".docx" not in file_path:
            # disable button
            collect_Button.setText("Scanning...")
            collect_Button.setEnabled(False)
            file_path = file_path_textbox.text()
            doc2bscanned = r"" + str(file_path) + "\\" + str(file_name)

            print(doc2bscanned)

            # scan the doc for words
            scan_products = collect_doc_words(doc2bscanned)
            doc_words = scan_products[0]
            performance = scan_products[1]

            # doc found now
            collect_Button.hide()
            performance_display.setText(str(performance))

            warning.setText("")
            file_name_textbox.setEnabled(False)
            file_path_textbox.setEnabled(False)

            # enable lower buttons
            unknown_Button.setEnabled(True)
            correct_Button.setEnabled(True)
            blacklist_Button.setEnabled(True)
            ok_Button.setEnabled(True)
            customer_Button.setEnabled(True)

            # pick out acronyms in word doc and possible definitions
            definitions = detect_pairs_above(doc_words)

            # intake definitions and determine glossary matches and blacklist deletions
            data_processing_products = data_processing(definitions)
            g_matches = data_processing_products[0]
            f_ups = data_processing_products[1]
            correct_Button.setText("Correct")
            blacklist_Button.setText("Blacklist")
            unknown_Button.setText("Unknown")


            gloss_matches.append(g_matches)
            follow_ups.append(f_ups)

            # present the next list element
            next_element = follow_ups[0][next(index_iterable)]
            if next_element[1] == "":
                display_title.setText("Found Acronym")
                display = str(next_element[0])
                correct_Button.hide()
                display_title.setGeometry(225, 260, 142, 30)
                correction_title.setText("Please Supply Definition Here:")
                correction_title.setGeometry(180, 450, 225, 30)

            else:
                display_title.setText("Found Pair")
                display = "" + str(next_element[0]) + " : " + str(next_element[1]) + ""
                correct_Button.show()
                display_title.setGeometry(250, 260, 100, 30)
                correction_title.setText("Wrong Definition? Correct Here:")
                correction_title.setGeometry(177, 450, 240, 30)


            print("next element is " + str(next_element))
            print(display)
            display_window.setText(display)
            index_iterable_babysitter.append(next_element)
        elif ".docx" in file_path:
            warning.setText("Do not include the document itself in 'Path Name'")
        elif file_name == "":
            warning.setText("No 'File Name' input detected.")
        elif file_path == "":
            warning.setText("No 'Path Name' input detected.")
        else:
            warning.setText("No suffix detected in 'File Name'")

    except FileNotFoundError:
        collect_Button.setText("Find and Scan")
        collect_Button.setEnabled(True)
        collect_Button.show()
        warning.setText("Requested document not found")


# connect collect button to real button
collect_Button.clicked.connect(collectButton)

def correctButton():
    try:
        nodef.setText("")

        # found glossary dictionary from what's already on csv
        gloss_values = read_from_csv(dictionary={}, csv_doc="lite_gloss.csv")

        # append correct pair into to_customer_dictand gloss values
        to_customer_dict[index_iterable_babysitter[len(index_iterable_babysitter)-1][0]] = index_iterable_babysitter[len(index_iterable_babysitter)-1][1]
        gloss_values[index_iterable_babysitter[len(index_iterable_babysitter)-1][0]] = index_iterable_babysitter[len(index_iterable_babysitter)-1][1]

        # make gloss matches a dictionary
        gloss_matches_dict = list_to_dict(gloss_matches[0])

        # combine gloss matches and to_customer
        to_customer = {**to_customer_dict, **gloss_matches_dict}

        # put to_customer in alphabetical order
        to_customer = sort_database_alphabetically(to_customer)

        # put glossary dict in alphabetical order here
        gloss_values = sort_database_alphabetically(gloss_values)

        # write to csv glossary
        write_to_csv(dictionary=gloss_values, csv_doc="lite_gloss.csv")

        # write to customer's list
        glossary_to_excel(to_customer)

        # present the next list element
        next_element = follow_ups[0][next(index_iterable)]
        if next_element[1] == "":
            display_title.setText("Found Acronym:")
            display = str(next_element[0])
            correct_Button.hide()
            display_title.setGeometry(235, 260, 135, 30)
            correction_title.setText("Please Supply Definition Here:")
            correction_title.setGeometry(180, 450, 225, 30)

        else:
            display_title.setText("Found Pair:")
            display = "" + str(next_element[0]) + " : " + str(next_element[1]) + ""
            correct_Button.show()
            display_title.setGeometry(250, 260, 100, 30)
            correction_title.setText("Wrong Definition? Correct Here:")
            correction_title.setGeometry(177, 450, 240, 30)

        display_window.setText(display)
        index_iterable_babysitter.append(next_element)

    except IndexError:
        unknown_Button.setEnabled(False)
        correct_Button.setEnabled(False)
        blacklist_Button.setEnabled(False)
        ok_Button.setEnabled(False)
        blacklist_Button.setText("")
        correct_Button.setText("")
        unknown_Button.setText("")
        display_window.setText("")
        display_title.setText("")
        correction_title.setText("")
        customer_Button.setStyleSheet(
            "QPushButton { color: white;background-color: green; border-style: outset; border-width: 2px; border-radius: 5px; border-color: green; padding: 4px; }"
            "QPushButton:pressed { background-color: darkblue }")





# connect correct button to real button
correct_Button.clicked.connect(correctButton)

def blackButton():
    try:
        nodef.setText("")

        # initialize blacklist for appending
        blacklist = csv_to_blacklist()

        # append blacklisted result to blacklist
        blacklist.append(index_iterable_babysitter[len(index_iterable_babysitter)-1][0])

        # send blacklist back
        blacklist_to_csv(blacklist)

        # present the next list element
        next_element = follow_ups[0][next(index_iterable)]
        if next_element[1] == "":
            display_title.setText("Found Acronym:")
            display = str(next_element[0])
            correct_Button.hide()
            display_title.setGeometry(235, 260, 135, 30)
            correction_title.setText("Please Supply Definition Here:")
            correction_title.setGeometry(180, 450, 225, 30)

        else:
            display_title.setText("Found Pair:")
            display = "" + str(next_element[0]) + " : " + str(next_element[1]) + ""
            correct_Button.show()
            display_title.setGeometry(250, 260, 100, 30)
            correction_title.setText("Wrong Definition? Correct Here:")
            correction_title.setGeometry(177, 450, 240, 30)

        print("next element is " + str(next_element))
        print(display)
        display_window.setText(display)
        index_iterable_babysitter.append(next_element)
    except IndexError:
        unknown_Button.setEnabled(False)
        correct_Button.setEnabled(False)
        blacklist_Button.setEnabled(False)
        ok_Button.setEnabled(False)
        blacklist_Button.setText("")
        correct_Button.setText("")
        unknown_Button.setText("")
        display_window.setText("")
        display_title.setText("")
        correction_title.setText("")
        customer_Button.setStyleSheet(
            "QPushButton { color: white; background-color: green; border-style: outset; border-width: 2px; border-radius: 5px; border-color: green; padding: 4px; }"
            "QPushButton:pressed { background-color: darkblue }")

# connect black button to real button
blacklist_Button.clicked.connect(blackButton)

def unknownButton():
    try:
        # present the next list element
        nodef.setText("")
        next_element = follow_ups[0][next(index_iterable)]
        if next_element[1] == "":
            display_title.setText("Found Acronym:")
            display = str(next_element[0])
            correct_Button.hide()
            display_title.setGeometry(235, 260, 135, 30)
            correction_title.setText("Please Supply Definition Here:")
            correction_title.setGeometry(180, 450, 225, 30)

        else:
            display_title.setText("Found Pair:")
            display = "" + str(next_element[0]) + " : " + str(next_element[1]) + ""
            correct_Button.show()
            display_title.setGeometry(250, 260, 100, 30)
            correction_title.setText("Wrong Definition? Correct Here:")
            correction_title.setGeometry(177, 450, 240, 30)

        print("next element is " + str(next_element))
        print(display)
        display_window.setText(display)
        index_iterable_babysitter.append(next_element)
    except IndexError:
        unknown_Button.setEnabled(False)
        correct_Button.setEnabled(False)
        blacklist_Button.setEnabled(False)
        ok_Button.setEnabled(False)
        blacklist_Button.setText("")
        correct_Button.setText("")
        unknown_Button.setText("")
        display_window.setText("")
        display_title.setText("")
        correction_title.setText("")
        customer_Button.setStyleSheet(
            "QPushButton { color: white; background-color: green; border-style: outset; border-width: 2px; border-radius: 5px; border-color: green; padding: 4px; }"
            "QPushButton:pressed { background-color: darkblue }")
# connect unknown button to real button
unknown_Button.clicked.connect(unknownButton)

# Create correction textbox
correction_textbox = QLineEdit(win)
correction_textbox.move(250, 490)
correction_textbox.resize(166, 30)
correction_textbox.setStyleSheet('background: lightblue;border-style: solid; border-width: 1px; border-radius: 5px; border-color: black; }"')



def okButton():
    correct_definition = correction_textbox.text()

    try:

        if correct_definition == "":
            nodef.setText("No user definition detected.")
        else:
            nodef.setText("")

            # found glossary dictionary from what's already on csv
            gloss_values = read_from_csv(dictionary={},csv_doc="lite_gloss.csv")

            # append correction and acronym to glossary, gloss dict here
            correct_definition = correction_textbox.text()
            acronym = index_iterable_babysitter[len(index_iterable_babysitter) - 1][0]
            gloss_values[acronym] = correct_definition

            # append new pair into to_customer_dict
            to_customer_dict[acronym] = correct_definition

            # make gloss matches a dictionary
            gloss_matches_dict = list_to_dict(gloss_matches[0])

            # combine gloss matches and to_customer
            to_customer = {**to_customer_dict, **gloss_matches_dict}

            # put to_customer in alphabetical order
            to_customer = sort_database_alphabetically(to_customer)

            # put glossary dict in alphabetical order here
            gloss_values = sort_database_alphabetically(gloss_values)

            # write to csv glossary
            write_to_csv(dictionary=gloss_values, csv_doc = "lite_gloss.csv")

            # write to customer's list
            glossary_to_excel(to_customer)

            # present the next list element
            next_element = follow_ups[0][next(index_iterable)]
            if next_element[1] == "":
                display_title.setText("Found Acronym:")
                display = str(next_element[0])
                correct_Button.hide()
                display_title.setGeometry(235, 260, 135, 30)
                correction_title.setText("Please Supply Definition Here:")
                correction_title.setGeometry(180, 450, 225, 30)

            else:
                display_title.setText("Found Pair:")
                display = "" + str(next_element[0]) + " : " + str(next_element[1]) + ""
                correct_Button.show()
                display_title.setGeometry(250, 260, 100, 30)
                correction_title.setText("Wrong Definition? Correct Here:")
                correction_title.setGeometry(177, 450, 240, 30)

            print("next element is " + str(next_element))
            print(display)
            display_window.setText(display)
            index_iterable_babysitter.append(next_element)
            correction_textbox.clear()

    except IndexError:
        unknown_Button.setEnabled(False)
        correct_Button.setEnabled(False)
        blacklist_Button.setEnabled(False)
        ok_Button.setEnabled(False)
        customer_Button.setEnabled(False)
        blacklist_Button.setText("")
        correct_Button.setText("")
        unknown_Button.setText("")
        display_window.setText("")
        display_title.setText("")
        correction_title.setText("")
        customer_Button.setStyleSheet(
            "QPushButton { color: white; background-color: green; border-style: outset; border-width: 2px; border-radius: 5px; border-color: green; padding: 4px; }"
            "QPushButton:pressed { background-color: darkblue }")

def Nic_Benedetto():
    sys.exit(app.exec_())

def helpMessage():
    message = "--File Name--" \
              "\nTo find the selected file, simply copy it's title (including whatever suffix like .docx/.pdf/etc) into the 'File Name' blank. " \
              "\n\n--Path Name--\nTo find the selected file path, right click on the selected file. Select the 'properties' option. Go to the 'general' tab in the interface, and find the 'Location' information. Copy this path, making sure that the desired document is not in it, and paste it into the 'Path Name' blank." \
              "\n\nNOTE: If Acrolite cannot find the file, try going to the doc's properties (as instructed above) and 'Select All', copy, and paste into the respective blanks for both the name and the path. You may simply be leaving off portions of their names/path if they're particularly long."
    QMessageBox.about(win,"Help", str(message))

def aboutMessage():
    message = "'Acrolite' is an intern-designed and intern-built tool for scrubbing any word document and producing a customer-ready acronym list for those used in the document." \
              "\n\nThe process is largely automated by the creation of a local glossary with each use, scanning mechanisms of applicable documents that predict the definition," \
              " and can be user-pruned to block false acronyms through use of a local blacklist. \n\nAcrolite was designed to ensure at least one engineer has seen and approved the acronym" \
              " and definition before adding to the customer list.\n\nAlso: Acrolite can scan directly from Gunbarrel files. Just make sure to get the name and path right -- follow the 'Help' guidelines if needed."
    QMessageBox.about(win, "About", str(message))

def openCustomerList():
    customer_Button.setText("Opening List...")
    file = "Customer_List.xlsx"
    os.startfile(file)


# wire customer list retrieval button
customer_Button.clicked.connect(openCustomerList)

# define button for about
about_Button = QPushButton('About', win)
about_Button.setGeometry(493, 23,75,30)
about_Button.setStyleSheet("QPushButton { background-color: lightgray; border-style: outset; border-width: 2px; border-radius: 5px; border-color: lightgray; padding: 4px; }"
                               "QPushButton:pressed { background-color: darkblue }" )
about_Button.clicked.connect(aboutMessage)

# create about holding cell
about_holder = QLabel(win)
about_holder.setText("")
about_holder.setGeometry(475, 12,110,50)
about_holder.setStyleSheet('color: black; background: darkgray;border-style: solid; border-width: 1px; border-radius: 5px; border-color: black; }"')
about_holder.lower()

# define button for quit
quit_button = QPushButton('Quit', win)
quit_button.setGeometry(32, 23,80,30)
quit_button.setStyleSheet("QPushButton { background-color: lightgray; border-style: outset; border-width: 2px; border-radius: 5px; border-color: lightgray; padding: 4px; }"
                               "QPushButton:pressed { background-color: red }" )
quit_button.clicked.connect(Nic_Benedetto)

# create about holding cell
quit_holder = QLabel(win)
quit_holder.setText("")
quit_holder.setGeometry(15, 12,115,50)
quit_holder.setStyleSheet('color: black; background: darkgray;border-style: solid; border-width: 1px; border-radius: 5px; border-color: black; }"')
quit_holder.lower()

# define button for help
help_Button = QPushButton('Help', win)
help_Button.setGeometry(480, 100,90,30)
help_Button.setStyleSheet("QPushButton { background-color: lightgray; border-style: outset; border-width: 2px; border-radius: 5px; border-color: lightgray; padding: 4px; }"
                               "QPushButton:pressed { background-color: darkblue }" )
help_Button.clicked.connect(helpMessage)

# create SpaceCows plug
spacecows = QLabel(win)
spacecows.setText("Product of SpaceCows")
spacecows.setGeometry(20, 545, 170, 50)
spacecows.setStyleSheet('color: darkblue;')

# create version mark
version = QLabel(win)
version.setText("Version 1.3")
version.setGeometry(500, 545, 170, 50)
version.setStyleSheet('color: darkblue;')


# connect ok button to real button
ok_Button.clicked.connect(okButton)

# show window, exit
win.show()
sys.exit(app.exec_()) # standard, provides a clean exit
