# -*- coding: utf-8 -*-

"""
/***************************************************************************
 Tree Volume Analysis System (T-VAS)
 A Tree Volume Analysis Algorithm
 This application (algorithm) is fully based on python and pandas modules.
 This application is based on 'Volume equations and biomass prediction of forest trees in Nepal (Sharma and Pukkala, 1990)
 and 'वन नियमावली, २०७९'. It can process and analyze of millions of data within few seconds efficiently.
                              -------------------
        begin                : 2022-11-01
        copyright            : (C) 2023 by Kapil Dev Adhikari
        email                : kapildevadk@gmail.com

***************************************************************************/
"""

# Import necessary modules
import pandas as pd
import numpy as np
import sys
import os
import xlsxwriter
from PyQt5 import QtCore
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

# Create a new class that inherits from QStyledItemDelegate for class General_Tab
class MyDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super(MyDelegate, self).__init__(parent)
        self.Times_font = QFont("Times New Roman", 11)
        self.kalimati_font = QFont("Kalimati", 11)

    def paint(self, painter, option, index):
        if index.column() == 12:
            option.font = self.Times_font
        else:
            option.font = self.kalimati_font
        super(MyDelegate, self).paint(painter, option, index)

class Nepal_TVAS(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(Nepal_TVAS, self).__init__(*args, **kwargs)
        self.setWindowTitle('Tree Volume Analysis System (T-VAS) Beta 1.0')
        icon = QIcon('icon.ico')
        self.setWindowIcon(QIcon('icon.ico'))
        self.resize(1000, 850)
        self._createToolBars()

        self.tab1 = Tab1() 
        self.tab2 = Tab2(self.tab1)
        self.tabs = QTabWidget()
        # self.tabs.addTab(self.tab1, "General Information")
        # self.tabs.addTab(self.tab2, "Main Data")
        # self.tabs.setFont(QFont("Times New Roman", 13))
        self.tabs.setContentsMargins(0, 50, 0, 0) # set top margin to 30
        self.tabs.addTab(self.tab1, "General Information")
        self.tabs.addTab(self.tab2, "Main Data")
        self.tabs.setFont(QFont("Helvetica", 13))
        icon = QIcon('pandas.png')
        self.tabs.setTabIcon(0, icon)
        self.tabs.setTabText(0, "General Information")
        icon = QIcon('data.png')
        self.tabs.setTabIcon(1, icon)
        self.tabs.setTabText(1, "Main Data")

        self.setCentralWidget(self.tabs)

        self.tab1.dataSubmitted.connect(self.update_df_gi)
    def update_df_gi(self):
        self.tab2.df_gi = self.tab1.main_df_gi.copy()

    def _createToolBars(self):

        # Create a QMenuBar object
        self.bar = QMenuBar(self)
        self.bar.setVisible(True)
        font = QFont("Times New Roman", 11)
        self.bar.setFont(font)

        # Add a menu to the menu bar
        self.file = self.bar.addMenu("File")
        self.file.setStyleSheet("QComboBox::item:hover { background-color: #0000FF; } QComboBox::item:selected { background-color: #079400; }") 
        
        # create a close action and add it to the File menu
        close_action = QAction('Close', self)
        close_action.setFont(QFont('Times New Roman', 11))
        close_action.setShortcut("Ctrl+Q")
        close_action.triggered.connect(self.closeEvent)
        self.file.addAction(close_action)

        self.help = self.bar.addMenu("Help")
        about_action = QAction('About', self)
        about_action.setFont(QFont('Times New Roman', 11))
        about_action.triggered.connect(self.about)
        self.help.addAction(about_action)

        about_action = QAction('Guide', self)
        about_action.setFont(QFont('Times New Roman', 11))
        about_action.triggered.connect(self.guide)
        self.help.addAction(about_action)

        about_action = QAction('Developer', self)
        about_action.setFont(QFont('Times New Roman', 11))
        about_action.triggered.connect(self.developer)
        self.help.addAction(about_action)

        # Set the menu bar on the main window
        self.setMenuBar(self.bar)

    # Define close function for main window
    # def close(self):
    #     # self.close()
    #     QApplication.quit()

    # Define function to add description to the "About" menu
    def about(self):
        # create a message box with the desired description
        about_msg = QMessageBox()
        about_msg.setFont(QFont('Times New Roman', 11))
        about_msg.setText("<p>Welcome to 'Tree Volume Analysis System (T-VIS) Beta 1.0' application! </p> \n <p>This application (algorithm) is fully based on python and pandas modules </p> \n <p>This application is based on: <a href='https://www.researchgate.net/publication/313927612_Volume_equations_and_biomass_prediction_of_forest_trees_in_Nepal'>Volume equations and biomass prediction of forest trees in Nepal (Sharma and Pukkala, 1990)</a> and 'वन नियमावली, २०७९'. </p> \n <p>It has a beta version of window interface with many widgets to add but user-friendly. </p> \n <p>This is easy to use and perform analysis of millions of data within few seconds efficiently. </p> \n <p>If you have any questions or suggestions, things to improvement, please don't hesitate to contact. </p>")
        about_msg.setWindowTitle("About")
        about_msg.exec_()

    def guide(self):
        guideline_msg = QMessageBox()
        guideline_msg.setFont(QFont('Times New Roman', 11))
        guideline_msg.setText("Coming soon...!")
        guideline_msg.setWindowTitle("Guide")
        guideline_msg.exec_()
    
    def developer(self):
        developer_msg = QMessageBox()
        developer_msg.setFont(QFont('Times New Roman', 11))
        developer_msg.setText("<p>Kapil Dev Adhikari </p> \n <p>Email: kapildevadk@gmail.com </p> \n <p>Contact No.: 9852083606 </p> \n <p> Github Profile: <a href='https://github.com/kapildevadk'>kapildevadk</a> </p>")
        developer_msg.setWindowTitle("Developer")
        developer_msg.exec_()
    
    # @pyqtSlot(QCloseEvent)
    def closeEvent(self, event):
        result = QMessageBox.question(
            self,
            "Confirm Exit...",
            "Are you sure you want to close ?",
            QMessageBox.Yes| QMessageBox.No
        )
        if result == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

class Tab1(QWidget):
    dataSubmitted = pyqtSignal()
    def __init__(self):
        super().__init__()
        """
        Create the General Information tab. Allows the user enter informations.
        """
        # self.main_df_gi = pd.DataFrame()
        self.main_df_gi = pd.DataFrame(columns=['प्रदेश', 'जिल्ला', 'कार्यालयको नाम', 'सव-डि.व.का./रे.पो.', 'वनको नाम', 'गा.पा./न.पा.',
                    'वडा नं.', 'खण्ड', 'उप-खण्ड', 'प्लट नं.','तथ्यांक संकलन मिति', 'तथ्यांक संकलकको नाम, पद', 'Coordinate System', 'आर्थिक वर्ष'])
        self.row_num = 0
        self.InitWindow()

    def InitWindow(self):
        # create table view
        self.table_view1 = QTableView(self)
        # ResizeToContents as per headers length
        self.table_view1.resizeColumnsToContents()

        # create lebels for input fields
        self.lbl_province = QLabel('प्रदेश:', self) # ComboBox
        self.lbl_province.setFont(QFont('kalimati', 11))
        self.lbl_district = QLabel('जिल्ला:', self) # ComboBox
        self.lbl_district.setFont(QFont('kalimati', 11))
        self.lbl_office_name = QLabel('कार्यालयको नाम:', self)
        self.lbl_office_name.setFont(QFont('kalimati', 11))
        self.lbl_sdfo_rp = QLabel('सव-डि.व.का./रे.पो.:', self)
        self.lbl_sdfo_rp.setFont(QFont('kalimati', 11))
        self.lbl_forest_name = QLabel('वनको नाम:', self)
        self.lbl_forest_name.setFont(QFont('kalimati', 11))
        self.lbl_gapa_napa = QLabel('गा.पा./न.पा.:', self)
        self.lbl_gapa_napa.setFont(QFont('kalimati', 11))
        self.lbl_ward = QLabel('वडा नं.:', self)
        self.lbl_ward.setFont(QFont('kalimati', 11))
        self.lbl_block = QLabel('खण्ड:', self)
        self.lbl_block.setFont(QFont('kalimati', 11))
        self.lbl_sub_block = QLabel('उप-खण्ड:', self)
        self.lbl_sub_block.setFont(QFont('kalimati', 11))
        self.lbl_plot = QLabel('प्लट नं.:', self)
        self.lbl_plot.setFont(QFont('kalimati', 11))
        self.lbl_date = QLabel('तथ्यांक संकलन मिति:', self) #self date 
        self.lbl_date.setFont(QFont("kalimati", 11))
        self.lbl_res_person = QLabel('तथ्यांक संकलकको नाम, पद:', self)
        self.lbl_res_person.setFont(QFont('kalimati', 11))
        self.lbl_crs = QLabel('Coordinate System:', self) #CRS
        self.lbl_crs.setFont(QFont("Times New Roman", 11))
        self.lbl_fy = QLabel('आर्थिक वर्ष:', self)
        self.lbl_fy.setFont(QFont('kalimati', 11))
        
        # create input fields, set the width and height of the input field
        self.province = QComboBox(self) # प्रदेश
        self.province.setStyleSheet("QComboBox { border: 1px solid black; }")
        self.province.setFixedSize(200, 30)
        self.province.setFont(QFont("kalimati", 11))
        self.province.setToolTip("Please select Province")
        self.province.setStyleSheet("QComboBox::item:hover { background-color: #0000FF; } QComboBox::item:selected { background-color: #079400; }") 
        # add options to combobox 
        # for Prvince
        self.province.addItems(['प्रदेश १', 'मधेश प्रदेश', 'बागमती प्रदेश', 'गण्डकी प्रदेश', 'लुम्बिनी प्रदेश', 'कर्णाली प्रदेश', 'सुदूरपश्चिम प्रदेश'])
        self.province.setCurrentIndex(0)  # Set default selected index to 0

        self.district = QComboBox(self) # जिल्ला
        self.district.setStyleSheet("QComboBox { border: 1px solid black; }")
        self.district.setFixedSize(200, 30)
        self.district.setFont(QFont("kalimati", 11))
        self.district.setToolTip("Please select District")
        self.district.setStyleSheet("QComboBox::item:hover { background-color: #0000FF; } QComboBox::item:selected { background-color: #079400; }") 
        self.province.currentIndexChanged.connect(self.category_changed)
        self.category_changed()

        self.office_name = QLineEdit(self) # कार्यालयको नाम
        self.office_name.setAlignment(QtCore.Qt.AlignLeft)
        self.office_name.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.office_name.setFixedSize(200, 30)
        self.office_name.setFont(QFont("kalimati", 11))
        self.office_name.setToolTip("Please enter your office name")

        self.sdfo_rp = QLineEdit(self) # सव-डि.व.का./रे.पो.
        self.sdfo_rp.setAlignment(QtCore.Qt.AlignLeft)
        self.sdfo_rp.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.sdfo_rp.setFixedSize(200, 30)
        self.sdfo_rp.setFont(QFont("kalimati", 11))
        self.sdfo_rp.setToolTip("Please enter your Sub-DFO/Range Post name")
        
        self.forest_name = QLineEdit(self) # वनको नाम 
        self.forest_name.setAlignment(QtCore.Qt.AlignLeft)
        self.forest_name.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.forest_name.setFixedSize(200, 30)
        self.forest_name.setFont(QFont("kalimati", 11))
        self.forest_name.setToolTip("Please enter the forest name") 

        self.gapa_napa = QLineEdit(self) # गा.पा./न.पा.
        self.gapa_napa.setAlignment(QtCore.Qt.AlignLeft)
        self.gapa_napa.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.gapa_napa.setFixedSize(200, 30)
        self.gapa_napa.setFont(QFont("kalimati", 11))

        self.ward = QLineEdit(self) # वडा नं. 
        self.ward.setAlignment(QtCore.Qt.AlignRight)
        self.ward.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.ward.setFixedSize(200, 30)
        self.ward.setFont(QFont("kalimati", 11))

        self.block = QLineEdit(self) 
        self.block.setAlignment(QtCore.Qt.AlignRight)
        self.block.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.block.setFixedSize(200, 30)
        self.block.setFont(QFont("kalimati", 11))
        self.block.setToolTip("Please enter the forest block name/no.") 

        self.sub_block = QLineEdit(self) 
        self.sub_block.setAlignment(QtCore.Qt.AlignRight)
        self.sub_block.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.sub_block.setFixedSize(200, 30)
        self.sub_block.setFont(QFont("kalimati", 11))
        self.sub_block.setToolTip("Please enter the sub block name/no.") 

        self.plot = QLineEdit(self) 
        self.plot.setAlignment(QtCore.Qt.AlignRight)
        self.plot.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.plot.setFixedSize(200, 30)
        self.plot.setFont(QFont("kalimati", 11))
        self.plot.setToolTip("Please enter the plot no.") 

        self.crs = QComboBox(self)
        self.crs.setStyleSheet("QComboBox { border: 1px solid black; }")
        self.crs.setFixedSize(200, 30)
        self.crs.setFont(QFont("Times New Roman", 11))
        self.crs.setToolTip("Please select the coordinate System (CRS/EPSG)")
        self.crs.setStyleSheet("QComboBox::item:hover { background-color: #0000FF; } QComboBox::item:selected { background-color: #079400; }") 
        self.crs.addItems(['WGS84 UTM 45N', 'WGS84 UTM 44N'])

        self.fy = QLineEdit(self)
        self.fy.setAlignment(QtCore.Qt.AlignLeft)
        self.fy.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.fy.setFixedSize(200, 30)
        self.fy.setFont(QFont("kalimati", 11))

        self.date = QDateEdit(self)
        self.date.setAlignment(QtCore.Qt.AlignCenter)
        current_date = QDate.currentDate()
        self.date.setDate(current_date)
        self.date.setCalendarPopup(True)
        self.date.setDisplayFormat("yyyy-MM-dd")
        self.date.setFixedSize(200, 30)
        self.date.setFont(QFont("kalimati", 11)) 

        self.res_person = QLineEdit(self) # लगत संकलकको नाम, पद
        self.res_person.setAlignment(QtCore.Qt.AlignLeft)
        self.res_person.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.res_person.setFixedSize(200, 30)
        self.res_person.setFont(QFont("kalimati", 11))
        self.res_person.setToolTip("Please enter the name of resource person") 

        # create submit button 
        self.bt_submit = QPushButton('Submit Data', self)
        self.bt_submit.setStyleSheet("background-color: #00A505; color: #fff;") 
        self.bt_submit.setFixedSize(120, 40)
        font = QFont()
        font.setBold(True)
        self.bt_submit.setFont(QFont("Times New Roman", 11))

        # create the delete button
        self.bt_delete = QPushButton("Delete Row", self)
        self.bt_delete.setStyleSheet("background-color: #d9534f; color: #fff;") 
        self.bt_delete.setFixedSize(120, 40)
        font = QFont()
        font.setBold(True)
        self.bt_delete.setFont(QFont("Times New Roman", 11))
        self.bt_delete.setEnabled(False)

        # create layout and add widgets
        # groupBox1 = QGroupBox("General Information")
        grid1 = QGridLayout(self)
        grid1.setSpacing(10)
        grid1.addWidget(self.lbl_province, 1, 0)
        grid1.setAlignment(self.lbl_province, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.province, 1, 1)
        grid1.setAlignment(self.province, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_district, 1, 2)
        grid1.setAlignment(self.lbl_district, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.district, 1, 3)
        grid1.setAlignment(self.district, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_office_name, 2, 0)
        grid1.setAlignment(self.lbl_office_name, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.office_name, 2, 1)
        grid1.setAlignment(self.office_name, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_sdfo_rp, 2, 2)
        grid1.setAlignment(self.lbl_sdfo_rp, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.sdfo_rp, 2, 3)
        grid1.setAlignment(self.sdfo_rp, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_forest_name, 3, 0)
        grid1.setAlignment(self.lbl_forest_name, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.forest_name, 3, 1)
        grid1.setAlignment(self.forest_name, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_gapa_napa, 3, 2)
        grid1.setAlignment(self.lbl_gapa_napa, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.gapa_napa, 3, 3)
        grid1.setAlignment(self.gapa_napa, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_ward, 4, 0)
        grid1.setAlignment(self.lbl_ward, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.ward, 4, 1)
        grid1.setAlignment(self.ward, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_block, 4, 2)
        grid1.setAlignment(self.lbl_block, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.block, 4, 3)
        grid1.setAlignment(self.block, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_sub_block, 5, 0)
        grid1.setAlignment(self.lbl_sub_block, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.sub_block, 5, 1)
        grid1.setAlignment(self.block, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_plot, 5, 2)
        grid1.setAlignment(self.lbl_plot, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.plot, 5, 3)
        grid1.setAlignment(self.plot, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_crs, 6, 0)
        grid1.setAlignment(self.lbl_crs, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.crs, 6, 1)
        grid1.setAlignment(self.crs, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_fy, 6, 2)
        grid1.setAlignment(self.lbl_fy, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.fy, 6, 3)
        grid1.setAlignment(self.fy, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_date, 7, 0)
        grid1.setAlignment(self.lbl_date, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.date, 7, 1)
        grid1.setAlignment(self.date, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.lbl_res_person, 7, 2)
        grid1.setAlignment(self.lbl_res_person, QtCore.Qt.AlignLeft)
        grid1.addWidget(self.res_person, 7, 3)
        grid1.setAlignment(self.res_person, QtCore.Qt.AlignLeft)

        grid1.addWidget(self.bt_submit, 8, 2)
        grid1.addWidget(self.bt_delete, 8, 3)
        # grid1.setSizeConstraint(QLayout.SetFixedSize)

        grid1.addWidget(self.table_view1, 10, 0, 14, 0)
        # Set the layout of the group box
        self.setLayout(grid1)
        
        # create the signal handler
        self.bt_submit.clicked.connect(self.submit_gi)
        self.bt_delete.clicked.connect(self.confirm_delete)     
        self.show()

    def submit_gi(self):
        # get values from input fields
        province = self.province.currentText()
        district = self.district.currentText()
        office_name = self.office_name.text()
        sdfo_rp = self.sdfo_rp.text()
        forest_name = self.forest_name.text()
        gapa_napa = self.gapa_napa.text()
        ward = self.ward.text()
        block = self.block.text()
        sub_block = self.sub_block.text()
        plot = self.plot.text()
        fy = self.fy.text()
        date = self.date.date().toPyDate()
        res_person = self.res_person.text()
        crs = self.crs.currentText()

        self.row_num += 1
        self.main_df_gi.loc[self.row_num, 'प्रदेश'] = province
        self.main_df_gi.loc[self.row_num, 'जिल्ला'] = district
        self.main_df_gi.loc[self.row_num, 'कार्यालयको नाम'] = office_name
        self.main_df_gi.loc[self.row_num, 'सव-डि.व.का./रे.पो.'] = sdfo_rp
        self.main_df_gi.loc[self.row_num, 'वनको नाम'] = forest_name
        self.main_df_gi.loc[self.row_num, 'गा.पा./न.पा.'] = gapa_napa
        self.main_df_gi.loc[self.row_num, 'वडा नं.'] = ward
        self.main_df_gi.loc[self.row_num, 'खण्ड'] = block
        self.main_df_gi.loc[self.row_num, 'उप-खण्ड'] = sub_block
        self.main_df_gi.loc[self.row_num, 'प्लट नं.'] = plot
        self.main_df_gi.loc[self.row_num, 'तथ्यांक संकलन मिति'] = date
        self.main_df_gi.loc[self.row_num, 'तथ्यांक संकलकको नाम, पद'] = res_person
        self.main_df_gi.loc[self.row_num, 'Coordinate System'] = crs
        self.main_df_gi.loc[self.row_num, 'आर्थिक वर्ष'] = fy

        self.dataSubmitted.emit()

        # clear forms input fields
        self.province.setCurrentText('')
        self.district.setCurrentText('')
        self.office_name.setText('')
        self.sdfo_rp.setText('')
        self.forest_name.setText('')
        self.gapa_napa.setText('')
        self.ward.setText('')
        self.block.setText('')
        self.sub_block.setText('')
        self.plot.setText('')
        self.fy.setText('')
        self.date.setDate(QDate.currentDate())
        self.res_person.setText('')
        self.crs.setCurrentText('')

        # update table view with dataframe
        self.model = PandasModel1(self.main_df_gi)
        self.table_view1.setModel(self.model)

        # Import QStyledItemDelegate class function here
        self.table_view1.setItemDelegate(MyDelegate(self))

        # ResizeToContents as per headers length
        self.table_view1.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        
        self.table_view1.horizontalHeader().setFont(QFont('Kalimati', 11))
        # activate dlete button here after submit
        self.bt_delete.setEnabled(True)

    def confirm_delete(self):
        selected_row = self.table_view1.currentIndex().row()
        if selected_row != -1:
            response = QMessageBox.question(self, 'Confirm Deletion', 'Are you sure you want to delete this row?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if response == QMessageBox.Yes:
                self.delete_data(selected_row)
            else:
                return
        else:
            QMessageBox.warning(self, 'No Row Selected', 'Please select a row to delete.')

    def delete_data(self, selected_row):
        self.main_df_gi = self.main_df_gi.drop(self.main_df_gi.index[selected_row])
        self.main_df_gi.reset_index(drop=True, inplace=True)
        self.row_num -= 1
        self.table_view1.setModel(PandasModel1(self.main_df_gi))
        self.table_view1.resizeColumnsToContents()
        if self.row_num == 0:
            self.bt_delete.setEnabled(False)

    def category_changed(self):
        category = self.province.currentText()
        if category == 'प्रदेश १':
            self.district.clear()
            self.district.addItems(["ताप्लेजुङ्ग","पाँचथर","इलाम","झापा","संखुवासभा","भोजपुर","तेह्रथुम","धनकुटा","मोरङ्ग","सुनसरी","सोलुखुम्बु","ओखलढुंगा","खोटाङ्ग","उदयपुर"])
        elif category == 'मधेश प्रदेश':
            self.district.clear()
            self.district.addItems(["पर्सा","बारा","रौतहट","सर्लाही","धनुषा","सिराहा","महोत्तरी","सप्तरी"])
        elif category == 'बागमती प्रदेश':
            self.district.clear()
            self.district.addItems(["सिन्धुली","रामेछाप","दोलखा","भक्तपुर","धादिङ","काठमाण्डौँ","काभ्रेपलान्चोक","ललितपुर","नुवाकोट","रसुवा","सिन्धुपाल्चोक","चितवन","मकवानपुर"])
        elif category == 'गण्डकी प्रदेश':
            self.district.clear()
            self.district.addItems(["बागलुङ्ग","गोरखा","कास्की","लमजुङ","मनाङ","मुस्ताङ","म्याग्दी","नवलपरासी (बर्दघाट सुस्ता पूर्व)","पर्वत","स्याङ्गजा","तनहुँ"])    
        elif category == 'लुम्बिनी प्रदेश':
            self.district.clear()
            self.district.addItems(["कपिलवस्तु","नवलपरासी (बर्दघाट सुस्ता पश्चिम)","रुपन्देही","अर्घाखाँची","गुल्मी","पाल्पा","दाङ","प्युठान","रोल्पा","पूर्वी रूकुम","बाँके","बर्दिया"])  
        elif category == 'कर्णाली प्रदेश':
            self.district.clear()
            self.district.addItems(["पश्चिमी रूकुम","सल्यान","डोल्पा","हुम्ला","जुम्ला","कालिकोट","मुगु","सुर्खेत","दैलेख","जाजरकोट"])
        elif category == 'सुदूरपश्चिम प्रदेश':
            self.district.clear()
            self.district.addItems(["कैलाली","अछाम","डोटी","बझाङ","बाजुरा","कञ्चनपुर","डडेलधुरा","बैतडी","दार्चुला"])

class PandasModel1(QAbstractTableModel):

    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parnet=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])

            column_count = self.columnCount()
            for column in range(0, column_count):
                if (index.column() == column and role == Qt.TextAlignmentRole):
                    value = self._data.iloc[index.row(), index.column()]
                    if isinstance(value, (int, float)):
                        return Qt.AlignRight | Qt.AlignVCenter
                    else:
                        return Qt.AlignHCenter | Qt.AlignVCenter

        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None

class Tab2(QWidget):
    def __init__(self, tab1_instance):
        super().__init__()
        
        """
        Create the main Data tab. Allows the user enter informations.
        """

        self.df_gi = pd.DataFrame()
        # For HTML table view create df5
        self.df5 = pd.DataFrame()
        # Create dataframe with parameters        
        self.df_par = pd.DataFrame({"प्रजाती (Species)": ['ठिङ्ग्रेसल्ला', 'खयर', 'हल्दु/कर्मा', 'शिरीष', 'उत्तिस', 'धौटी', 'सिमल', 'टुनी', 'सिसौँ', 'जामुन', 'भुडकुल', 'बोटधँगेरो', 'चाँप', 'खोटेसल्ला', 'गोब्रेसल्ला', 'खस्रु', 'चिलाउने', 'साल', 'असना', 'गुटेल', 'धूपी सल्ला', 'गम्हारी (अन्य तराई)', 'पाटेसल्ला (अन्य पहाड)', 'सतिसाल (अन्य तराई)', 'विजयसाल (अन्य तराई)', 'सागवन (अन्य तराई)', 'ओखर (अन्य पहाड)', 'दार (अन्य पहाड)', 'सन्दन/पाजन (अन्य तराई)', 'हर्रो (अन्य तराई)', 'बर्रो (अन्य तराई)', 'फल्दु (अन्य पहाड)', 'कटुस (अन्य पहाड)', 'चिलाउने (अन्य पहाड)', 'सौर (अन्य पहाड)', 'तालिसपत्र (अन्य पहाड)', 'देवदार (अन्य पहाड)', 'स्प्रुस (अन्य पहाड)', 'आँप (अन्य तराई)', 'मसला (अन्य तराई)', 'पपलर (अन्य तराई)', 'टिक (अन्य तराई)', 'बाँझी (अन्य तराई)', 'श्वेत चन्दन (अन्य पहाड)', 'तराईका अन्य प्रजाति', 'पहाडका अन्य प्रजाति'],
                                "a": [-2.4453, -2.3256, -2.5626, -2.4284, -2.7761, -2.272, -2.3865, -2.1832, -2.1959, -2.5693, -2.585, -2.3411, -2.0152, -2.977, -2.8195, -2.36, -2.7385, -2.4554, -2.4616, -2.4585, -2.5293, -2.3993, -2.3204, -2.3993, -2.3993, -2.3993, -2.3204, -2.3204, -2.3993, -2.3993, -2.3993, -2.3204, -2.3204, -2.3204, -2.3204, -2.3204, -2.3204, -2.3204, -2.3993, -2.3993, -2.3993, -2.3993, -2.3993, -2.3204, -2.3993, -2.3204],
                                "b": [1.722, 1.6476, 1.8598, 1.7609, 1.9006, 1.7499, 1.7414, 1.8679, 1.6567, 1.8816, 1.9437, 1.7246, 1.8555, 1.9235, 1.725, 1.968, 1.8155, 1.9026, 1.8497, 1.8043, 1.7815, 1.7836, 1.8507, 1.7836, 1.7836, 1.7836, 1.8507, 1.8507, 1.7836, 1.7836, 1.7836, 1.8507, 1.8507, 1.8507, 1.8507, 1.8507, 1.8507, 1.8507, 1.7836, 1.7836, 1.7836, 1.7836, 1.7836, 1.8507, 1.7836, 1.8507],
                                "c": [1.0757, 1.0552, 0.8783, 0.9662, 0.9428, 0.9174, 1.0063, 0.7569, 0.9899, 0.8498, 0.7902, 0.9702, 0.763, 1.0019, 1.1623, 0.7469, 1.0072, 0.8352, 0.88, 0.922, 1.0369, 0.9546, 0.8223, 0.9546, 0.9546, 0.9546, 0.8223, 0.8223, 0.9546, 0.9546, 0.9546, 0.8223, 0.8223, 0.8223, 0.8223, 0.8223, 0.8223, 0.8223, 0.9546, 0.9546, 0.9546, 0.9546, 0.9546, 0.8223, 0.9546, 0.8223],
                                "a1": [5.4443, 5.4401, 5.4681, 4.4031, 6.019, 4.9502, 4.5554, 4.9705, 4.358, 5.1749, 5.5572, 5.3349, 3.3499, 6.2696, 5.7216, 4.8511, 7.4617, 5.2026, 4.5968, 5.3475, 5.2774, 4.8991, 5.5323, 4.8991, 4.8991, 4.8991, 5.5323, 5.5323, 4.8991, 4.8991, 4.8991, 5.5323, 5.5323, 5.5323, 5.5323, 5.5323, 5.5323, 5.5323, 4.8991, 4.8991, 4.8991, 4.8991, 4.8991, 5.5323, 4.8991, 5.5323],
                                "b1": [-2.6902, -2.491, -2.491, -2.2094, -2.7271, -2.3353, -2.3009, -2.3436, -2.1559, -2.3636, -2.496, -2.4428, -2.0161, -2.8252, -2.6788, -2.4494, -3.0676, -2.4788, -2.2305, -2.4774, -2.6483, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406, -2.3406],
                                "small": [0.436, 0.443, 0.443, 0.443, 0.803, 0.443, 0.443, 0.443, 0.684, 0.443, 0.443, 0.443, 0.443, 0.189, 0.683, 0.747, 0.52, 0.055, 0.443, 0.443, 0.436, 0.443, 0.436, 0.443, 0.443, 0.443, 0.443, 0.443, 0.443, 0.443, 0.443, 0.443, 0.443, 0.443, 0.443, 0.436, 0.436, 0.436, 0.443, 0.443, 0.443, 0.443, 0.443, 0.443, 0.443, 0.443],
                                "medium": [0.372, 0.511, 0.511, 0.511, 1.226, 0.511, 0.511, 0.511, 0.684, 0.511, 0.511, 0.511, 0.511, 0.256, 0.488, 0.96, 0.186, 0.341, 0.511, 0.511, 0.372, 0.511, 0.372, 0.511, 0.511, 0.511, 0.511, 0.511, 0.511, 0.511, 0.511, 0.511, 0.511, 0.511, 0.511, 0.372, 0.372, 0.372, 0.511, 0.511, 0.511, 0.511, 0.511, 0.511, 0.511, 0.511],
                                "big": [0.355, 0.71, 0.71, 0.71, 1.51, 0.71, 0.71, 0.71, 0.684, 0.71, 0.71, 0.71, 0.71, 0.3, 0.41, 1.06, 0.168, 0.357, 0.71, 0.71, 0.355, 0.71, 0.355, 0.71, 0.71, 0.71, 0.71, 0.71, 0.71, 0.71, 0.71, 0.71, 0.71, 0.71, 0.71, 0.355, 0.355, 0.355, 0.71, 0.71, 0.71, 0.71, 0.71, 0.71, 0.71, 0.71]})

        self.df = pd.DataFrame()
        self.main_df = pd.DataFrame()
        self.df2 = pd.DataFrame()
        self.df3 = pd.DataFrame()
        self.df4 = pd.DataFrame()
        self.row_num = 0
        # self.InitWindow()

    # def InitWindow(self):
        # create table view
        self.table_view = QTableView(self)
        # self.table_view.setModel(self.model)
        # ResizeToContents as per headers length
        self.table_view.resizeColumnsToContents()

        # create lebels for input fields
        self.lbl_sn = QLabel('S.N.:', self)
        self.lbl_sn.setFont(QFont('Times New Roman', 11))
        self.lbl_sn.setAlignment(QtCore.Qt.AlignVCenter)
        self.lbl_xcor = QLabel('X-Coordinate:', self)
        self.lbl_xcor.setFont(QFont("Times New Roman", 11))
        self.lbl_ycor = QLabel('Y-Coordinate:', self)
        self.lbl_ycor.setFont(QFont("Times New Roman", 11))
        self.lbl_species = QLabel('प्रजाती (Species):', self)
        self.lbl_species.setFont(QFont("kalimati", 11))
        self.lbl_gd = QLabel('मापनः गोलाई/ब्यास:', self)
        self.lbl_gd.setFont(QFont("kalimati", 11))
        self.lbl_girthdiam = QLabel('गोलाई/ब्यास (cm):', self)
        self. lbl_girthdiam.setFont(QFont("kalimati", 11))
        self.lbl_t_height = QLabel('उचाई (Height) m.:', self)
        self.lbl_t_height.setFont(QFont("kalimati", 11))
        self.lbl_tree_class = QLabel('श्रेणी/दर्जा (Tree class):', self)
        self.lbl_tree_class.setFont(QFont("kalimati", 11))
        self.lbl_remark = QLabel('कैफियत:', self)
        self.lbl_remark.setFont(QFont("kalimati", 11))

        # create input fields
        self.sn = QLineEdit(self)
        self.sn.setAlignment(QtCore.Qt.AlignRight)
        self.sn.setStyleSheet("QLineEdit { border: 1px solid black; }")
        # Set the width and height of the input field
        self.sn.setFixedSize(100, 30)
        self.sn.setFont(QFont("kalimati", 11))
        self.sn.setToolTip("Please enter the serial number of tree")

        self.xcor = QLineEdit(self)
        self.xcor.setAlignment(QtCore.Qt.AlignRight)
        self.xcor.setStyleSheet("QLineEdit { border: 1px solid black; }")
        x = QDoubleValidator()
        self.xcor.setValidator(x)
        self.xcor.setValidator(QRegExpValidator(QRegExp("\d{6}")))
        self.xcor.setFixedSize(180, 30)
        self.xcor.setFont(QFont("kalimati", 11))
        self.xcor.setToolTip("Please enter the 'X-Coordinate' of tree")
        self.ycor = QLineEdit(self)
        self.ycor.setAlignment(QtCore.Qt.AlignRight)
        self.ycor.setStyleSheet("QLineEdit { border: 1px solid black; }")
        x = QDoubleValidator()
        self.ycor.setValidator(x)
        self.ycor.setValidator(QRegExpValidator(QRegExp("\d{7}")))
        self.ycor.setFixedSize(180, 30)
        self.ycor.setFont(QFont("kalimati", 11))
        self.ycor.setToolTip("Please enter the 'Y-Coordinate' of tree")
        self.species = QComboBox(self)
        self.species.setStyleSheet("QComboBox { border: 1.5px solid black; }")
        self.species.setFixedSize(250, 35)
        self.species.setFont(QFont("kalimati", 11))
        self.species.setToolTip("Please select the tree species")
        self.species.setStyleSheet("QComboBox::item:hover { background-color: #0000FF; } QComboBox::item:selected { background-color: #079400; }") 
        self.gd = QComboBox(self)
        self.gd.setStyleSheet("QComboBox { border: 1.5px solid black; }")
        self.gd.setFixedSize(180, 30)
        self.gd.setFont(QFont("kalimati", 11))
        self.gd.setToolTip("Please select measurement method: Girth or Diameter")
        self.gd.setStyleSheet("QComboBox::item:hover { background-color: #0000FF; } QComboBox::item:selected { background-color: #079400; }") 
        self.girthdiam = QLineEdit(self)
        self.girthdiam.setAlignment(QtCore.Qt.AlignRight)
        self.girthdiam.setStyleSheet("QLineEdit { border: 1px solid black; }")
        x = QDoubleValidator()
        self.girthdiam.setValidator(x)
        self.girthdiam.setFixedSize(180, 30)
        self.girthdiam.setFont(QFont("kalimati", 11))
        self.girthdiam.setToolTip("Please enter the DBH or Girth of tree in cm")
        # self.girthdiam.setValidator(QtGui.QDoubleValidator())
        self.t_height = QLineEdit(self)
        self.t_height.setAlignment(QtCore.Qt.AlignRight)
        self.t_height.setStyleSheet("QLineEdit { border: 1px solid black; }")
        x = QDoubleValidator()
        self.t_height.setValidator(x)
        self.t_height.setFixedSize(180, 30)
        self.t_height.setFont(QFont("kalimati", 11))
        self.t_height.setToolTip("Please enter height of tree in meter")
        self.tree_class = QComboBox(self)
        self.tree_class.setStyleSheet("QComboBox { border: 1.5px solid black; }")
        self.tree_class.setFixedSize(180, 30)
        self.tree_class.setFont(QFont("kalimati", 11))
        self.tree_class.setToolTip("Please select the tree class")
        self.tree_class.setStyleSheet("QComboBox::item:hover { background-color: #0000FF; } QComboBox::item:selected { background-color: #079400; }") 
        self.remark = QLineEdit(self)
        self.remark.setStyleSheet("QLineEdit { border: 1px solid black; }")
        self.remark.setFixedSize(180, 30)
        self.remark.setFont(QFont("kalimati", 12))
        self.remark.setToolTip('प्रजाती (Species) मा तराईका अन्य प्रजाति वा पहाडका अन्य प्रजाति भए स्थानिय नाम कैफियतमा लेख्‍न सक्नुहुन्छ ।')

        # add options to combobox for species
        self.species.addItems(['ठिङ्ग्रेसल्ला', 'खयर', 'हल्दु/कर्मा', 'शिरीष', 'उत्तिस', 'धौटी', 'सिमल', 'टुनी', 'सिसौँ', 
                                     'जामुन', 'भुडकुल', 'बोटधँगेरो', 'चाँप', 'खोटेसल्ला', 'गोब्रेसल्ला', 'खस्रु', 'चिलाउने', 'साल', 
                                     'असना', 'गुटेल', 'धूपी सल्ला', 'गम्हारी (अन्य तराई)', 'पाटेसल्ला (अन्य पहाड)', 
                                     'सतिसाल (अन्य तराई)', 'विजयसाल (अन्य तराई)', 'सागवन (अन्य तराई)', 'ओखर (अन्य पहाड)', 
                                     'दार (अन्य पहाड)', 'सन्दन/पाजन (अन्य तराई)', 'हर्रो (अन्य तराई)', 'बर्रो (अन्य तराई)', 
                                     'फल्दु (अन्य पहाड)', 'कटुस (अन्य पहाड)', 'चिलाउने (अन्य पहाड)', 'सौर (अन्य पहाड)', 
                                     'तालिसपत्र (अन्य पहाड)', 'देवदार (अन्य पहाड)', 'स्प्रुस (अन्य पहाड)', 'आँप (अन्य तराई)', 
                                     'मसला (अन्य तराई)', 'पपलर (अन्य तराई)', 'टिक (अन्य तराई)', 'बाँझी (अन्य तराई)', 
                                     'श्वेत चन्दन (अन्य पहाड)', 'तराईका अन्य प्रजाति', 'पहाडका अन्य प्रजाति'])
       
        # add options to combobox for selection of गोलाई/ब्यास
        self.gd.addItems(['ब्यास (Diameter)', 'गोलाई (Girth)'])
        # add options to combobox for selection of tree class
        self.tree_class.addItems(['१', '२', '३', '४'])

        # create submit button to do...
        self.btn_submit = QPushButton('Submit', self)
        self.btn_submit.setStyleSheet("background-color: #00A505; color: #fff;") 
        self.btn_submit.setFixedSize(80, 40)
        font = QFont()
        font.setBold(True)
        self.btn_submit.setFont(QFont("Times New Roman", 11))

        self.btn_show_table = QPushButton("Show Table", self)
        self.btn_show_table.setStyleSheet("background-color: #00BFFF;") 
        self.btn_show_table.setFixedSize(120, 40)
        font = QFont()
        font.setBold(True)
        self.btn_show_table.setFont(QFont("Times New Roman", 11))
        self.btn_show_table.setToolTip("Press 'Show Table' to view the analzed data")
        # self.btn_show_table.setEnabled(False)

        # create export button and connect to click event handler
        self.btn_export = QPushButton("Export to Excel", self)
        self.btn_export.setStyleSheet("background-color: #00BFFF;") 
        self.btn_export.setFixedSize(150, 40)
        font = QFont()
        font.setBold(True)
        self.btn_export.setFont(QFont("Times New Roman", 11))
        self.btn_export.setToolTip("Press 'Export to Excel' to export the analzed data as excel file")
        # self.btn_export.setEnabled(False)

        # create the delete button
        self.btn_delete = QPushButton("Delete Row", self)
        # Set the button's CSS properties
        self.btn_delete.setStyleSheet("background-color: #d9534f; color: #fff;") 
        self.btn_delete.setFixedSize(120, 40)
        font = QFont()
        font.setBold(True)
        self.btn_delete.setFont(QFont("Times New Roman", 11))
        self.btn_delete.setToolTip("Press 'Delete Row' to delete the selected row data")
        self.btn_delete.setEnabled(False)

        # create layout and add widgets
        # groupBox = QGroupBox("Main Data")
        grid = QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(self.lbl_sn, 1, 0)
        grid.setAlignment(self.lbl_sn, QtCore.Qt.AlignLeft)
        grid.addWidget(self.sn, 1, 1)
        grid.setAlignment(self.sn, QtCore.Qt.AlignLeft)
        grid.addWidget(self.lbl_xcor, 2, 0)
        grid.setAlignment(self.lbl_xcor, QtCore.Qt.AlignLeft)
        grid.addWidget(self.xcor, 2, 1)
        grid.setAlignment(self.xcor, QtCore.Qt.AlignLeft)
        grid.addWidget(self.lbl_ycor, 3, 0)
        grid.setAlignment(self.lbl_ycor, QtCore.Qt.AlignLeft)
        grid.addWidget(self.ycor, 3, 1)
        grid.setAlignment(self.ycor, QtCore.Qt.AlignLeft)
        grid.addWidget(self.lbl_species, 4, 0)
        grid.setAlignment(self.lbl_species, QtCore.Qt.AlignLeft)
        grid.addWidget(self.species, 4, 1)
        grid.setAlignment(self.species, QtCore.Qt.AlignLeft)
        grid.addWidget(self.lbl_gd, 5, 0)
        grid.setAlignment(self.lbl_gd, QtCore.Qt.AlignLeft)
        grid.addWidget(self.gd, 5, 1)
        grid.setAlignment(self.gd, QtCore.Qt.AlignLeft)
        grid.addWidget(self.lbl_girthdiam, 6, 0)
        grid.setAlignment(self.lbl_girthdiam, QtCore.Qt.AlignLeft)
        grid.addWidget(self.girthdiam, 6, 1)
        grid.setAlignment(self.girthdiam, QtCore.Qt.AlignLeft)
        grid.addWidget(self.lbl_t_height, 7, 0)
        grid.setAlignment(self.lbl_t_height, QtCore.Qt.AlignLeft)
        grid.addWidget(self.t_height, 7, 1)
        grid.setAlignment(self.t_height, QtCore.Qt.AlignLeft)
        grid.addWidget(self.lbl_tree_class, 8, 0)
        grid.setAlignment(self.lbl_tree_class, QtCore.Qt.AlignLeft)
        grid.addWidget(self.tree_class, 8, 1)
        grid.setAlignment(self.tree_class, QtCore.Qt.AlignLeft)
        grid.addWidget(self.lbl_remark, 9, 0)
        grid.setAlignment(self.lbl_remark, QtCore.Qt.AlignLeft)
        grid.addWidget(self.remark, 9, 1)
        grid.setAlignment(self.remark, QtCore.Qt.AlignLeft)

        grid.addWidget(self.btn_submit, 10, 1)
        grid.addWidget(self.btn_show_table, 10, 3)
        grid.addWidget(self.btn_export, 10, 5)
        grid.addWidget(self.btn_delete, 10, 7)
        grid.addWidget(self.table_view, 12, 0, 23, 0)
        # Set the layout of the group box
        self.setLayout(grid)

        # create the signal handler
        self.btn_submit.clicked.connect(self.submit)
        
        # Create an attribute for the instance of Tab1 in Tab2
        self.tab1 = tab1_instance
        # self.btn_show_table.clicked.connect(self.show_data)
        # self.btn_export.clicked.connect(self.export_data)
        self.show()

    def submit(self):
        # get values from input fields
        # if (self.sn.text() != "" and self.xcor.text() != "" and self.ycor.text() != "" and self.species.currentText() != "" and self.gd.currentText() != "" and self.girthdiam.text() != "" and self.t_height.text() != "" and self.tree_class.currentText() != "" and self.remark.text() != ""):
        # get values from input fields
        sn = self.sn.text()
        # crs = self.crs.currentText()
        xcor = self.xcor.text()
        ycor = self.ycor.text()
        species = self.species.currentText()
        gd = self.gd.currentText()
        girthdiam = self.girthdiam.text()
        t_height = self.t_height.text()
        tree_class = self.tree_class.currentText()
        remark = self.remark.text()

        # Check if any input field is empty
        # if not all([sn, crs, species, gd, girthdiam, t_height, tree_class]):
            # # Check if the 'S.N.' value is already in the dataframe
            # if sn in self.df['S.N.'].values:
            # # Display a message if the 'S.N.' value is already in the dataframe
            #     QMessageBox.warning(self, 'Error', 'This serial number of tree value is already entered!')
            # else:
        if not all([sn, girthdiam, t_height]):
            # Display an error message if any input field is empty
            QMessageBox.warning(self, "Error", "Please fill up all input fields!")
        else:
            self.row_num += 1
            self.main_df.loc[self.row_num, 'S.N.'] = sn
            self.main_df.loc[self.row_num, 'X-Coordinate'] = xcor
            self.main_df.loc[self.row_num, 'Y-Coordinate'] = ycor
            self.main_df.loc[self.row_num, 'प्रजाती (Species)'] = species
            self.main_df.loc[self.row_num, 'मापनः गोलाई/ब्यास'] = gd
            self.main_df.loc[self.row_num, 'गोलाई/ब्यास (cm)'] = girthdiam
            self.main_df.loc[self.row_num, 'उचाई (Height) m.'] = t_height
            self.main_df.loc[self.row_num, 'श्रेणी/दर्जा (Tree class)'] = tree_class
            self.main_df.loc[self.row_num, 'कैफियत'] = remark

            # merge the df with the df_par
            self.df = pd.merge(self.main_df, self.df_par, on='प्रजाती (Species)', how='left')

            self.df = self.df[['S.N.', 'X-Coordinate', 'Y-Coordinate', 'प्रजाती (Species)','मापनः गोलाई/ब्यास','गोलाई/ब्यास (cm)','उचाई (Height) m.', 'श्रेणी/दर्जा (Tree class)', 'कैफियत', 'a', 'b', 'c', 'a1', 'b1', 'small', 'medium', 'big']]
            # self.df[['Coordinate System']] = self.df[['Coordinate System']].apply(pd.to_numeric)
            self.df[['X-Coordinate']] = self.df[['X-Coordinate']].apply(pd.to_numeric)
            self.df[['Y-Coordinate']] = self.df[['Y-Coordinate']].apply(pd.to_numeric)
            self.df[['गोलाई/ब्यास (cm)']] = self.df[['गोलाई/ब्यास (cm)']].apply(pd.to_numeric)
            self.df[['उचाई (Height) m.']] = self.df[['उचाई (Height) m.']].apply(pd.to_numeric)
            # self.df[['X-Coordinate', 'Y-Coordinate', 'गोलाई/ब्यास (cm)','उचाई (Height) m.', 'a', 'b', 'c', 'a1', 'b1', 'small', 'medium', 'big']] = self.df[['गोलाई/ब्यास (cm)','उचाई (Height) m.', 'a', 'b', 'c', 'a1', 'b1', 'small', 'medium', 'big']].apply(pd.to_numeric)
        
            # Define ब्यास (Diameter) cm
            def categorise(row):
                    if row['मापनः गोलाई/ब्यास'] == 'गोलाई (Girth)':
                        return row['गोलाई/ब्यास (cm)']/(np.pi)
                    else:
                        return row['गोलाई/ब्यास (cm)']

            # Apply function
            self.df['ब्यास (Diameter) cm'] = self.df.apply(categorise, axis=1)
            self.df[['ब्यास (Diameter) cm']] = self.df[['ब्यास (Diameter) cm']].apply(pd.to_numeric)
            
            # Define stem volume
            # Steam_Volume = (np.exp(self.df['a']+(self.df['b']*np.log(float(self.df['ब्यास (Diameter) cm'])))+(self.df['c']*np.log(float(self.df['उचाई (Height) m.']))))/1000)
            Steam_Volume = (np.exp(self.df['a']+(self.df['b']*np.log(self.df['ब्यास (Diameter) cm']))+(self.df['c']*np.log(self.df['उचाई (Height) m.'])))/1000)
            self.df['Steam Volume (m3)'] = Steam_Volume
            
            # Define R-Value
            def categorise(row):
                if row['ब्यास (Diameter) cm'] <= 10:
                    return row['small']
                elif row['ब्यास (Diameter) cm'] < 40:
                    return (((row['ब्यास (Diameter) cm']-10)*row['medium']+(40-row['ब्यास (Diameter) cm'])*row['small'])/30)
                elif row['ब्यास (Diameter) cm'] <= 70:
                    return (((row['ब्यास (Diameter) cm']-40)*row['big']+(70-row['ब्यास (Diameter) cm'])*row['medium'])/30)
                elif row['ब्यास (Diameter) cm'] > 70:
                    return row['big']
                else:
                    return 0

            # Apply function
            self.df['R-Value'] = self.df.apply(categorise, axis=1)
            
            # Define '<10 cm DBH Ratio'
            Top_diam_ratio_10cm = (np.exp(self.df['a1']+(self.df['b1']*np.log(self.df['ब्यास (Diameter) cm']))))
            self.df['<10 cm DBH Ratio'] = Top_diam_ratio_10cm

            # Define branch volume
            Branch_Volume = (self.df['Steam Volume (m3)'] * self.df['R-Value'])
            self.df['Branch Volume (m3)'] = Branch_Volume

            # Define branch volume
            Total_Volume = (self.df['Steam Volume (m3)'] + self.df['Branch Volume (m3)'])
            self.df['Tree Volume (m3)'] = Total_Volume

            # Define Gross Timber Volume
            Gross_Timber_Volume = (self.df['Steam Volume (m3)'] - (self.df['Steam Volume (m3)']*self.df['<10 cm DBH Ratio']))
            self.df['Gross Timber Volume (m3)'] = Gross_Timber_Volume
            
            # df.apply() method for Volume with Classes (I-IV)
            # Class I
            # Define function
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '१': 
                    return row['Gross Timber Volume (m3)']*80/100
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '१':
                    return 0

            # Apply function
            self.df['I Class Wood'] = self.df.apply(categorise, axis=1) #self.df.apply(lambda row: categorise(row), axis=1)
                                            
            # Fill in any missing values with 0
            self.df['I Class Wood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['I Class Wood'] = pd.to_numeric(self.df['I Class Wood'])
            
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '१':
                    return row['Tree Volume (m3)'] - row['I Class Wood']
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '१':
                    return 0 

            # Apply function
            self.df['I Class Fuelwood'] = self.df.apply(categorise, axis=1)
            
            # Fill in any missing values with 0
            self.df['I Class Fuelwood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['I Class Fuelwood'] = pd.to_numeric(self.df['I Class Fuelwood'])

            # Class II
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '२': 
                    return row['Gross Timber Volume (m3)']*60/100
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '२':
                    return 0

            # Apply function    
            self.df['II Class Wood'] = self.df.apply(categorise, axis=1)

            # Fill in any missing values with 0
            self.df['II Class Wood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['II Class Wood'] = pd.to_numeric(self.df['II Class Wood'])
            
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '२':
                    return row['Tree Volume (m3)'] - row['II Class Wood']
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '२':
                    return 0

            # Apply function
            self.df['II Class Fuelwood'] = self.df.apply(categorise, axis=1)
            # Fill in any missing values with 0
            self.df['II Class Fuelwood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['II Class Fuelwood'] = pd.to_numeric(self.df['II Class Fuelwood'])

            # Class III
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '३': 
                    return row['Gross Timber Volume (m3)']*30/100
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '३':
                    return 0

            # Apply function
            self.df['III Class Wood'] = self.df.apply(categorise, axis=1)
            
            # Fill in any missing values with 0
            self.df['III Class Wood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['III Class Wood'] = pd.to_numeric(self.df['III Class Wood'])

            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '३':
                    return row['Tree Volume (m3)'] - row['III Class Wood']
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '३':
                    return 0

            # Apply function
            self.df['III Class Fuelwood'] = self.df.apply(categorise, axis=1)

            # Fill in any missing values with 0
            self.df['III Class Fuelwood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['III Class Fuelwood'] = pd.to_numeric(self.df['III Class Fuelwood'])

            # Class IV
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '४': 
                    return row['Gross Timber Volume (m3)']*100/100
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '४':
                    return 0

            # Apply function
            self.df['IV Class Fuelwood'] = self.df.apply(categorise, axis=1)

            # Fill in any missing values with 0
            self.df['IV Class Fuelwood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['IV Class Fuelwood'] = pd.to_numeric(self.df['IV Class Fuelwood'])
            
            # Define net volume for other than 'खयर'
            Net_Volume_others = (self.df['I Class Wood'] + self.df['II Class Wood'] + self.df['III Class Wood'])
            self.df['Net_Vol_others'] = Net_Volume_others
            
            #Function of Net volume (m3)
            def categorise(row):                                 
                if row['प्रजाती (Species)'] == 'खयर':
                    return row['Steam Volume (m3)']
                else:
                    return row['Net_Vol_others']

            #Apply function
            self.df['Net Volume (m3)'] = self.df.apply(categorise, axis=1)
            
            # Define Total Fuelwood
            Total_Fuelwood = (self.df['Tree Volume (m3)'] - self.df['Net Volume (m3)'])
            self.df['Total Fuelwood (m3)'] = Total_Fuelwood

            # Define Net Volume in Cubic Feet
            Net_Volume_Cuft = (self.df['Net Volume (m3)'] *35.3147)
            self.df['Net Volume (cuft)'] = Net_Volume_Cuft

            # Define Total Fuelwood in Chatta (चट्टा)
            Total_Fuelwood_Chatta = (self.df['Total Fuelwood (m3)'] *0.105944)
            self.df['Fuelwood (चट्टा)'] = Total_Fuelwood_Chatta
            
            # drop and copy self.df as new df for detailed calculation sheet
            self.df = self.df[['S.N.', 'X-Coordinate', 'Y-Coordinate', 'प्रजाती (Species)','मापनः गोलाई/ब्यास','गोलाई/ब्यास (cm)', 'a','b','c','a1','b1','small','medium','big','ब्यास (Diameter) cm',
                        'उचाई (Height) m.','श्रेणी/दर्जा (Tree class)','Steam Volume (m3)','Branch Volume (m3)', 
                        'Tree Volume (m3)', 'Gross Timber Volume (m3)', 'I Class Wood', 'I Class Fuelwood', 'II Class Wood',
                        'II Class Fuelwood','III Class Wood', 'III Class Fuelwood', 'IV Class Fuelwood', 'Net Volume (m3)', 'Net Volume (cuft)',
                        'Total Fuelwood (m3)', 'Fuelwood (चट्टा)', 'कैफियत']]
            self.df.drop(columns=['मापनः गोलाई/ब्यास','गोलाई/ब्यास (cm)', 'a','b','c','a1','b1','small','medium','big'], inplace=True)
            
            # do round down by 3 here
            self.df[['Steam Volume (m3)','Branch Volume (m3)', 'Tree Volume (m3)', 'Gross Timber Volume (m3)', 'I Class Wood', 'I Class Fuelwood', 'II Class Wood',
                        'II Class Fuelwood','III Class Wood', 'III Class Fuelwood', 'IV Class Fuelwood', 'Net Volume (m3)', 'Net Volume (cuft)',
                        'Total Fuelwood (m3)', 'Fuelwood (चट्टा)']] = self.df[['Steam Volume (m3)','Branch Volume (m3)', 'Tree Volume (m3)', 'Gross Timber Volume (m3)', 'I Class Wood', 'I Class Fuelwood', 'II Class Wood',
                        'II Class Fuelwood','III Class Wood', 'III Class Fuelwood', 'IV Class Fuelwood', 'Net Volume (m3)', 'Net Volume (cuft)',
                        'Total Fuelwood (m3)', 'Fuelwood (चट्टा)']].astype(float).round(3)#.astype(int)

            self.df.rename(columns = {'S.N.':'रुख नं.', 'प्रजाती (Species)':'प्रजाति', 'ब्यास (Diameter) cm':'ब्यास (Diameter) (से.मि.)', 
                        'उचाई (Height) m.':'उचाई (मि.)', 'श्रेणी/दर्जा (Tree class)':'श्रेणी/दर्जा', 'Steam Volume (m3)':'कान्डको आयतन (घ.मि.)',
                        'Branch Volume (m3)':'हाँगाको आयतन (घ.मि.)', 'Tree Volume (m3)':'रुखको आयतन (घ.मि.)',
                        'Gross Timber Volume (m3)':'ग्रस आयतन (घ.मि.)', 'Net Volume (m3)':'नेट आयतन (घ.मि.)',
                        'Net Volume (cuft)':'नेट आयतन (घन फिट)', 'Total Fuelwood (m3)':'दाउराको आयतन (घ.मि.)',
                        'Fuelwood (चट्टा)':'दाउरा (चट्टा)'}, inplace = True)


                
            # clear forms input fields
            self.sn.setText('')
            self.species.setCurrentText('0')
            def dropdown_item_changed(self, index):
                # Store the index of the selected item
                self.last_selected_index = index
                # Connect the itemChanged signal of the dropdown list to the handler function
                self.dropdown.itemChanged.connect(self.dropdown_item_changed)
                self.gd.setCurrentIndex(self.last_selected_index)
            self.girthdiam.setText('')
            self.t_height.setText('')
            self.tree_class.setCurrentText('0')
            self.remark.setText('')

            #Update button state based on submit button
            self.btn_show_table.setEnabled(True)
            # connect signal to a function
            self.btn_show_table.clicked.connect(self.show_data)
            self.df.to_excel("data_TVAS.xlsx")
        
    def show_data(self):
        # update table view with dataframe
        self.model = PandasModel(self.df)
        self.table_view.setModel(self.model)
        self.table_view.setFont(QFont('Kalimati', 11))
        # ResizeToContents as per headers length
        self.table_view.resizeColumnsToContents()

        self.btn_delete.setEnabled(False) # initially set delete button to unclickable
        self.btn_export.setEnabled(True)
        self.btn_delete.clicked.connect(self.delete_data)
        self.btn_export.clicked.connect(self.export_data)
        self.table_view.selectionModel().selectionChanged.connect(self.enable_delete) # connect selection changed signal to enable_delete function

    def enable_delete(self):
        if self.table_view.selectionModel().hasSelection():
            self.btn_delete.setEnabled(True) # enable delete button if a row is selected
        else:
            self.btn_delete.setEnabled(False) # disable delete button if no row is selected

    def delete_data(self):
        indexes = self.table_view.selectionModel().selectedRows()
        if not indexes:
            self.btn_delete.setEnabled(False)
            # QMessageBox.warning(self, 'Error', 'Please select a row to delete.')
        else:
            for index in sorted(indexes, reverse=True):
                response = QMessageBox.question(self, 'Caution!',
                                       'Are you sure you want to delete the row?',
                                       QMessageBox.Yes | QMessageBox.No,
                                       QMessageBox.No)
                if response == QMessageBox.Yes:
                    self.main_df = self.main_df.drop(self.main_df.index[index.row()])
            self.df = pd.merge(self.main_df, self.df_par, on='प्रजाती (Species)', how='left')
            self.df = self.df[['S.N.', 'X-Coordinate', 'Y-Coordinate', 'प्रजाती (Species)','मापनः गोलाई/ब्यास','गोलाई/ब्यास (cm)','उचाई (Height) m.', 'श्रेणी/दर्जा (Tree class)', 'कैफियत', 'a', 'b', 'c', 'a1', 'b1', 'small', 'medium', 'big']]
            # self.df[['Coordinate System']] = self.df[['Coordinate System']].apply(pd.to_numeric)
            self.df[['X-Coordinate']] = self.df[['X-Coordinate']].apply(pd.to_numeric)
            self.df[['Y-Coordinate']] = self.df[['Y-Coordinate']].apply(pd.to_numeric)
            self.df[['गोलाई/ब्यास (cm)']] = self.df[['गोलाई/ब्यास (cm)']].apply(pd.to_numeric)
            self.df[['उचाई (Height) m.']] = self.df[['उचाई (Height) m.']].apply(pd.to_numeric)
            # self.df[['X-Coordinate', 'Y-Coordinate', 'गोलाई/ब्यास (cm)','उचाई (Height) m.', 'a', 'b', 'c', 'a1', 'b1', 'small', 'medium', 'big']] = self.df[['गोलाई/ब्यास (cm)','उचाई (Height) m.', 'a', 'b', 'c', 'a1', 'b1', 'small', 'medium', 'big']].apply(pd.to_numeric)
        
            # Define ब्यास (Diameter) cm
            def categorise(row):
                    if row['मापनः गोलाई/ब्यास'] == 'गोलाई (Girth)':
                        return row['गोलाई/ब्यास (cm)']/(np.pi)
                    else:
                        return row['गोलाई/ब्यास (cm)']

            # Apply function
            self.df['ब्यास (Diameter) cm'] = self.df.apply(categorise, axis=1)
            self.df[['ब्यास (Diameter) cm']] = self.df[['ब्यास (Diameter) cm']].apply(pd.to_numeric)
            
            # Define stem volume
            # Steam_Volume = (np.exp(self.df['a']+(self.df['b']*np.log(float(self.df['ब्यास (Diameter) cm'])))+(self.df['c']*np.log(float(self.df['उचाई (Height) m.']))))/1000)
            Steam_Volume = (np.exp(self.df['a']+(self.df['b']*np.log(self.df['ब्यास (Diameter) cm']))+(self.df['c']*np.log(self.df['उचाई (Height) m.'])))/1000)
            self.df['Steam Volume (m3)'] = Steam_Volume
            
            # Define R-Value
            def categorise(row):
                if row['ब्यास (Diameter) cm'] <= 10:
                    return row['small']
                elif row['ब्यास (Diameter) cm'] < 40:
                    return (((row['ब्यास (Diameter) cm']-10)*row['medium']+(40-row['ब्यास (Diameter) cm'])*row['small'])/30)
                elif row['ब्यास (Diameter) cm'] <= 70:
                    return (((row['ब्यास (Diameter) cm']-40)*row['big']+(70-row['ब्यास (Diameter) cm'])*row['medium'])/30)
                elif row['ब्यास (Diameter) cm'] > 70:
                    return row['big']
                else:
                    return 0

            # Apply function
            self.df['R-Value'] = self.df.apply(categorise, axis=1)
            
            # Define '<10 cm DBH Ratio'
            Top_diam_ratio_10cm = (np.exp(self.df['a1']+(self.df['b1']*np.log(self.df['ब्यास (Diameter) cm']))))
            self.df['<10 cm DBH Ratio'] = Top_diam_ratio_10cm

            # Define branch volume
            Branch_Volume = (self.df['Steam Volume (m3)'] * self.df['R-Value'])
            self.df['Branch Volume (m3)'] = Branch_Volume

            # Define branch volume
            Total_Volume = (self.df['Steam Volume (m3)'] + self.df['Branch Volume (m3)'])
            self.df['Tree Volume (m3)'] = Total_Volume

            # Define Gross Timber Volume
            Gross_Timber_Volume = (self.df['Steam Volume (m3)'] - (self.df['Steam Volume (m3)']*self.df['<10 cm DBH Ratio']))
            self.df['Gross Timber Volume (m3)'] = Gross_Timber_Volume
            
            # df.apply() method for Volume with Classes (I-IV)
            # Class I
            # Define function
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '१': 
                    return row['Gross Timber Volume (m3)']*80/100
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '१':
                    return 0

            # Apply function
            self.df['I Class Wood'] = self.df.apply(categorise, axis=1) #self.df.apply(lambda row: categorise(row), axis=1)
                                            
            # Fill in any missing values with 0
            self.df['I Class Wood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['I Class Wood'] = pd.to_numeric(self.df['I Class Wood'])
            
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '१':
                    return row['Tree Volume (m3)'] - row['I Class Wood']
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '१':
                    return 0 

            # Apply function
            self.df['I Class Fuelwood'] = self.df.apply(categorise, axis=1)
            
            # Fill in any missing values with 0
            self.df['I Class Fuelwood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['I Class Fuelwood'] = pd.to_numeric(self.df['I Class Fuelwood'])

            # Class II
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '२': 
                    return row['Gross Timber Volume (m3)']*60/100
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '२':
                    return 0

            # Apply function    
            self.df['II Class Wood'] = self.df.apply(categorise, axis=1)

            # Fill in any missing values with 0
            self.df['II Class Wood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['II Class Wood'] = pd.to_numeric(self.df['II Class Wood'])
            
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '२':
                    return row['Tree Volume (m3)'] - row['II Class Wood']
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '२':
                    return 0

            # Apply function
            self.df['II Class Fuelwood'] = self.df.apply(categorise, axis=1)
            # Fill in any missing values with 0
            self.df['II Class Fuelwood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['II Class Fuelwood'] = pd.to_numeric(self.df['II Class Fuelwood'])

            # Class III
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '३': 
                    return row['Gross Timber Volume (m3)']*30/100
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '३':
                    return 0

            # Apply function
            self.df['III Class Wood'] = self.df.apply(categorise, axis=1)
            
            # Fill in any missing values with 0
            self.df['III Class Wood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['III Class Wood'] = pd.to_numeric(self.df['III Class Wood'])

            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '३':
                    return row['Tree Volume (m3)'] - row['III Class Wood']
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '३':
                    return 0

            # Apply function
            self.df['III Class Fuelwood'] = self.df.apply(categorise, axis=1)

            # Fill in any missing values with 0
            self.df['III Class Fuelwood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['III Class Fuelwood'] = pd.to_numeric(self.df['III Class Fuelwood'])

            # Class IV
            def categorise(row):
                if row['प्रजाती (Species)'] != 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '४': 
                    return row['Gross Timber Volume (m3)']*100/100
                elif row['प्रजाती (Species)'] == 'खयर' and row['श्रेणी/दर्जा (Tree class)'] == '४':
                    return 0

            # Apply function
            self.df['IV Class Fuelwood'] = self.df.apply(categorise, axis=1)

            # Fill in any missing values with 0
            self.df['IV Class Fuelwood'].fillna(0, inplace=True)
            # Convert the column to a numeric type
            self.df['IV Class Fuelwood'] = pd.to_numeric(self.df['IV Class Fuelwood'])
            
            # Define net volume for other than 'खयर'
            Net_Volume_others = (self.df['I Class Wood'] + self.df['II Class Wood'] + self.df['III Class Wood'])
            self.df['Net_Vol_others'] = Net_Volume_others
            
            #Function of Net volume (m3)
            def categorise(row):                                 
                if row['प्रजाती (Species)'] == 'खयर':
                    return row['Steam Volume (m3)']
                else:
                    return row['Net_Vol_others']

            #Apply function
            self.df['Net Volume (m3)'] = self.df.apply(categorise, axis=1)
            
            # Define Total Fuelwood
            Total_Fuelwood = (self.df['Tree Volume (m3)'] - self.df['Net Volume (m3)'])
            self.df['Total Fuelwood (m3)'] = Total_Fuelwood

            # Define Net Volume in Cubic Feet
            Net_Volume_Cuft = (self.df['Net Volume (m3)'] *35.3147)
            self.df['Net Volume (cuft)'] = Net_Volume_Cuft

            # Define Total Fuelwood in Chatta (चट्टा)
            Total_Fuelwood_Chatta = (self.df['Total Fuelwood (m3)'] *0.105944)
            self.df['Fuelwood (चट्टा)'] = Total_Fuelwood_Chatta
            
            # drop and copy self.df as new df for detailed calculation sheet
            self.df = self.df[['S.N.', 'X-Coordinate', 'Y-Coordinate', 'प्रजाती (Species)','मापनः गोलाई/ब्यास','गोलाई/ब्यास (cm)', 'a','b','c','a1','b1','small','medium','big','ब्यास (Diameter) cm',
                        'उचाई (Height) m.','श्रेणी/दर्जा (Tree class)','Steam Volume (m3)','Branch Volume (m3)', 
                        'Tree Volume (m3)', 'Gross Timber Volume (m3)', 'I Class Wood', 'I Class Fuelwood', 'II Class Wood',
                        'II Class Fuelwood','III Class Wood', 'III Class Fuelwood', 'IV Class Fuelwood', 'Net Volume (m3)', 'Net Volume (cuft)',
                        'Total Fuelwood (m3)', 'Fuelwood (चट्टा)', 'कैफियत']]
            self.df.drop(columns=['मापनः गोलाई/ब्यास','गोलाई/ब्यास (cm)', 'a','b','c','a1','b1','small','medium','big'], inplace=True)
            
            # do round down by 3 here
            self.df[['Steam Volume (m3)','Branch Volume (m3)', 'Tree Volume (m3)', 'Gross Timber Volume (m3)', 'I Class Wood', 'I Class Fuelwood', 'II Class Wood',
                        'II Class Fuelwood','III Class Wood', 'III Class Fuelwood', 'IV Class Fuelwood', 'Net Volume (m3)', 'Net Volume (cuft)',
                        'Total Fuelwood (m3)', 'Fuelwood (चट्टा)']] = self.df[['Steam Volume (m3)','Branch Volume (m3)', 'Tree Volume (m3)', 'Gross Timber Volume (m3)', 'I Class Wood', 'I Class Fuelwood', 'II Class Wood',
                        'II Class Fuelwood','III Class Wood', 'III Class Fuelwood', 'IV Class Fuelwood', 'Net Volume (m3)', 'Net Volume (cuft)',
                        'Total Fuelwood (m3)', 'Fuelwood (चट्टा)']].astype(float).round(3)#.astype(int)

            self.df.rename(columns = {'S.N.':'रुख नं.', 'प्रजाती (Species)':'प्रजाति', 'ब्यास (Diameter) cm':'ब्यास (Diameter) (से.मि.)', 
                        'उचाई (Height) m.':'उचाई (मि.)', 'श्रेणी/दर्जा (Tree class)':'श्रेणी/दर्जा', 'Steam Volume (m3)':'कान्डको आयतन (घ.मि.)',
                        'Branch Volume (m3)':'हाँगाको आयतन (घ.मि.)', 'Tree Volume (m3)':'रुखको आयतन (घ.मि.)',
                        'Gross Timber Volume (m3)':'ग्रस आयतन (घ.मि.)', 'Net Volume (m3)':'नेट आयतन (घ.मि.)',
                        'Net Volume (cuft)':'नेट आयतन (घन फिट)', 'Total Fuelwood (m3)':'दाउराको आयतन (घ.मि.)',
                        'Fuelwood (चट्टा)':'दाउरा (चट्टा)'}, inplace = True)
            self.show_data()

        # ResizeToContents as per headers length
        self.table_view.resizeColumnsToContents()

    # Define export function
    def export_data(self):  #def export_data(self, df3_1, df4=None, df6=None): #def export_data(self): 
        # if not self.df.empty:
        # Apply the export function with updated dataframe                                  
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getSaveFileName(self, "QFileDialog.getSaveFileName()", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        if file_name:
            if self.df.shape[0] > 0:
            # Create an Excel writer using the xlsxwriter engine
                writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

            # Write each dataframe to a different worksheet
            # self.df = self.df.style.set_properties(**{'border-color': 'black', 'border-width': '1px', 'border-style': 'solid'})
                self.df.to_excel(writer, sheet_name='Detailed calculation sheet', index=False)

            # subset_df.to_excel(writer, sheet_name='Brief data', index=False)
            # subset_df = subset_df.style.set_properties(**{'border-color': 'black', 'border-width': '1px', 'border-style': 'solid'})

            # self.df3 = self.df3.style.set_properties(**{'border-color': 'black', 'border-width': '1px', 'border-style': 'solid'})
            # self.df.to_excel(writer, sheet_name='Summary Report', startrow=11, startcol=0, index=True)

            # workbook = writer.book
            # worksheet1 = writer.sheets['Detailed calculation sheet']
            # # worksheet2 = writer.sheets['Brief data']
            # worksheet3 = writer.sheets['Summary Report']

            # Create the pivot table
                table = pd.pivot_table(self.df, index='प्रजाति', values=['कान्डको आयतन (घ.मि.)','हाँगाको आयतन (घ.मि.)','रुखको आयतन (घ.मि.)','ग्रस आयतन (घ.मि.)','नेट आयतन (घ.मि.)','नेट आयतन (घन फिट)', 'दाउराको आयतन (घ.मि.)', 'दाउरा (चट्टा)'], aggfunc='sum')

                # Add grand total
                table.loc['Grand Total'] = table.sum()

                # Add 'status_count' column to pivot table
                table['रुख संख्या'] = self.df.प्रजाति.value_counts()

                table.to_excel(writer, sheet_name='Summary Report', startrow=11, index=True)

                worksheet1 = writer.sheets['Detailed calculation sheet']
                worksheet3 = writer.sheets['Summary Report']

                # Set the font for the worksheet1
                # Set the font and border for the entire worksheet
                font_w = writer.book.add_format({'font_name': 'Kalimati'})
                border = writer.book.add_format({'border':1})
                worksheet1.set_column('A:W', None, font_w)
                worksheet1.set_column('A:W', None, border)
                worksheet3.set_column('A:J', None, font_w)
                worksheet3.set_column('A:J', None, border)

                # set paper size and layout
                worksheet1.set_paper(9)  # A4 
                worksheet1.set_landscape() # Landscape page layout
                # worksheet2.set_paper(9)
                # worksheet2.set_landscape() # Landscape page layout
                worksheet3.set_paper(9)
                worksheet3.set_portrait() # portrait page layout

                # Set the column width to match the header width for each sheet
                for i, width in enumerate(self.df.columns.str.len()):
                    worksheet1.set_column(i, i, width)
                # for i, width in enumerate(subset_df.columns.str.len()):
                #     worksheet2.set_column(i, i, width)
                for i, width in enumerate(self.df.columns.str.len()):
                    worksheet3.set_column(i, i, width)

                # self.df = self.df.style.set_properties(**{'border-color': 'black', 'border-width': '1px', 'border-style': 'solid'})

                # Merge cells 'A1:J1' for Main Heading using xlsxwriter
                # worksheet3.write('A1:J1', 'रुखको आयतन मुल्याङ्कन सारांस विवरण')
                merge_format = writer.book.add_format()
                merge_format.set_align('center')
                merge_format.set_align('Vcenter')
                merge_format.set_underline()
                merge_format.set_bold()
                merge_format.set_font_name('Kalimati')
                merge_format.set_font_size(12)
                # Merge the cells
                worksheet3.merge_range('A1:J1', 'रुखको आयतन मुल्याङ्कन सारांस विवरण', merge_format)

                # autofit columns by width
                worksheet1.set_column('A:W', None, None, {'autofit': True})
                worksheet3.set_column('A11:J11', None, None, {'autofit': True})

                # Add an Excel date format.
                date_format = writer.book.add_format({'num_format': 'yyyy/mm/dd', 'font_name': 'Kalimati','font_size':11,'bold':False})

                # write_blank() format
                cell_format = writer.book.add_format({'bold': False})
                worksheet3.set_column(0, 0, None, cell_format)

                # cell properties format for auto adjusting cells lengths
                cell_font = writer.book.add_format({'font_name': 'Kalimati','font_size':11,'bold':False})
                times_font = writer.book.add_format({'font_name': 'Times New Roman','font_size':12,'bold':False})

                # Write the labels to specific cells.
                worksheet3.write('A3', 'प्रदेशः', cell_font)
                worksheet3.write('A4', 'कार्यालयको नामः', cell_font)
                worksheet3.write('A5', 'वनको नामः', cell_font)
                worksheet3.write('A6', 'वडा नं.:', cell_font)
                worksheet3.write('A7', 'उप-खण्डः', cell_font)
                worksheet3.write('A8', 'Coordinate System:', times_font)
                worksheet3.write('A9', 'तथ्याङ्क संकलन मितिः', cell_font)
                worksheet3.write('E3', 'जिल्लाः', cell_font)
                worksheet3.write('E4', 'सव-डि.व.का./रे.पो.:', cell_font)
                worksheet3.write('E5', 'गा.पा./न.पा.:', cell_font)
                worksheet3.write('E6', 'खण्ड:', cell_font)
                worksheet3.write('E7', 'प्लट नं.:', cell_font)
                worksheet3.write('E8', 'आर्थिक वर्षः', cell_font)
                worksheet3.write('E9', 'तथ्याङ्क संकलकको नाम, पदः', cell_font)

                # Write the df_gi dataframe column values to specific cells of df5.
                if self.df_gi.empty:
                    worksheet3.write('B3', None, cell_format) # write_blank() 
                    worksheet3.write('B4', None, cell_format)
                    worksheet3.write('B5', None, cell_format)
                    worksheet3.write('B6', None, cell_format)
                    worksheet3.write('B7', None, cell_format)
                    worksheet3.write('B8', None, cell_format)
                    worksheet3.write('B9', None, cell_format)

                    worksheet3.write('F3', None, cell_format)
                    worksheet3.write('F4', None, cell_format)
                    worksheet3.write('F5', None, cell_format)
                    worksheet3.write('F6', None, cell_format)
                    worksheet3.write('F7', None, cell_format)
                    worksheet3.write('F8', None, cell_format)
                    worksheet3.write('F9', None, cell_format)
                    # raise warning message
                    QMessageBox.information(self,'Warning','Data Exported: Informations are missing from General Information window to generate the Summary Report')
                else:
                    worksheet3.write('B3', self.df_gi.iloc[0]['प्रदेश'], cell_font)
                    worksheet3.write('B4', self.df_gi.iloc[0]['कार्यालयको नाम'], cell_font)
                    worksheet3.write('B5', self.df_gi.iloc[0]['वनको नाम'], cell_font)
                    worksheet3.write('B6', self.df_gi.iloc[0]['वडा नं.'], cell_font)
                    worksheet3.write('B7', self.df_gi.iloc[0]['उप-खण्ड'], cell_font)
                    worksheet3.write('B8', self.df_gi.iloc[0]['Coordinate System'], times_font)
                    worksheet3.write('B9', self.df_gi.iloc[0]['तथ्यांक संकलन मिति'], date_format)

                    worksheet3.write('F3', self.df_gi.iloc[0]['जिल्ला'], cell_font)
                    worksheet3.write('F4', self.df_gi.iloc[0]['सव-डि.व.का./रे.पो.'], cell_font)
                    worksheet3.write('F5', self.df_gi.iloc[0]['गा.पा./न.पा.'], cell_font)
                    worksheet3.write('F6', self.df_gi.iloc[0]['खण्ड'], cell_font)
                    worksheet3.write('F7', self.df_gi.iloc[0]['प्लट नं.'], cell_font)
                    worksheet3.write('F8', self.df_gi.iloc[0]['आर्थिक वर्ष'], cell_font)
                    worksheet3.write('F9', self.df_gi.iloc[0]['तथ्यांक संकलकको नाम, पद'], cell_font)

                    # raise success message
                    QMessageBox.information(self,'Success','Data Exported Succesfully!')
                            # Save the Excel file
                # writer.save()
                writer.close()
        # else:
        #     QMessageBox.warning(self, 'Warning', 'Table has no values to export')

class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parnet=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])

            column_count = self.columnCount()
            for column in range(0, column_count):
                if (index.column() == column and role == Qt.TextAlignmentRole):
                    value = self._data.iloc[index.row(), index.column()]
                    if isinstance(value, (int, float)):
                        return Qt.AlignRight | Qt.AlignVCenter
                    else:
                        return Qt.AlignHCenter | Qt.AlignVCenter

        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None

# Run main window
if __name__ == "__main__":
    sys.stdout = open(os.devnull, "w")
    sys.stderr = open(os.devnull, "w")
    app = QApplication(sys.argv)
    font = QFont("Times New Roman", 12)
    app.setFont(font)
    main = Nepal_TVAS()
    main.setStyleSheet("background-color: Dark") 
    main.show()
    sys.exit(app.exec_())

### -------------------------------------------------------------- Contribute here ... ### 
