#standard library imports
import os
import sys
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime

import xlwings as xw

from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import QApplication, QLabel, QDialog,\
                            QLineEdit, QGridLayout,QPushButton, QComboBox

#import all the path class containg all the paths
from parameters import paths
from aux_functions import source_sps_to_df, get_line_df, get_matrix_stats,\
                          deployed_sps_to_df, recovered_sps_to_df, get_rcv_line_df, get_line_matrix_stats
from extra_functions import get_file_stats

__version__ = "prod. 1.0.7" #April 2023

#icon here
NEO = 'Neo.jpg'

#columns names
rcv_lines_col = "A"
preplot_rcv_count_col = "B"
dep_count_col = "C"
dep_start_col = "D"
dep_stop_col = "E"

rec_count_col = "F"
rec_start_col = "G"
rec_stop_col = "H"


src_lines_col = "K"
preplot_src_count_col = "L"
src_count_col = "M"
src_start_col = "N"
src_stop_col = "O"

#colors
accept_color = "#82E0AA"
alert_color = "#922B21"
info_color = "#27AE60"
warning_color = "#F4D03F"

no_color = "#FFFFFF"

#main Dialog window
class DumbDialog(QDialog):
    def __init__(self,parent = None):
        super(DumbDialog,self).__init__(parent, flags = Qt.WindowMinimizeButtonHint|Qt.WindowCloseButtonHint)
        optionLabel = QLabel(f"<font color={info_color}><b>Select Line type:</font></b>")
        self.optionBox = QComboBox()
        self.optionBox.addItems(["Deployed","Recovered","Source"])
        self.lineInput = QLineEdit()
        self.lineInput.setFocus()

        self.checkButton = QPushButton("Check")
        self.updateButton = QPushButton("Update Matrix")
        self.dataButton = QPushButton("Update SPS")         #update the input files
        self.updateLabel = QLabel("")
        self.updateLabel.setText(startUp_message)
        self.updateButton.setEnabled(False)
        # validator
        reg_ex = QRegExp(r"\d{3-4}")
        validator = QRegExpValidator(reg_ex, self.lineInput)
        self.lineInput.setValidator(QIntValidator())
        self.lineInput.setMaxLength(4)
        # layout
        grid = QGridLayout()
        grid.addWidget(optionLabel,0,0)
        grid.addWidget(self.optionBox,0, 1, 1, 3)
        grid.addWidget(self.lineInput, 1, 0, 1, 2)
        grid.addWidget(self.checkButton, 1, 2)
        grid.addWidget(self.updateButton, 1, 3)
        grid.addWidget(self.updateLabel, 2, 0, 1, 3)
        grid.addWidget(self.dataButton)
        
        # set icon
        icon = QIcon()
        icon.addFile(os.path.join(paths.app_image(), NEO))
        # set layout
        self.setLayout(grid)
        self.setMinimumSize(500,100)
        self.setMaximumSize(501,101)
        self.setWindowIcon(icon)
        self.setWindowTitle(f"Matrix helper {__version__}")

        # connecting slots
        self.optionBox.currentTextChanged.connect(self.updateUi)
        self.checkButton.clicked.connect(self.checkInput)
        self.updateButton.clicked.connect(self.applyUpdate)
        self.dataButton.clicked.connect(self.updateData)

        # set the data sources into the app scope
        # so it could be updated with the class instance method
        self.deployed_df = deployed_sps_to_df(sps_r_deploy)
        self.deployed_stats = get_file_stats(sps_r_deploy)

        self.recovered_df = recovered_sps_to_df(sps_r_recover)
        self.recovered_stats = get_file_stats(sps_r_recover)

        self.source_df = source_sps_to_df(sps_s_clean)
        self.source_stats = get_file_stats(sps_s_clean)

    def checkInput(self):
        '''prevalidation function'''
        option = self.optionBox.currentText()
        line_nb = self.lineInput.text()
        # general input check
        if not line_nb:
            message = f"<font color = {alert_color}><b>Please, enter line number</b></font>"
            self.updateLabel.setText(message)
            self.lineInput.setFocus()
            logger.warning("No value input attempt")
            return
        elif len(line_nb) < 3:
            message = f"<b>Incorrect value: <font color={alert_color}>{line_nb}</font></b>"
            self.updateLabel.setText(message)
            self.lineInput.selectAll()
            self.lineInput.setFocus()
            logger.warning(f"Short value {line_nb} input attempt")
            return
        else:
            if option == "Deployed":
                if int(line_nb) in self.deployed_df.line.unique():
                    message = f"Can update {option.lower()} line: <font color={info_color}><b>{line_nb}</font></b>"
                    self.updateLabel.setText(message)
                    self.updateButton.setEnabled(True)
                    self.updateButton.setFocus()
                    return
                else:
                    message = f"{option} line: <font color={alert_color}><b>{line_nb}</font></b> not found"
                    self.updateLabel.setText(message)
                    self.updateButton.setEnabled(False)
                    self.lineInput.selectAll()
                    self.lineInput.setFocus()
                    logger.warning(f"Non existing in {option} file line: {line_nb} input attempt")
                    return

            elif option == "Recovered":
                if self.recovered_df is not None:
                    if int(line_nb) in self.recovered_df.line.unique():
                        message = f"Can update {option.lower()} line: <font color={info_color}><b>{line_nb}</font></b>"
                        self.updateLabel.setText(message)
                        self.updateButton.setEnabled(True)
                        self.updateButton.setFocus()
                        return
                    else:
                        message = f"{option} line: <font color={alert_color}><b>{line_nb}</font></b> not found"
                        self.updateLabel.setText(message)
                        self.updateButton.setEnabled(False)
                        self.lineInput.selectAll()
                        self.lineInput.setFocus()
                        logger.warning(f"Non existing in {option} file line: {line_nb} input attempt")
                        return
                else:
                    message = f"<font color={alert_color}><b>{option} SPS file doesn't exist or empty</font></b>"
                    self.updateLabel.setText(message)
                    self.updateButton.setEnabled(False)
                    self.optionBox.setFocus()
                    logger.warning(f"Non existing {option} file access attempt")
                    return

            elif option == "Source":
                #if self.source_df != None:
                if int(line_nb) in self.source_df.line.unique():
                    message = f"Can update {option.lower()} line: <font color={info_color}><b>{line_nb}</font></b>"
                    self.updateLabel.setText(message)
                    self.updateButton.setEnabled(True)
                    self.updateButton.setFocus()
                    return
                else:
                    message = f"{option} line: <font color={alert_color}><b>{line_nb}</font></b> not found"
                    self.updateLabel.setText(message)
                    self.updateButton.setEnabled(False)
                    self.lineInput.selectAll()
                    self.lineInput.setFocus()
                    logger.warning(f"Non existing in {option} file line: {line_nb} input attempt")
                    return
                #else:
                #   message = f"<font color={alert_color}><b>{option} SPS file doesn't exist or empty</font></b>"
                #   self.updateLabel.setText(message)
                #   self.updateButton.setEnabled(False)
                #   self.optionBox.setFocus()
                #   logger.warning(f"Non existing {option} file access attempt")
                #   return

    def applyUpdate(self):
        '''function for updating matrix sheet'''
        option = self.optionBox.currentText()
        line_nb = self.lineInput.text()
        if option == "Deployed":
            try:
                line, dep_count, dep_start, dep_stop = \
                get_line_matrix_stats(get_rcv_line_df(self.deployed_df, int(line_nb), short = False))
                #update matrix file with attributes
                row = int(rlines_col.index(float(line_nb)))+1

                #first update with values then check the color
                ws[f"{dep_count_col}{row}"].value = dep_count
                if  ws[f"{dep_count_col}{row}"].value == ws[f"{preplot_rcv_count_col}{row}"].value:
                    ws[f"{dep_count_col}{row}"].color = f"{accept_color}"
                else:
                    ws[f"{dep_count_col}{row}"].color = f"{warning_color}"

                if  ws[f"{dep_start_col}{row}"].value != datetime.strptime(dep_start, "%Y-%m-%d %H:%M:%S"):
                    ws[f"{dep_start_col}{row}"].value = dep_start
                    ws[f"{dep_start_col}{row}"].color = f"{warning_color}"
                else:
                    ws[f"{dep_start_col}{row}"].color = f"{no_color}"
                
                if  ws[f"{dep_stop_col}{row}"].value != datetime.strptime(dep_stop, "%Y-%m-%d %H:%M:%S"):
                    ws[f"{dep_stop_col}{row}"].value = dep_stop
                    ws[f"{dep_stop_col}{row}"].color = f"{warning_color}"
                else:
                    ws[f"{dep_stop_col}{row}"].color = f"{no_color}"
                #always save
                wb.save(matrix_file)
                #updateUi for further use
                log_time = datetime.strftime(datetime.now(),"%H:%M %d-%m-%y")
                message = f"{option} line <font color={info_color}><b>{line}</b></font> updated at <font color={info_color}><b>{log_time}</font></b>"
                logger.info(f"{option} line {line} successfully updated")
                self.updateLabel.setText(message)
                self.updateUi()
            except Exception as exc:
                message = f"<font color={alert_color}><b>Something went wrong. Check the log file</b></font>"
                logger.error(f"Error occured during matrix file update: {exc}", exc_info = True)
                self.updateLabel.setText(message)
                self.updateUi()
            return
        elif option == "Recovered":
            try:
                line, dep_count, dep_start, dep_stop, rec_count, rec_start, rec_stop = \
                get_line_matrix_stats(get_rcv_line_df(self.recovered_df,int(line_nb), short = False)) 
                row = int(rlines_col.index(float(line_nb)))+1

                ws[f"{rec_count_col}{row}"].value = rec_count
                if  ws[f"{rec_count_col}{row}"].value == ws[f"{dep_count_col}{row}"].value: #check against the number of deployed nodes
                    ws[f"{rec_count_col}{row}"].color = f"{accept_color}"
                else:
                    ws[f"{rec_count_col}{row}"].color = f"{warning_color}"

                if  ws[f"{rec_start_col}{row}"].value != datetime.strptime(rec_start, "%Y-%m-%d %H:%M:%S"):
                    ws[f"{rec_start_col}{row}"].value = rec_start
                    ws[f"{rec_start_col}{row}"].color = f"{warning_color}"
                else:
                    ws[f"{rec_start_col}{row}"].color = f"{no_color}"
                
                if  ws[f"{rec_stop_col}{row}"].value != datetime.strptime(rec_stop, "%Y-%m-%d %H:%M:%S"):
                    ws[f"{rec_stop_col}{row}"].value = rec_stop
                    ws[f"{rec_stop_col}{row}"].color = f"{warning_color}"
                else:
                    ws[f"{rec_stop_col}{row}"].color = f"{no_color}"
                # save file as usual
                wb.save(matrix_file)
                # update GUI for further usage
                log_time = datetime.strftime(datetime.now(),"%H:%M %d-%m-%y")
                message = f"{option} line <font color={info_color}><b>{line}</b></font> updated at <font color={info_color}><b>{log_time}</font></b>"
                logger.info(f"{option} line {line} successfully updated")
                self.updateLabel.setText(message)
                self.updateUi()
            except Exception as exc:
                message = f"<font color={alert_color}><b>Something went wrong. Check the log file</b></font>"
                logger.error(f"Error occured during matrix file update: {exc}", exc_info = True)
                self.updateLabel.setText(message)
                self.updateUi()
            return

        elif option == "Source":
            #source line attributes
            try:
                line, src_count, src_start, src_stop = \
                get_matrix_stats(get_line_df(self.source_df,int(line_nb)))
                row = int(srclines_col.index(float(line_nb)))+1
                
                ws[f"{src_count_col}{row}"].value = src_count
                if  (ws[f"{src_count_col}{row}"].value/ws[f"{preplot_src_count_col}{row}"].value)*100 < 97.0:
                    ws[f"{src_count_col}{row}"].color = f"{warning_color}"
                else:
                    ws[f"{src_count_col}{row}"].color = f"{accept_color}"

                if  ws[f"{src_start_col}{row}"].value != datetime.strptime(src_start,"%Y-%m-%d %H:%M:%S"):
                    ws[f"{src_start_col}{row}"].value = src_start
                    ws[f"{src_start_col}{row}"].color =f"{warning_color}"
                else:
                    ws[f"{src_start_col}{row}"].color =f"{no_color}"

                if  ws[f"{src_stop_col}{row}"].value != datetime.strptime(src_stop, "%Y-%m-%d %H:%M:%S"):
                    ws[f"{src_stop_col}{row}"].value = src_stop
                    ws[f"{src_stop_col}{row}"].color = f"{warning_color}"
                else:
                    ws[f"{src_stop_col}{row}"].color = f"{no_color}"
                # save file
                wb.save(matrix_file)
                # update UI
                log_time = datetime.strftime(datetime.now(),"%H:%M %d-%m-%y")
                message = f"{option} line <font color={info_color}><b>{line}</b></font> updated at <font color={info_color}><b>{log_time}</font></b>"
                logger.info(f"{option} line {line} successfully updated")
                self.updateLabel.setText(message)
                self.updateUi()
            except Exception as exc:
                message = f"<font color={alert_color}><b>Something went wrong. Check the log file</b></font>"
                logger.error(f"Error occured during matrix file update: {exc}", exc_info = True)
                self.updateLabel.setText(message)
            return

    def updateUi(self):
        '''update to initial app state'''
        self.updateButton.setEnabled(False)
        self.lineInput.setFocus()
        self.lineInput.selectAll()
        return

    def updateData(self):
        '''update data '''
        src_updated = False
        #avoid updating just for nothing if there is no change in file
        #still won't work as we are rebuilding all files
        temp_stats = get_file_stats(sps_s_clean)
        if self.source_stats and temp_stats:
            if  temp_stats != self.source_stats: 
                self.source_df = source_sps_to_df(sps_s_clean)
                self.source_stats = get_file_stats(sps_s_clean)
                src_updated = True
            else:
                pass
        else:
            pass

        self.deployed_df = deployed_sps_to_df(sps_r_deploy)
        self.recovered_df = recovered_sps_to_df(sps_r_recover)

        log_time = datetime.strftime(datetime.now(),"%H:%M %d-%m-%y")
        message = f"Input data updated at <font color={info_color}><b>{log_time}</font></b>"
        logger.info(f"Input data successfully updated")
        self.updateLabel.setText(message)
        self.updateButton.setEnabled(False)
        self.lineInput.setFocus()
        return

if __name__ == "__main__":
    #create the directory for logs
    if not os.path.exists('matrix_helper_logs'):
        os.mkdir('matrix_helper_logs')

    #set up the logger 
    logger = logging.getLogger("matrix helper")
    logger.setLevel(logging.DEBUG)
    
    handler = RotatingFileHandler(paths.log_file(), maxBytes=100240, backupCount=10)
    handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s', datefmt = '%d-%b-%Y %H:%M:%S')
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    #set paths to production files, if some file doesn't exist or empty it is handled later 
    sps_s_clean = paths.sps_s_clean()
    sps_r_deploy = paths.spr_deployed()
    sps_r_recover = paths.spr_recovered()

    #try to open the matrix file
    #application started with an error message 
    #in case file is unavailable
    matrix_file = paths.matrix_file()
    if not os.path.exists(matrix_file):
        startUp_message = f"<font_color={alert_color}><b>Create matrix file first</b></font>"
        logger.error("Application started without available matrix file")
    else:
    #open worksheet and save it
    #as xlwings tends to crash on application starts, use try-except block
    #in case of crash on start up 
    #notification will appear on the start up 
        try:
            MAX_SOURCE = 300
            wb = xw.Book(matrix_file)
            wb.save(matrix_file)
            ws = xw.sheets["RL-SL"]
            
            rlines_col = ws.range(f"{rcv_lines_col}:{rcv_lines_col}")[:MAX_SOURCE].value
            srclines_col = ws.range(f"{src_lines_col}:{src_lines_col}")[:MAX_SOURCE].value
            
            log_time = datetime.strftime(datetime.now(),"%H:%M %d-%m-%y")
            startUp_message = f"Suceessfully started app at <font color={info_color}><b>{log_time}</font></b>"
            logger.info("Application successfully started")
        except:
            startUp_message = f"<font color={alert_color}><b>Error during app start. Check log file.</font></b>"
            logger.error("Application didn't start correctly:", exc_info=True)

    app = QApplication(sys.argv)
    dialog = DumbDialog()
    dialog.show()
    app.exec_()
