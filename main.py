# -*- coding: utf-8 -*-
# Form implementation generated from reading ui file 'SVG To PDF Converter Beta.ui'
#
# Created by: Vo Hao Hao
#
# WARNING! All changes made in this file will be lost!

import os
import subprocess
import time
import traceback

import natsort as natsort
from PyPDF2 import PdfFileMerger, PdfFileReader
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from tkinter import *
from tkinter import filedialog
from PyQt5.QtWidgets import QAbstractItemView

dataz = [" ", " ", " ", " ", " ", " ", " "]
header = ["Files location"]
state = 0

class WorkerSignals(QObject):
    '''
    Defines the signals available from a running worker thread.

    Supported signals are:

    finished
        No data

    error
        `tuple` (exctype, value, traceback.format_exc() )

    result
        `object` data returned from processing, anything

    progress
        `int` indicating % progress

    '''
    finished = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(object)
    progress = pyqtSignal(int)


class Worker(QRunnable):
    '''
    Worker thread

    Inherits from QRunnable to handler worker thread setup, signals and wrap-up.

    :param callback: The function callback to run on this worker thread. Supplied args and
                     kwargs will be passed through to the runner.
    :type callback: function
    :param args: Arguments to pass to the callback function
    :param kwargs: Keywords to pass to the callback function

    '''

    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()

        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

        # Add the callback to our kwargs
        self.kwargs['progress_callback'] = self.signals.progress #check progress_callbank

    @pyqtSlot()
    def run(self):
        '''
        Initialise the runner function with passed args, kwargs.
        '''

        # Retrieve args/kwargs here; and fire processing using them
        try:
            result = self.fn(*self.args, **self.kwargs)
        except:
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            self.signals.result.emit(result)  # Return the result of the processing
        finally:
            self.signals.finished.emit()  # Done

class TableModel(QtCore.QAbstractTableModel):
    '''
    TableModel

    Table with custom header model & count.
    Check headerdata to find need information :))

    '''

    def __init__(self, data, header):
        super(TableModel, self).__init__()
        self._data = data
        self.header = header

    def data(self, index, role):
        if not index.isValid():
            return None
        if (index.column() == 0):
            value = self._data[index.row()][index.column()]
        if role == QtCore.Qt.EditRole:
            return value
        elif role == QtCore.Qt.DisplayRole:
            return value

    def rowCount(self, index):
        # The length of the outer list.
        return len(self._data)

    def columnCount(self, index):
        # The following takes the first sub-list, and returns
        # the length (only works if all rows are an equal length)
        return len(self._data[0])

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.header[col]
        if orientation == Qt.Vertical and role == Qt.DisplayRole:
            return col + 1

class Ui_Dialog(object):
    '''
    Ui_Dialog

    Main UI Dialog. Connection with button

    '''

    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(597, 410)
        self.horizontalLayoutWidget = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(10, 10, 361, 51))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.pushButton_2 = QtWidgets.QPushButton(self.horizontalLayoutWidget)
        self.pushButton_2.setObjectName("pushButton_2")
        self.horizontalLayout.addWidget(self.pushButton_2)
        self.pushButton_2.clicked.connect(self.selectFiles)
        self.pushButton_3 = QtWidgets.QPushButton(self.horizontalLayoutWidget)
        self.pushButton_3.setObjectName("pushButton_3")
        self.horizontalLayout.addWidget(self.pushButton_3)
        self.pushButton_3.clicked.connect(self.selectFolder)
        self.pushButton_4 = QtWidgets.QPushButton(self.horizontalLayoutWidget)
        self.pushButton_4.setObjectName("pushButton_4")
        self.horizontalLayout.addWidget(self.pushButton_4)
        self.pushButton_4.clicked.connect(self.delete)
        self.pushButton_5 = QtWidgets.QPushButton(self.horizontalLayoutWidget)
        self.pushButton_5.setObjectName("pushButton_5")
        self.horizontalLayout.addWidget(self.pushButton_5)
        self.pushButton_5.clicked.connect(self.delete_all)
        self.verticalLayoutWidget = QtWidgets.QWidget(Dialog)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(380, 10, 201, 51))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(self.verticalLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.label_2 = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.verticalLayout.addWidget(self.label_2)
        self.tableView = QtWidgets.QTableView(Dialog)
        self.tableView.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tableView.horizontalHeader().setHighlightSections(False)
        self.tableView.horizontalHeader().setDefaultSectionSize(571)
        self.tableView.setGeometry(QtCore.QRect(10, 71, 571, 235))
        self.tableView.setObjectName("tableView")
        self.progressBar = QtWidgets.QProgressBar(Dialog)
        self.progressBar.setGeometry(QtCore.QRect(320, 340, 261, 21))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(10, 330, 81, 41))
        self.label_3.setObjectName("label_3")
        self.verticalLayoutWidget_2 = QtWidgets.QWidget(Dialog)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(99, 330, 151, 42))
        self.verticalLayoutWidget_2.setObjectName("verticalLayoutWidget_2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_2)
        self.verticalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.radioButton = QtWidgets.QRadioButton(self.verticalLayoutWidget_2)
        self.radioButton.setObjectName("radioButton")
        self.verticalLayout_2.addWidget(self.radioButton)
        self.radioButton.setChecked(True)
        self.radioButton.toggled.connect(lambda: self.btnstate(self.radioButton))
        self.radioButton_2 = QtWidgets.QRadioButton(self.verticalLayoutWidget_2)
        self.radioButton_2.setObjectName("radioButton_2")
        self.verticalLayout_2.addWidget(self.radioButton_2)
        self.radioButton_2.toggled.connect(lambda: self.btnstate(self.radioButton_2))
        self.lineEdit = QtWidgets.QLineEdit(Dialog)
        self.lineEdit.setGeometry(QtCore.QRect(122, 381, 201, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton = QtWidgets.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(330, 380, 84, 23))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.selectSave)
        self.label_4 = QtWidgets.QLabel(Dialog)
        self.label_4.setGeometry(QtCore.QRect(260, 330, 61, 41))
        self.label_4.setObjectName("label_4")
        self.pushButton_6 = QtWidgets.QPushButton(Dialog)
        self.pushButton_6.setGeometry(QtCore.QRect(444, 372, 131, 31))
        self.pushButton_6.setObjectName("pushButton_6")
        self.pushButton_6.clicked.connect(self.convertAll)
        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)
        self.tableView.keyPressEvent = self.keyPressEvent
        self.threadpool = QThreadPool()

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowIcon(QIcon('lib/icon.png'))
        Dialog.setWindowTitle(_translate("Dialog", "SVG To PDF Converter Beta"))
        self.pushButton_2.setText(_translate("Dialog", "Add Files"))
        self.pushButton_3.setText(_translate("Dialog", "Add Folder"))
        self.pushButton_4.setText(_translate("Dialog", "Remove Selected"))
        self.pushButton_5.setText(_translate("Dialog", "Remove All"))
        self.label.setText(_translate("Dialog", "SVG To PDF Converter Beta"))
        self.label_2.setText(_translate("Dialog", "version: 1.12 | <a href='http://elearning-sos.com/'>www.elearning-sos.com</a>"))
        self.label_3.setText(_translate("Dialog", "Saving options:"))
        self.radioButton.setText(_translate("Dialog", "Same as source folder"))
        self.radioButton_2.setText(_translate("Dialog", "Following folder"))
        self.pushButton.setText(_translate("Dialog", "Change ..."))
        self.label_4.setText(_translate("Dialog", "Progress:"))
        self.pushButton_6.setText(_translate("Dialog", "Convert All"))

    '''
    Delete 

    Select row and delete

    '''

    def delete(self):
        if self.tableView.selectionModel().hasSelection():
            indexes = [QPersistentModelIndex(index) for index in self.tableView.selectionModel().selectedRows()]
            global dataz
            for index in indexes:
                dataz[index.row()] = " ";
            dataz = ui.updateRows(dataz) #custom ways to :v delete
            model = TableModel(dataz, header)
            ui.tableView.setModel(model)
        else:
            print('No row selected')

    '''
    Delete all

    Reset the table

    '''

    def delete_all(self):
        global dataz
        dataz = [" ", " ", " ", " ", " ", " ", " "]
        dataz = ui.updateRows(dataz)
        model = TableModel(dataz, header)
        ui.tableView.setModel(model)

    '''
    KeyPressEvent

    Check delete and backspace key to delete seleted row

    '''

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Backspace, Qt.Key_Delete):
            self.delete()

    '''
    UpdateRows

    Custom progress updating table

    '''

    def updateRows(self, data):
        space_array = []
        o_rows = data
        i = 0
        for row in data:
            if(row ==  " "):
                space_array.append(i)
            i += 1
        space_array.reverse()
        for pos in space_array:
            del o_rows[pos]
        while len(o_rows) < 7:
            o_rows.append(" ")
        return  o_rows

    '''
    SelectFiles

    Select multifiles with extension and update table

    '''

    def selectFiles(self):
        global dataz
        global header
        root.filename = filedialog.askopenfilenames(initialdir="/", title="Select files",
                                                    filetypes=(("SVG Files", ".svg .svgz"), ("HTML Files", ".html .xhtml"), ("All Files", "*.*")))
        for item in (list(map(str, root.filename))):  # comma, or other
           if [item] not in dataz:
            dataz.append([item])

        dataz = ui.updateRows(dataz)
        model = TableModel(dataz, header)
        ui.tableView.setModel(model)

    '''
    SelectFolder

    Select all files in folders

    '''

    def selectFolder(self):
        global dataz
        global header
        root.directory = filedialog.askdirectory()
        if root.directory != "":
            for file in os.listdir(root.directory):
                if (file.endswith(".svg") or file.endswith(".svgz")) and [root.directory + "/" + file] not in dataz:
                    dataz.append([root.directory + "/" + file])
            dataz = ui.updateRows(dataz)
            dataz = natsort.natsorted(dataz) #fixed folder correct order
            model = TableModel(dataz, header)
            ui.tableView.setModel(model)

    '''
    SelectSave

    Select save location

    '''

    def selectSave(self):
        root.directory = filedialog.askdirectory()
        self.lineEdit.setText(root.directory)
        self.radioButton_2.setChecked(True)

    '''
    CovertAll

    Covert buttons

    '''

    def convertAll(self):
        # Pass the function to execute
        worker = Worker(self.convertProgress)  # Any other args, kwargs are passed to the run function
        worker.signals.result.connect(self.print_output)
        worker.signals.finished.connect(self.thread_complete)
        worker.signals.progress.connect(self.progress_fn)

        # Execute
        self.threadpool.start(worker)

    '''
    CutAll

    Cut border pdf files

    '''

    def cutAll(self):
        # Pass the function to execute
        worker = Worker(self.cutProgress)  # Any other args, kwargs are passed to the run function
        worker.signals.result.connect(self.cut_output)
        worker.signals.finished.connect(self.cut_complete)
        worker.signals.progress.connect(self.cut_progress_fn)
        # Execute
        self.threadpool.start(worker)

    '''
    convertProgress

    Covert with chrome

    '''

    def convertProgress(self, progress_callback):
        n = 0
        i = 0
        for data in dataz:
            if(data[0] != " "):
                process = subprocess.Popen('chrome_ex\App\Chrome-bin\chrome.exe --headless --disable-gpu --print-to-pdf="' + os.path.dirname(os.path.abspath(__file__)) + '/temp/' + os.path.splitext(os.path.basename(data[0]))[0] + '.pdf" "' + data[0] + '"', shell=True)
                n += 1
                i += 1
                progress_callback.emit(n)
                if (i >= self.threadpool.maxThreadCount()/2):
                    process.wait()
                    i = 0
        process.wait()
        return "Converted"

    '''
    cutProgress

    Cut progress

    '''

    def cutProgress(self, progress_callback):
        global state
        n = 0
        mergedObject = PdfFileMerger()
        for data in dataz:
            if data[0] != " ":
                mergedObject.append(PdfFileReader(os.path.dirname(os.path.abspath(__file__)) + '/temp/' + os.path.splitext(os.path.basename(data[0]))[0] + '.pdf', 'rb'))
                n += 1
                progress_callback.emit(n)
        name = "svg_convet"
        mergedObject.write(os.path.dirname(os.path.abspath(__file__))  + '/temp/' + name +".pdf")
        mother_folder = ""
        if state == 0:
            mother_folder = os.path.dirname(os.path.realpath(dataz[0][0]))
        if state == 1:
            mother_folder = self.lineEdit.text()
        process = subprocess.Popen('lib\crop.exe -u -s -ap4 0 25 0 25 "' + os.path.dirname(
                os.path.abspath(__file__)) + '/temp/' + name + '.pdf" -o "' + mother_folder + "/" + name + '.pdf"',
                                       shell=True)
        process.wait()
        os.remove(os.path.dirname(os.path.abspath(__file__)) + '/temp/' + name + '.pdf')
        time.sleep(1)
        for data in dataz:
            if data[0] != " ":
                os.remove(os.path.dirname(os.path.abspath(__file__)) + '/temp/' + os.path.splitext(os.path.basename(data[0]))[0] + '.pdf')
        time.sleep(1)
        return "Converted"


    def print_output(self, s):
        print(s)

    def thread_complete(self):
        print("THREAD COMPLETE!")
        self.cutAll() #Run cut progress

    def progress_fn(self, n):
        self.progressBar.setProperty("value", n*50/(len(dataz) - dataz.count(" ")))

    def cut_output(self, s):
       print(s)

    def cut_complete(self):
        name = "svg_convet"
        print("THREAD COMPLETE!")
        self.progressBar.setProperty("value", 100)
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText('Your file name is \n' + name + ".pdf")
        msg.setWindowTitle("Done")
        msg.setWindowIcon((QIcon('lib/icon.png')))
        msg.exec_()

    def cut_progress_fn(self, n):
        self.progressBar.setProperty("value",50 + n*49/(len(dataz) - dataz.count(" ")) )

    def btnstate(self, b):
        global state
        if b.text() == "Same as source folder" and b.isChecked():
            state = 0
        if b.text() == "Following folder" and b.isChecked():
            state = 1


    '''
    Main 

    '''
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(MainWindow)
    root = Tk()
    root.withdraw()
    dataz = ui.updateRows(dataz)
    model = TableModel(dataz, header)
    ui.tableView.setModel(model)
    MainWindow.show()
    sys.exit(app.exec_())
