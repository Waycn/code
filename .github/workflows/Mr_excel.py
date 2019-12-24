# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""

from PyQt4.QtGui import QMainWindow
from PyQt4.QtCore import pyqtSignature
from Ui_Mr_Excel import Ui_MainWindow

from PyQt4 import QtCore, QtGui
from win32com import client as wc
from csv import excel


class MainWindow(QMainWindow, Ui_MainWindow):
    """
    Class documentation goes here.
    """
    def __init__(self, parent = None):
        """3
        
        Constructor
        """
        QMainWindow.__init__(self, parent)
        self.setupUi(self)
  
   #选择excel按钮
    @pyqtSignature("")
    def on_pushButton_clicked(self):
        """
        Slot documentation goes here.
        """
        # TODO: not implemented yet
        my_file_path=QtGui.QFileDialog.getOpenFileName(self, "Select File", '../')#选择excel文件，返回路径
        if unicode(my_file_path).endswith('.lxsx'):
            excel=wc.Dispatch('Excel.Application')
            excel.Visible=0  #设置文件为不可见
            my_excel=excel.Workbooks.Open(my_file_path)
            
           



if __name__ == "__main__":
    import sys
    app = QtGui.QApplication(sys.argv)
    ui = MainWindow()
    ui.show()
    sys.exit(app.exec_())
