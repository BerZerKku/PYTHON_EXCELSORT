# -*- coding: cp1251 -*-
import sys
from PyQt4 import QtCore, QtGui, uic
import icons
import excelSort
 
class MainWindow(QtGui.QDialog):
   def __init__(self):
      QtGui.QDialog.__init__(self)
      uic.loadUi('excelSortGUI.ui', self)

      self.pbSave.setDisabled(True)
      self.leSave.setReadOnly(True)
      self.leOpen.setReadOnly(True)
##      self.leSave.setDisabled(True)
##      self.leOpen.setDisabled(True)
      
      self.pbSave.clicked.connect(self.fileSave)
      self.pbOpen.clicked.connect(self.fileOpen)
      self.sort = excelSort.excelSort()
      
      
   def fileSaveAs(self, checked=False):
      ''' (MainWindow, bool) -> None
     
         Сохранение файла с помощью диалога Save.
      '''
      name = unicode(QtGui.QFileDialog.getSaveFileName(self, u"Сохранить",
               filter=u"XLS files (*.xls)"))
      # если имя было выбрано, передход к функции сохранения
      if name:
         self.fileSave(name=name)

   def fileOpen(self, checked=False):
      ''' (MainWindow, bool) ->
            Загрузка файла.      
      '''
      name = QtGui.QFileDialog.getOpenFileName(self, u"Открыть",
               filter=u"XLS files (*.xls *.xlsx)")
      
      if name:
         ext = name.split('.')[-1]
         saveName = name[0:-(len(ext)+1)] + u'_sort.xls'
         self.leOpen.setText(name)
         self.leSave.setText(saveName)
         self.pbSave.setEnabled(True)
         self.sort.fileOpen(name)
         
      

   def fileSave(self, checked=False):
      name = self.leSave.text()
      if name:
         self.sort.sortOne()
         self.sort.fileSave(name)
   
app = QtGui.QApplication(sys.argv)
w = MainWindow()
w.show()
sys.exit(app.exec_())
