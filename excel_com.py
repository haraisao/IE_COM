# -*- coding: utf-8 -*-
import sys
import types
import win32com.client
from time import sleep

#
#  COM Wrapper for Excel
#
class excel_com:
  def __init__(self, new_win=True):
    if new_win :
      self.app = win32com.client.Dispatch("Excel.Application")
      self.app.Visible=1
    self.shell = None
    self.wshell = None
    self.workbook = None
    self.sheet = None

  def setShell(self) :
    if not self.shell :
      self.shell = win32com.client.Dispatch("Shell.Application")
    if not self.wshell :
      self.wshell = win32com.client.Dispatch("WScript.Shell")

  def open(self, fname):
    try:
      self.workbook = self.app.Workbooks.Open(fname)
      self.sheet = self.workbook.Worksheets(1).Activate()
    except:
      print "Error in open %s" % fname

  def save(self, fname=""):
    try:
      if fname:
        self.workbook.SaveAs(fname)
      else:
        self.workbook.Save()
    except:
        print "Error in open %s" % fname

  def newWorkbook(self):
    try:
      self.workbook = self.app.Workbooks.Add()
      self.sheet = self.workbook.Worksheets(1).Activate()
    except:
      print "Error in newWorkbook"

  def newSheet(self):
    try:
      self.sheet = self.workbook.Add()
    except:
      print "Error in newSheet"

  def getSheetNames(self):
    res = []
    try:
      for x in range(self.workbook.Worksheets.Count):
        res.append(self.workbook.Worksheets(x+1).Name)
    except:
      print "Error in getSheetNames"
    return  res

  def selectSheet(self, name):
    try:
      self.sheet = self.workbook.Worksheets(name).Activate()
    except:
      print "Error in selectSheet"

  def deletetSheet(self, name):
    try:
      if self.sheet == self.workbook.Worksheets(name):
        self.sheet == None
      self.workbook.Worksheets(name).Delete()
      if not self.sheet:
        self.sheet = self.workbook.Worksheets(name).Activate()
    except:
      print "Error in deleteSheet"


#---- sample navigation
def main():
  exl=excel_com()
  exl.Visible=1



if __name__ == "__main__":
  main()

