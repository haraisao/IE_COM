# -*- coding: utf-8 -*-
import sys
import types
import win32com.client
from time import sleep


def getFileName(pth):
  pths = pth.split('/')
  return pths[len(pths) -1]

def findString(info, key) :
  if type(info) == types.StringType or type(info) == types.UnicodeType :
    if info.find(key) >= 0:
      return True
  return False

def isWebUrl(info) :
  if type(info) == types.StringType or type(info) == types.UnicodeType :
    if info.find('http') >=0 : return True
  return False

#
#  COM Wrapper for Excel
#
class excel_com:
  def __init__(self, new_win=True):
    if new_win :
      self.excel = win32com.client.Dispatch("Excel.Application")
      self.excel.Visible=1
    self.shell = None
    self.wshell = None

  def setShell(self) :
    if not self.shell :
      self.shell = win32com.client.Dispatch("Shell.Application")
    if not self.wshell :
      self.wshell = win32com.client.Dispatch("WScript.Shell")




#---- sample navigation
def main():
  exl=excel_com()
  exl.Visible=1



if __name__ == "__main__":
  main()

