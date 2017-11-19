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
#  COM Wrapper for PowerPoint
#
class ppt_com:
  def __init__(self, new_win=True):
    if new_win :
      self.app = win32com.client.Dispatch("PowerPoint.Application")
      self.app.Visible=1
    self.shell = None
    self.presentation = None
    self.slide = None
    self.view = None

  def setShell(self) :
    if not self.shell :
      self.shell = win32com.client.Dispatch("Shell.Application")
    if not self.wshell :
      self.wshell = win32com.client.Dispatch("WScript.Shell")

  def open(self, fname):
    try:
      self.presentation = self.app.Presentations.Open(fname)
    except:
      print "Error in open %s" % fname

  def save(self, fname=""):
    try:
      if fname:
        self.presentation.SaveAs(fname)
      else:
        self.presentation.Save()
    except:
        print "Error in open %s" % fname

  def close(self):
    try:
      self.presentation.Close()
    except:
      print "Error in open %s" % fname

  def newSlide(self, pos, cat):
    try:
      self.slide = self.presentation.Add()
    except:
      print "Error"

  def runSlideShow(self):
    try:
      self.view = self.presentation.SlideShowSettings.Run()
    except:
      print "Error"

  def next(self):
    try:
      self.view.View.Next()
    except:
      print "Error"

  def prev(self):
    try:
      self.view.View.Previous()
    except:
      print "Error"

  def end(self):
    try:
      self.view.View.Exit()
      self.view=None
    except:
      print "Error"

  def goto(self, n):
    try:
      self.view.View.GotoSlideEnd(n)
    except:
      print "Error"

#---- sample navigation
def main():
  exl=excel_com()
  exl.Visible=1



if __name__ == "__main__":
  main()

