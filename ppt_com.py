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
    self.window = None

  def setShell(self) :
    if not self.shell :
      self.shell = win32com.client.Dispatch("Shell.Application")
    if not self.wshell :
      self.wshell = win32com.client.Dispatch("WScript.Shell")

  def open(self, fname=""):
    try:
      if fname :
        self.presentation = self.app.Presentations.Open(fname)
      else:
        self.presentation = self.app.Presentations.Add()
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
      self.slide = self.presentation.Slides.Add(pos, cat)
    except:
      print "Error"

  def addSlide(self, cat):
    n = self.presentation.Slides.Count + 1
    self.newSlide(n, cat)

  def runSlideShow(self):
    try:
      self.window = self.presentation.SlideShowSettings.Run()
    except:
      print "Error"

  def next(self):
    try:
      self.window.View.Next()
    except:
      print "Error"

  def prev(self):
    try:
      self.window.View.Previous()
    except:
      print "Error"

  def end(self):
    try:
      self.window.View.Exit()
      self.window=None
    except:
      print "Error"

  def goto(self, n):
    try:
      self.window.View.GotoSlideEnd(n)
    except:
      print "Error"


#---- sample navigation
def main():
  exl=ppt_com()
  exl.Visible=1

if __name__ == "__main__":
  main()

