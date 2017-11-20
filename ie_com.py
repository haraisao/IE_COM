# -*- coding: utf-8 -*-
import sys
import time
import types
import win32com.client
from time import sleep


#-----------------
home = "http://www.yahoo.co.jp"
#-----------------

def getFileName(pth):
  pths = pth.split('/')
  return pths[len(pths) -1]

def getAnchorInfo(ele):
  if not ele.firstChild or ele.firstChild.nodeName == '#text':
    return ele.innerHTML
  child = ele.firstChild
  if child.tagName == 'IMG':
    if child.title : return "IMG:"+child.title
    if child.alt : return "IMG:"+child.alt
    return "IMG:"+getFileName(child.src)

def findString(info, key) :
  if type(info) == types.StringType or type(info) == types.UnicodeType :
    if info.find(key) >= 0:
      return True
  return False

def isWebUrl(info) :
  if type(info) == types.StringType or type(info) == types.UnicodeType :
    if info.find('http') >=0 : return True
    if not info : return True
  return False

####

def ActivateWin(wshell,name):
  if wshell.AppActivate(name) :
    time.sleep(1)
    print "---"
    wshell.AppActivate('TEST')


#################################
#
#  COM Wrapper for IE
#
class ie_com:

  def __init__(self, new_win=False):
    self.ie =None
    self.shell = None
    self.wshell = None
    if new_win : self.newWindow()

  def newWindow(self, visible=1) :
    self.ie = win32com.client.Dispatch("InternetExplorer.Application")
    self.ie.Visible=visible

  def setShell(self) :
    if not self.shell :
      self.shell = win32com.client.Dispatch("Shell.Application")
    if not self.wshell :
      self.wshell = win32com.client.Dispatch("WScript.Shell")

  def getNumOfWindows(self) :
    self.setShell()
    return self.shell.Windows().Count

  def getWebWindow(self):
    count = self.getNumOfWindows()
    res=[]
    for i in range(count) :
      if self.shell.Windows().Item(i) :
        loc = self.shell.Windows().Item(i).LocationName
        if isWebUrl(self.shell.Windows().Item(i).LocationURL) :
          res.append(i)
    return res    

  def setNthWindow(self, n):
    wins = self.getWebWindow()
    try:
      self.setTargetWindow(wins[n])
    except:
      print "Error in setNthWindow"
    return

  def setFirstWindow(self):
    wins = self.setNthWindow(0)
    return

  def setLastWindow(self):
    wins = self.setNthWindow(-1)
    return

  def getWindowList(self) :
    count = self.getNumOfWindows()
    res = {}
    for i in range(count) :
      if self.shell.Windows().Item(i) :
        loc = self.shell.Windows().Item(i).LocationName
        if isWebUrl(self.shell.Windows().Item(i).LocationURL) :
          res[i] = "[Web] %s" % (loc)
        else:
          res[i] = "[File] %s" % (loc)
    return res

  def listWindow(self) :
    wins = self.getWindowList()
    print "==== List of IE Window ===="
    for k,info in wins :
      print "%d: %s" % (k, info)
    print "==========================="
    return 

  def getWindowsList(self) :
    count = self.getNumOfWindows()
    res = {}
    for i in range(count) :
      if self.shell.Windows().Item(i) :
        loc = self.shell.Windows().Item(i).LocationName
        if isWebUrl(self.shell.Windows().Item(i).LocationURL) :
          res[i]= "[Web] %s" % (loc)
        else:
          res[i]= "[File] %s" % (loc)
    return res

  def setTargetWindow(self, idx) :
    if type(idx) == types.IntType :
      count = self.getNumOfWindows()
      if idx < count and idx >= 0:
        self.ie = self.shell.Windows().Item(idx)
        print "set Window to '%s'" % (self.ie.LocationName)
        return True
      else:
        print "Invalid Window index, (0 < idx < %d )" % (count)
    else:
      print "Invalid Argument: the arg1 should be IntType."
    return False

  def closeWindow(self, idx=-1) :
    count = self.getNumOfWindows()
    if idx < 0:
      self.ie.Quit()
      self.ie=None
    elif idx < count :
      self.ie = self.shell.Windows().Item(idx)
      self.ie.Quit()
    else:
      print "Invalid Window index, (0 < idx < %d )" % (count)

  def navigate(self, url):
    self.ie.Navigate(url)
    while self.ie.Busy : sleep(1)

  def getDocument(self):
    return self.ie.Document.Body

  def getHTML(self):
    return self.getDocument().innerHTML

  def getElementsByTagName(self, tag):
    return self.getDocument().getElementsByTagName(tag)

  def getElementByName(self, tag, name):
    itms = self.getElementsByTagName(tag)
    for i in range(itms.length):
      if itms[i].name == name:
        return itms[i]
    return None

  def getElementByValue(self, tag, val):
    itms = self.getElementsByTagName(tag)
    for i in range(itms.length):
      if itms[i].value == val:
        return itms[i]
    return None

  def getAnchorList(self, key=None):
    itms = self.getElementsByTagName('a')
    res = {}
    for i in range(itms.length):
      info = getAnchorInfo(itms[i])
      if key :
        if findString(info, key) :
          res[i] = info
      else:
        res[i]= info
    return res

  def listAnchor(self, key=None, start=0, count=20):
    itms = self.getElementsByTagName('a')
    countIdx = 0
    print "==== List of anchors (%d/%d)====" % (start, itms.length)
    for i in range(itms.length):
      if i >= start:
        if countIdx < count :
          info = getAnchorInfo(itms[i])
          if key :
            if findString(info, key) :
              print "[ %d: %s ]" % (i, info)
              countIdx += 1
          else:
            print "%d: %s" % (i, info)
            countIdx += 1
    print "========================="
    return 

  def getAnchorByIndex(self, n):
    itms = self.getElementsByTagName('a')
    if n>=0 and n < itms.length:
      return itms[n]
    else:
      print "Invaid Index: 0 =< %d < %d" % (n. items.length)
      return None

  def clickAnchorByIndex(self, n):
    itm = self.getAnchorByIndex(n)
    if itm : itm.click()

  def moveAnchorSiteByIndex(self, n):
    itm = self.getAnchorByIndex(n)
    if itm : self.navigate(itm.href)

  def getAnchorByValue(self, val, flag='all'):
    itms = self.getElementsByTagName('a')
    for i in range(itms.length):
      if flag == 'all' :
        if itms[i].innerHTML == val:
          return itms[i]
      else:
        if  findString(itms[i].innerHTML, val) :
          return itms[i]

    print "No such an anchor: %s" % (val)
    return None

  def clickAnchor(self, val, flag='all'):
    itm = self.getAnchorByValue(val, flag)
    if itm : itm.click()

  def moveAnchorSite(self, val, flag='all'):
    itm = self.getAnchorByValue(val, flag)
    if itm : self.navigate(itm.href)

  def click(self, x, flag='all'):
    if type(x) == types.IntType :
      self.clickAnchorByIndex(x)
    elif type(x) == types.StringType or type(x) == types.UnicodeType :
      self.clickAnchor(x, flag)

  def move(self, x, flag='all'):
    if type(x) == types.IntType :
      self.moveAnchorSiteByIndex(x)
    elif type(x) == types.StringType or type(x) == types.UnicodeType :
      self.moveAnchorSite(x, flag)

  def listInputs(self, val=None):
    itms = self.getElementsByTagName('input')
    res = "==== List of anchors ====\n"
    for i in range(itms.length):
      if itms[i].type != 'hidden' :
        if val :
          info = itms[i].value
          name = itms[i].name
          if findString(info, val) or findString(name, val) :
            res += "%d: [ type = %s, name = %s, value = %s ]\n" % (i, itms[i].type, itms[i].name, itms[i].value)
        else:
          res += "%d: type = %s, name = %s, value = %s\n" % (i, itms[i].type, itms[i].name, itms[i].value)
    res += "=========================\n"

    return res

  def listButtons(self, val=None):
    itms = self.getElementsByTagName('input')
    res = "==== List of anchors ====\n"
    for i in range(itms.length):
      if itms[i].type == 'button' or itms[i].type == 'submit' :
        if val :
          info = itms[i].value
          name = itms[i].name
          if findString(info, val) or findString(name, val) :
            res += "%d: [ type = %s, name = %s, value = %s ]\n" % (i, itms[i].type, itms[i].name, itms[i].value)
        else:
          res += "%d: type = %s, name = %s, value = %s\n" % (i, itms[i].type, itms[i].name, itms[i].value)
    res += "=========================\n"

    return res

  def listTextInputs(self, val=None):
    itms = self.getElementsByTagName('input')
    res = "==== List of anchors ====\n"
    for i in range(itms.length):
      if itms[i].type == 'text' or itms[i].type == '' :
        if val :
          info = itms[i].value
          name = itms[i].name
          if findString(info, val) or findString(name, val) :
            res += "%d: [ type = %s, name = %s, value = %s ]\n" % (i, itms[i].type, itms[i].name, itms[i].value)
        else:
          res += "%d: type = %s, name = %s, value = %s\n" % (i, itms[i].type, itms[i].name, itms[i].value)
    res += "=========================\n"

    return res

  def findInput(self, key):
    itms = self.getElementsByTagName('input')
    res += "==== List of anchors ====\n"
    for i in range(itms.length):
      if itms[i].type != 'hidden' :
        info = itms[i].value
        if findString(info, key)  :
          res += "%d: type = %s, name = %s, value = %s\n" % (i, itms[i].type, itms[i].name, itms[i].value)
    res += "=========================\n"
  
    return res

  def getInputByIndex(self, n):
    itms = self.getElementsByTagName('input')
    if n >=0 and n < itms.length :
      return itms[n]
    else:
      print "Invaild Index"
      return None

  def clickInputByIndex(self, n):
    itm = self.getInputByIndex(n)
    if itm : itm.click()

  def getButton(self, val):
    itms = self.getElementsByTagName('input')
    for i in range(itms.length):
      if itms[i].type == 'button' or itms[i].type == 'submit' :
        info = itms[i].value
        if findString(info,val)  :
          return itms[i]
    print "Button: %s not found." % (val)
    return None

  def clickButtonByValue(self, val):
    itm = self.getButton(val)
    if itm : itm.click()

  def clickButton(self, x):
    if type(x) == types.IntType :
      self.clickInputByIndex(x)
    elif type(x) == types.StringType or type(x) == types.UnicodeType :
      self.clickButtonByValue(x)


  def clickSubmit(self, val):
    itm = self.getButton(val)
    if itm and itm.type == 'submit': itm.click()


  def clickInputByValue(self, val):
    itms = self.getElementsByTagName('input')
    for i in range(itms.length):
      if itms[i].value == val:
        itms[i].click()
        return

  def getSubmits(self):
    res = []
    itms = self.getElementsByTagName('input')
    for i in range(itms.length):
      if itms[i].type == 'submit':
        res.append(itms[i])
    return res

  def setValue(self, name, val):
    in_obj = self.getElementByName('input', name)
    if in_obj :
      in_obj.value = val
      return in_obj.value
    return None

#---- Control
  def GoHome(self):
    self.ie.GoHome()
    return

  def GoBack(self):
    self.ie.GoBack()
    return

  def GoForward(self):
    self.ie.GoForward()
    return

  def Stop(self):
    self.ie.Stop()
    return

  def ScrollV(self, val=100):
    self.ie.Document.ParentWindow.scrollBy(0, val)
    return

  def ScrollH(self, val=100):
    self.ie.Document.ParentWindow.scrollBy(val,0)
    return

  def ScrollTo(self, h, v):
    self.ie.Document.ParentWindow.scrollTo(h,v)
    return

  def Quit(self):
    self.ie.Quit()
    self.ie = None
    return

  def FullScreen(self):
    self.ie.FullScreen = not self.ie.FullScreen
    return


#---- sample navigation
def main():
  ie=ie_com()
  ie.navigate(home)

if __name__ == "__main__":
  main()

