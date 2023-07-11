
import ctypes

# Load the user32.dll library
user32 = ctypes.windll.LoadLibrary("user32.dll")

# Send the WM_INPUTLANGCHANGEREQUEST message to change the keyboard layout
user32.SendMessageW(
    ctypes.c_void_p(0xffff), # HWND_BROADCAST
    0x0050, # WM_INPUTLANGCHANGEREQUEST
    0, # wParam (not used)
    ctypes.c_long(0x04090409).value # lParam: 0x0409 = English (United States)
)

import sys
import pyautogui



from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtGui import QIcon,QDoubleValidator,QValidator

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIntValidator
from PyQt5.QtWidgets import QApplication, QLineEdit, QMainWindow
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget
from PyQt5.QtGui import QFont
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication, QMainWindow, QAction, QToolBar



from turtle import delay
import time
from pynput import mouse , keyboard             

from pynput.keyboard import Key, Controller





def window():
  m=mouse.Controller()
  kb = Controller()




  
  app = QApplication(sys.argv)
  win = QMainWindow()
  win.setGeometry(400,600,300,400)
  win.setWindowTitle("AKP(AutoKomodaaPanel)")
  win.setToolTip("Komodaa")
  win.setWindowIcon(QIcon("icon.ico"))
  win.setFixedSize(325,400)



  font_size = QFont()
  font_size.setPointSize(13)
  font_bold = QFont()
  font_bold.setBold(True)
  font_month = QFont()
  font_month.setPointSize(12)
  


  lbl_komoda = QtWidgets.QLabel(win)
  lbl_komoda.setText("سیستم پنل اتوماتیک کمدا")
  lbl_komoda.move(50,30)
  lbl_komoda.setFixedWidth(210)
  lbl_komoda.setFont(font_bold)
  lbl_komoda.setFont(font_size)
 

  

  lbl_ckomoda = QtWidgets.QLabel(win)
  lbl_ckomoda.setText('تعداد کمدا')
  lbl_ckomoda.move(150,80)
  
  lbl_ckomoda.setFont(font_month)
  lbl_ckomoda.setFixedSize(125,30)

  lbl_year = QtWidgets.QLabel(win)
  lbl_year.setText(' سال')
  lbl_year.move(162,120)
  
  lbl_year.setFont(font_month)
 

  lbl_month = QtWidgets.QLabel(win)
  lbl_month.setText('ماه ')
  lbl_month.move(151,160)
  lbl_month.setFont(font_bold)
  lbl_month.setFont(font_month)
  

  lbl_day = QtWidgets.QLabel(win)
  lbl_day.setText(' روز')
  lbl_day.move(154,200)
  lbl_day.setFont(font_month)
  

  lbl_hour = QtWidgets.QLabel(win)
  lbl_hour.setText(' ساعت')
  lbl_hour.move(168,240)
  lbl_hour.setFont(font_bold)
  lbl_hour.setFont(font_month)
  

  

  txt_komoda = QtWidgets.QLineEdit(win)
  txt_komoda.setPlaceholderText("140")
  txt_komoda.move(100,80)
  

  txt_year = QtWidgets.QLineEdit(win)
  txt_year.setPlaceholderText("1376")
  txt_year.move(100,120)
  

  txt_month = QtWidgets.QLineEdit(win)
  txt_month.move(100,160)
  txt_month.setPlaceholderText("09")

  txt_day = QtWidgets.QLineEdit(win)
  txt_day.move(100,200)
  txt_day.setPlaceholderText("09")

  txt_hour = QtWidgets.QLineEdit(win)
  txt_hour.move(100,240)
  txt_hour.setPlaceholderText("21")

  def excelmouse():

    m.position = (300,500)
    m.click(mouse.Button.left,1)
  def panelmouse():
    m.position = (1300,150)
    m.click(mouse.Button.left,1)
    time.sleep(0.1)
  def panelmouse2():
    m.position = (1129,15)
    m.click(mouse.Button.left,1)
    time.sleep(0.1)
  def findfirst():
    confidence_excel = 0.9
    condition_excel = False
    while not condition_excel :
      location = pyautogui.locateOnScreen("excel.png", confidence=confidence_excel)
      
      if location is not None: 
        condition_excel = True
        print("true")
        print("excel dide shod")
        print("Image found at: ", location)

        esc()

        kb.press(Key.ctrl)
        kb.press(Key.left)
        kb.release(Key.left)
        time.sleep(0.1)
        kb.press(Key.left)
        kb.release(Key.left)
        kb.release(Key.ctrl)
        
        kb.press(Key.ctrl)
        kb.press(Key.up)
        kb.release(Key.up)
        time.sleep(0.1)
        kb.press(Key.up)
        kb.release(Key.up)
        kb.release(Key.ctrl)

        kb.press(Key.right)
        kb.release(Key.right)

        kb.press(Key.down)
        kb.release(Key.down)
        time.sleep(0.1)

      else:
        print("false")
        print("excel hanouz dide nashode")
        pass
      
    esc()
    kb.press(Key.right)
    kb.release(Key.right)
    kb.press(Key.ctrl)
    kb.press(Key.up)
    kb.release(Key.up)
    time.sleep(0.1)
    kb.press(Key.up)
    kb.release(Key.up)
    kb.release(Key.ctrl)
    time.sleep(0.1)
    kb.press(Key.down)
    kb.release(Key.down)
    time.sleep(0.1)  
  def timeinsert():
    delete()
    day=txt_day.text()
    month=txt_month.text()
    year=txt_year.text()
    hour=txt_hour.text()
    kb.type(day)
    kb.type(month)
    kb.type(year)
    #insert time
    tab()
    time.sleep(0.1)

    kb.type(hour)
    tab()
    time.sleep(0.1)
    kb.type("0")
    tab()
  def copy():

    kb.press(Key.ctrl)
    kb.press("c")
    kb.release("c")
    kb.release(Key.ctrl)
    time.sleep(0.1)
  def paste():
    kb.press(Key.ctrl)
    kb.press("v")
    kb.release("v")
    kb.release(Key.ctrl)
    time.sleep(0.1)
  def enter(): 
    kb.press(Key.enter)
    kb.release(Key.enter)
    time.sleep(0.5)
  def tab():
    kb.press(Key.tab)
    kb.release(Key.tab)
    
  def shifttab():
    kb.press(Key.shift)
    kb.press(Key.tab)
    kb.release(Key.tab)
    kb.release(Key.shift)
    time.sleep(0.01)
  def delete():
    kb.press(Key.delete)
    kb.release(Key.delete)
  def down():
    kb.press(Key.down)
    kb.release(Key.down)
    time.sleep(0.1)
  def altTab():
    kb.press(Key.alt)
    kb.press(Key.tab)
    kb.release(Key.tab)
    kb.release(Key.alt)
  def ctrlTab():
    kb.press(Key.ctrl)
    tab()
    kb.release(Key.ctrl)
    time.sleep(0.1)

  
  
 

  def esc():
    kb.press(Key.esc)
    kb.release(Key.esc)

  def changeType():
    kb.press(Key.ctrl)
    kb.press("b")
    kb.release("b")
    kb.release(Key.ctrl)
    time.sleep(0.1)
    kb.press(Key.ctrl)
    kb.press("u")
    kb.release("u")
    kb.release(Key.ctrl)
    time.sleep(0.1)
    kb.press(Key.ctrl)
    kb.press("i")
    kb.release("i")
    kb.release(Key.ctrl)
    time.sleep(0.1)
    kb.press(Key.ctrl)
    kb.press(Key.shift)
    kb.press("e")
    kb.release("e")
    kb.release(Key.shift)
    kb.release(Key.ctrl)
  


  def cancelNote():
    
    cancelnote = pyautogui.locateOnScreen("cancelnote.png")
    if cancelnote is not None:
      editese = pyautogui.center(cancelnote)
      pyautogui.moveTo(editese)
      print("cancel hast")
      esc()
      ctrlTab()
      print("badi")
      changeType()
      down()
      print("payeen")
      
    else:
      print("cancel  nist")
      ctrlTab()
      time.sleep(0.1)
      down()
      time.sleep(0.1)
      print("payeen")
  
  def allow_First():
    
    time.sleep(0.5)  
    for i in range(7):
      tab()
    enter()
    for i in range(9):
      tab()
    timeinsert()
    for i in range(11):
      shifttab() 
    enter()
   
    
  
#-------------------------------------------------------------------------------
  def startKomodaSite():
        print("start komoda zade shod")
        m.position = (300,500)
        m.click(mouse.Button.left,1)
        
        
        kb.press(Key.ctrl)
        kb.press("t")
        kb.release("t")
        kb.release(Key.ctrl)
        time.sleep(0.1)
        kb.type("admin.komodaa.com/admin/dkLogistic")
        print("adres vared shod")
        time.sleep(2)
        enter()
        confidence = 0.9
        condition = False
        x=1
        while not condition :
          location = pyautogui.locateOnScreen("panel.png", confidence=confidence)
          
          if location is not None: 
            condition = True
            print("true")
            print("panel dide shod")
            print("Image found at: ", location)
            time.sleep(2)
            ctrlTab()
          else:
            if x < 20:
              print("false")
              print("panel hanouz dide nashode")
              x=x+1
              pass
            else:
              sys.exit()

            
            
  
        


  def firstCodeInsert():
  

    print("avalin soro")
    findfirst()
    print("avalin seloul peyda shod")
    copy()
    ctrlTab()
    for i in range(2):
      shifttab()
    paste()
    enter()
    time.sleep(0.1)
    print("code vared shod")
    allow_First()
    print("tarikh zade shod")
    cancelNote()
    print("cancel y raf shod")

  def secondCodeInsert():
    komoda_int = txt_komoda.text()
    komoda = int(komoda_int)-1
    
    print("dovomin shorou")
    for i in range(komoda):
      confidence_excel2 = 0.9
      condition_excel2 = False
      while not condition_excel2 :
          location = pyautogui.locateOnScreen("excel.png", confidence=confidence_excel2)
          
          if location is not None: 
            condition_excel2 = True
            print("true")
            print("excel2 dide shod")
            print("Image found at: ", location)

            copy()
            ctrlTab()

            for y in range(5):
              shifttab()
            paste()
            enter()
            print("code vared shod2")
            allow_First()
            print("tarikh zade shod2")
            cancelNote()
            print("cancel  raf shod2")
          else:
            print("false")
            print("panel hanouz dide nashode")
            pass

     
      
      

  def startAPK():
    startKomodaSite()
    firstCodeInsert()
    secondCodeInsert()
    print("barname tamououououm shod")


  btn_save = QtWidgets.QPushButton(win)
  btn_save.setText('ثبت خودکار')
  btn_save.clicked.connect(startAPK)
  btn_save.move(100,300)

  
  win.show()
  sys.exit(app.exec_())

window()