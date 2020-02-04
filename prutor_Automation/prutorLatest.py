from PIL import ImageGrab
import os
import time
import win32api, win32con
from PIL import ImageOps
from PIL import ImageGrab
from numpy import *
import math
import win32com.client
import xlrd 

    
#################mouse controls###################
    
def leftClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)

def rightClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,0,0)
    time.sleep(1)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,0,0)    

def doubleClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
    
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)
    time.sleep(2)
    

def leftDown():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(.1)
    print('left down')

def leftUp():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(.1)
    print('left release')

def mousePos(cord):
    win32api.SetCursorPos((cord[0],cord[1]))

def get_cords():
    x,y=win32api.GetCursorPos()
    print(x,y)

###################start telling########################
def set_screen():
    count=0
    
    while True:
      if (count==30):
          break
      else:
          get_cords()
    count=count+1
    
####################################### Action definition ####################################3    
def passSet(sch,initial,initialschno):
   mousePos((872,505))
   rightClick()
   time.sleep(1)
   shellb = win32com.client.Dispatch("WScript.Shell")
   shellb.SendKeys('{DOWN}')
   time.sleep(0.3)
   shellb = win32com.client.Dispatch("WScript.Shell")
   shellb.SendKeys('{DOWN}')
   time.sleep(0.3)
   shellb = win32com.client.Dispatch("WScript.Shell")
   shellb.SendKeys('{DOWN}')
   time.sleep(0.3)
   shellb = win32com.client.Dispatch("WScript.Shell")
   shellb.SendKeys('{ENTER}')
   time.sleep(0.3)

   while True:
      pass_window = Chk_pass_window_enabled()
      if(pass_window==1):
        password = FetchXLSchPass(sch,initial,initialschno)
        mousePos((575,494))
        leftClick()
        time.sleep(0.3)
        shellb = win32com.client.Dispatch("WScript.Shell")
        shellb.SendKeys(password)
        time.sleep(0.3)
        shellb = win32com.client.Dispatch("WScript.Shell")
        shellb.SendKeys('{TAB}')
        time.sleep(0.3)
        shellb = win32com.client.Dispatch("WScript.Shell")
        shellb.SendKeys(password)
        time.sleep(0.3)
        shellb = win32com.client.Dispatch("WScript.Shell")
        shellb.SendKeys('{TAB}')
        time.sleep(0.3)
        shellb = win32com.client.Dispatch("WScript.Shell")
        shellb.SendKeys('{ENTER}')
        time.sleep(1)
        print("Password for sch " + str(sch) + " changed." )
        while True:
          close_win = Chk_closing_window_after_passSet()
          if(close_win==1):
            mousePos((1339,11))
            leftClick()
            time.sleep(0.3)
            break
          else:
            continue
        break
      else:
        print("Password window not visible")
        continue

def click_new_student_btn():
    mousePos((673,334))
    leftClick()
    time.sleep(1)

def revState():
    selectBtn_active = Chk_selectBtn_if_null_name()
    if(selectBtn_active==1):
        mousePos((654,335))
        leftClick()
        time.sleep(1)
        mousePos((520,335))
        leftClick()
        time.sleep(1)
        
def enter_fullname(schNo,initial,initialschno,section):
    #sch = initial+schNo
    sch = initial + str(schNo)
    #email = sch+'test6@gmail.com'
    #stuName = 'student' + sch
    mousePos((804,410))
    leftClick()
    time.sleep(0.5)

    SchName = FetchXLSchName(schNo,initial,initialschno,section)

    if(SchName == "null"):
        return 0
    SchEmail = FetchXLEmail(schNo,initial,initialschno)
    
    shellb = win32com.client.Dispatch("WScript.Shell")
    shellb.SendKeys(SchName)
    
    shellb = win32com.client.Dispatch("WScript.Shell")
    shellb.SendKeys('{TAB}')
    
    shellb = win32com.client.Dispatch("WScript.Shell")
    shellb.SendKeys(SchEmail)
    
    shellb = win32com.client.Dispatch("WScript.Shell")
    shellb.SendKeys('{TAB}')
    
    shellb = win32com.client.Dispatch("WScript.Shell")
    shellb.SendKeys(section)
    
    shellb = win32com.client.Dispatch("WScript.Shell")
    shellb.SendKeys('{TAB}')
    
    shellb = win32com.client.Dispatch("WScript.Shell")
    shellb.SendKeys(sch)

    shellb = win32com.client.Dispatch("WScript.Shell")
    shellb.SendKeys('{ENTER}')

    return 1

    #mousePos((797,615))
    #leftClick()
    #time.sleep(1)
    

###############################computer vision #####################################################################
def Chk_closing_window_after_passSet():
   box=(616,467,733,499)
   ## box=(608,293,728,325)
   im=ImageOps.grayscale(ImageGrab.grab(box))
   b=array(im.getcolors())
   b=b.sum()
   if(b==23438):
      return 1

def Chk_selectBtn_if_null_name():
   box=(741,318,849,353)
   ## box=(608,293,728,325)
   im=ImageOps.grayscale(ImageGrab.grab(box))
   b=array(im.getcolors())
   b=b.sum()
   if(b==8778):
      return 1    

def Chk_pass_window_enabled():
   box=(521,477,828,510)
   ## box=(608,293,728,325)
   im=ImageOps.grayscale(ImageGrab.grab(box))
   b=array(im.getcolors())
   b=b.sum()
   if(b==19855):
      return 1
   
def chk_student_btn():
    box=(609,318,726,348)
   ## box=(608,293,728,325)
    im=ImageOps.grayscale(ImageGrab.grab(box))
    b=array(im.getcolors())
    b=b.sum()
    if(b==14645):
       return 1
    

def chk_fullname_btn():
    box=(763,395,1302,427)
    ##box=(762,371,1300,398)
    im=ImageOps.grayscale(ImageGrab.grab(box))
    b=array(im.getcolors())
    b=b.sum()
    if(b==25688):
        return 1

def chk_save_btn():
    box=(763,630,838,661)
    ##box=(763,606,835,637)
    im=ImageOps.grayscale(ImageGrab.grab(box))
    b=array(im.getcolors())
    b=b.sum()
    if(b==9787):
        return 1

##########################################################################fetching detail from excel sheets#########################################################
def FetchXLSchName(schno,initial,initialschno):
   sch = initial + str(schno)
   loc = ("student.xlsx")
   wb = xlrd.open_workbook(loc) 
   sheet = wb.sheet_by_index(0)

   print(schno-initialschno)

   strxl=sheet.cell_value((schno-initialschno), 0)
   print(strxl)
   print(sch)
   if(strxl==sch):
       name = sheet.cell_value((schno-initialschno), 1)
       if(name=="null"):
           print("here in if")
           return "null"
       else:
           print("here")
           print(name)
           return name
   else:
       print("in else")

def FetchXLEmail(schno,initial,initialschno):
   sch = initial + str(schno)
   em = sch+"@manit.ac.in"
   loc = ("student.xlsx")
   wb = xlrd.open_workbook(loc) 
   sheet = wb.sheet_by_index(0)

   strxl=sheet.cell_value((schno-initialschno), 0)
   if(strxl==sch):
      return em 
      #return sheet.cell_value((schno-initialschno), 2)

def FetchXLSchPass(schno,initial,initialschno):
   sch = initial + str(schno)
   loc = ("student.xlsx")
   wb = xlrd.open_workbook(loc) 
   sheet = wb.sheet_by_index(0)

   strxl=sheet.cell_value((schno-initialschno), 0)
   if(strxl==sch):
      return sheet.cell_value((schno-initialschno), 3)
 

#############################main funtction#####################################

def set_motion():
    student_active = 0
    fullName_active= 0
    saveBtn_active = 0
    resName = 0
    initialschno = input("initialschno : ")
    initialschno = int(initialschno)
    schNo = initialschno
    section = input("section : ")
    lastRoll = input("How many times you want the process to repeat i.e; if first student is 181116201 and you enter 25 then last student will be 181116225 : ")
    res = input("Enter new initial(The fixed alpha part) - ")
    initial=res
    print("\nNOTE : Do remember to enter the data as it is entered into the excel table i.e; column(2,1) should contain the starting roll and after that all rolls shoud be ascending.\n\n")    
    Command_in=input("All set! Press 1 to begain; Press 2 to witheld;")
    if(Command_in=="1"):
     while True:   
          student_active = chk_student_btn()
          if(student_active==1):
                click_new_student_btn()
                fullName_active = chk_fullname_btn()
                if(fullName_active==1):
                    saveBtn_active = chk_save_btn()
                    if(saveBtn_active==1):
                        if(schNo<int(lastRoll)):
                          schNo = schNo + 1
                          resName = enter_fullname(schNo,initial,initialschno,section)
                          if(resName==0):
                              revState()
                              continue
                          print("Scholar No "+ str(schNo) + "added")
                          time.sleep(5)
                          while True:
                             student_active = chk_student_btn()
                             if(student_active==1):
                                passSet(schNo,initial,initialschno)
                                break
                             else: 
                                  continue
                          continue    
                        else:
                         break   
                    else:
                        print("save button not active")
                else:
                    print("Details fields not active")
          else:
                print("add student button not active")
    else:
       exit()

                    
######################main function#################


if __name__ == '__main__':
   
    set_motion()
