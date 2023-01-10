# pyinstaller main.py --clean --add-data "img/*;img" --onefile --noconsole to build

import time, os, sys
import mss, cv2
import multiprocessing as mp
import numpy as np
import win32com.client as comclt
from win32gui import FindWindow, MoveWindow
from tkinter import *
from tkinter import ttk

READ_UPPER_PIXELS = False
DELAY = 0.13

direction = True #True = left
pixel_main_L = None
pixel_main_R = None
pixel_upper1_L = None
pixel_upper1_R = None
pixel_upper2_L = None
pixel_upper2_R = None

exit_cmd = False
paused = False
initial = True
proc = None

class Process(mp.Process):
    def __init__(self, *args, **kw):
        if hasattr(sys, 'frozen'):
            # We have to set original _MEIPASS2 value from sys._MEIPASS
            # to get --onefile mode working.
            os.putenv('_MEIPASS2', sys._MEIPASS)
        try:
            super().__init__(*args, **kw)
        finally:
            if hasattr(sys, 'frozen'):
                # On some platforms (e.g. AIX) 'os.unsetenv()' is not
                # available. In those cases we cannot delete the variable
                # but only set it to the empty string. The bootloader
                # can handle this case.
                if hasattr(os, 'unsetenv'):
                    os.unsetenv('_MEIPASS2')
                else:
                    os.putenv('_MEIPASS2', '')

def get_pixel(left,top):
    with mss.mss() as sct:
        pic = sct.grab({'mon':1, 'top':top, 'left':left, 'width':1, 'height':1})
        return pic.pixel(0,0)

def start_pixels():
    global pixel_main_L, pixel_main_R, pixel_upper1_L, pixel_upper1_R, pixel_upper2_L, pixel_upper2_R,wsh
    #pixel_main_L = get_pixel(1000,856) #lower left branch
    #pixel_main_R = get_pixel(1580,856) #lower right branch
    pixel_upper1_L = get_pixel(272,234) #middle left branch  #1000,656
    pixel_upper1_R = get_pixel(450,234) #middle right branch
    #pixel_upper2_L = get_pixel(1000,456) #upper left branch
    #pixel_upper2_R = get_pixel(1580,456) #upper left branch

    #start_time = time.time()
    #locate_branch(DELAY)
    #print("--- %s seconds ---" % (time.time() - start_time))
        
def dash():
    if direction:
        wsh.SendKeys("{LEFT}")
        wsh.SendKeys("{LEFT}")
        wsh.SendKeys("{LEFT}")
        wsh.SendKeys("{LEFT}")
    else:
        wsh.SendKeys("{RIGHT}")
        wsh.SendKeys("{RIGHT}")
        wsh.SendKeys("{RIGHT}")
        wsh.SendKeys("{RIGHT}")
        
def locate_branch(wsh):
    global pixel_main_L, pixel_main_R, pixel_upper1_L, pixel_upper1_R, pixel_upper2_L, pixel_upper2_R, direction
    if direction:
        current_pixel = get_pixel(272,234) #upper left branch check
        if pixel_upper1_L!=current_pixel:
            wsh.SendKeys("{LEFT}")
            time.sleep(DELAY)
            wsh.SendKeys("{RIGHT}")
            direction = False
        elif READ_UPPER_PIXELS:
            current_pixel_upper1 = get_pixel(1000,656) #middle left branch check
            if pixel_upper1_L==current_pixel_upper1:
                current_pixel_upper2 = get_pixel(1000,456) #upper left branch check
                if pixel_upper2_L==current_pixel_upper2:
                    wsh.SendKeys("{LEFT}")
                    time.sleep(0.005)
                    wsh.SendKeys("{LEFT}")
                    time.sleep(0.005)
                    wsh.SendKeys("{LEFT}")
                    print("triple L")
                else:
                    #print("single L 1")
                    wsh.SendKeys("{LEFT}")
            else:
                #print("single L 2")
                wsh.SendKeys("{LEFT}")
        else:
            wsh.SendKeys("{LEFT}")
    else:
        current_pixel = get_pixel(450,234) #lower right branch check
        if pixel_upper1_R!=current_pixel:
            wsh.SendKeys("{RIGHT}")
            time.sleep(DELAY)
            wsh.SendKeys("{LEFT}")
            direction = True
        elif READ_UPPER_PIXELS:
            current_pixel_upper1 = get_pixel(1580,656) #middle right branch check
            if pixel_upper1_R==current_pixel_upper1:
                current_pixel_upper2 = get_pixel(1580,456) #upper right branch check
                if pixel_upper2_R==current_pixel_upper2:
                    wsh.SendKeys("{RIGHT}")
                    time.sleep(0.005)
                    wsh.SendKeys("{RIGHT}")
                    time.sleep(0.005)
                    wsh.SendKeys("{RIGHT}")
                    print("triple R")
                else:
                    #print("single R 1")
                    wsh.SendKeys("{RIGHT}")
            else:
                #print("single R 2")
                wsh.SendKeys("{RIGHT}")
        else:
            wsh.SendKeys("{RIGHT}")
    time.sleep(DELAY)

def shell_main():
    
    wsh = comclt.Dispatch("WScript.Shell")
    act = wsh.AppActivate("Timberman")
    
    
    if act:
        game_win = FindWindow(None, "Timberman")
        MoveWindow(game_win,0,0,720,480,True)
    else:
        print("Timberman is not launched.")
        exit()
    
    START_PICTURE_cv2 = cv2.imread("img/start.png")
    with mss.mss() as sct:
        pic = sct.grab({'mon':1, 'top':0+415, 'left':0+325, 'width':80, 'height':80})
        pic_cv2 = cv2.cvtColor(np.array(pic), cv2.COLOR_RGB2BGR)

    #cv2.imshow('img',pic_cv2)
    #cv2.waitKey(0)
    #cv2.destroyAllWindows()

    comparison = cv2.matchTemplate(pic_cv2, START_PICTURE_cv2, cv2.TM_CCOEFF_NORMED)
    if (comparison >= 0.8).any():

        wsh.SendKeys("{ENTER}")
        time.sleep(1.5)
        wsh.SendKeys("{LEFT}")
        start_pixels()

        while True:
                locate_branch(wsh)
    else:
        print("No Start button detected.")
            
def act_start():
    global proc
    if initial:
        proc = Process(target=shell_main)
        proc.start()

        
def act_pause():
    if proc:
        proc.terminate()

def act_exit():
    if proc:
        proc.terminate()
    root.destroy()

if __name__ == '__main__':
    mp.freeze_support()
    os.chdir(sys._MEIPASS)

    root = Tk()
    root.overrideredirect(True)
    root.geometry("+0+480")

    root.grid()
    ttk.Button(root, text="Start", command=act_start).grid(column=0, row=0)
    ttk.Button(root, text="Pause", command=act_pause).grid(column=1, row=0)
    ttk.Button(root, text="Exit", command=act_exit).grid(column=2, row=0)
    
    root.mainloop()
    
        
    

    print("Exited normally at",time.strftime("%H:%M:%S"))