# Vindenatt 2023
# pyinstaller main.py --clean --onefile --noconsole to build

import time
from mss import mss
from multiprocessing import freeze_support, Process
import win32com.client as comclt
from win32gui import FindWindow, MoveWindow
import tkinter as tk
from tkinter import ttk

DELAY = 0.13 #smaller delay might be possible on fast machines

direction = True #True = left, False = right
pixel_upper1_L = None
pixel_upper1_R = None

proc = None


def get_pixel(left,top):
    with mss() as sct:
        pic = sct.grab({'mon':1, 'top':top, 'left':left, 'width':1, 'height':1})
        return pic.pixel(0,0)

def start_pixels():
    global pixel_upper1_L, pixel_upper1_R
    pixel_upper1_L = get_pixel(272,234) #remember the branch two "tiles" above the player on the left
    pixel_upper1_R = get_pixel(450,234) #remember the branch two "tiles" above the player on the right
        
def locate_branch(wsh):
    global pixel_upper1_L, pixel_upper1_R, direction
    if direction:
        current_pixel = get_pixel(272,234) #left branch check
        if pixel_upper1_L!=current_pixel:
            wsh.SendKeys("{LEFT}")
            time.sleep(DELAY) #might get better score by tweaking this
            wsh.SendKeys("{RIGHT}")
            direction = False
        else:
            wsh.SendKeys("{LEFT}")
    else:
        current_pixel = get_pixel(450,234) #right branch check
        if pixel_upper1_R!=current_pixel:
            wsh.SendKeys("{RIGHT}")
            time.sleep(DELAY) #might get better score by tweaking this
            wsh.SendKeys("{LEFT}")
            direction = True
        else:
            wsh.SendKeys("{RIGHT}")

    time.sleep(DELAY)

def shell_main():
    end_this = False
    try:
        wsh = comclt.Dispatch("WScript.Shell")
        wsh.AppActivate("Timberman")

        wsh.SendKeys("{ENTER}")
        time.sleep(1.5)
        wsh.SendKeys("{LEFT}")
        start_pixels()
    except:
        end_this = True
    
    while not end_this:
        locate_branch(wsh)
    
            

def act_start():
    if FindWindow(None, "Timberman"):
        game_win = FindWindow(None, "Timberman")
        MoveWindow(game_win,0,0,720,480,True)
        global proc
        proc = Process(target=shell_main)
        proc.start()
        btn_start.config(command=act_stop,text = "Stop")
        lbl_status.config(text = "Started successfully.")
    else:
        lbl_status.config(text = "Timberman is not launched.")


        
def act_stop():
    if proc:
        proc.terminate()
    btn_start.config(command=act_start,text = "Start")
    lbl_status.config(text = "Stopped successfully.")

def act_exit():
    if proc:
        proc.terminate()
    root.destroy()

if __name__ == '__main__':
    freeze_support()



    root = tk.Tk()
    root.overrideredirect(True)
    root.geometry("+0+480")

    frm = ttk.Frame(root, padding=18)
    frm.grid()

    lbl_title = ttk.Label(frm,text="TimberBotty",font="Helvetica 18 bold")
    lbl_title.grid(row=0,column=0,columnspan=2)
    lbl_desc = tk.Label(frm,text=" To use the bot:\n・Launch Timberman\n・Have it run in Windowed (720x480)\n・Press \"Singleplayer\", then \"Start\" here",font="Helvetica 14",justify="left")
    lbl_desc.grid(row=1,column=0,columnspan=2,pady=10)
    btn_start = tk.Button(frm, text="Start", command=act_start,font="sans 14 bold")
    btn_start.grid(row=2,column=0)
    btn_exit = tk.Button(frm, text="Exit", command=act_exit,font="sans 14 bold")
    btn_exit.grid(row=2,column=1,pady=6)
    lbl_status = ttk.Label(frm,text="Press \"Start\" if you're at the ▶ button.",font="Helvetica 14 italic",justify="left",background="lightgray")
    lbl_status.grid(row=3,column=0,columnspan=2,sticky = tk.W+tk.E,pady=4)
    
    root.mainloop()
    