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

multi = False


def get_pixel(left,top):
    with mss() as sct:
        pic = sct.grab({'mon':1, 'top':top, 'left':left, 'width':1, 'height':1})
        return pic.pixel(0,0)

def start_pixels():
    global pixel_upper1_L, pixel_upper1_R
    pixel_upper1_L = get_pixel(272,234) #remember the branch two "tiles" above the player on the left
    pixel_upper1_R = get_pixel(450,234) #remember the branch two "tiles" above the player on the right
        
def locate_branch(wsh,delay):
    global pixel_upper1_L, pixel_upper1_R, direction
    if direction:
        current_pixel = get_pixel(272,234) #left branch check
        if pixel_upper1_L!=current_pixel:
            wsh.SendKeys("{LEFT}")
            time.sleep(delay) #might get better score by tweaking this
            wsh.SendKeys("{RIGHT}")
            direction = False
        else:
            wsh.SendKeys("{LEFT}")
    else:
        current_pixel = get_pixel(450,234) #right branch check
        if pixel_upper1_R!=current_pixel:
            wsh.SendKeys("{RIGHT}")
            time.sleep(delay) #might get better score by tweaking this
            wsh.SendKeys("{LEFT}")
            direction = True
        else:
            wsh.SendKeys("{RIGHT}")

    time.sleep(delay)

def shell_main(delay):
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
        locate_branch(wsh,delay)
    
            

def act_start():
    if FindWindow(None, "Timberman"):
        game_win = FindWindow(None, "Timberman")
        MoveWindow(game_win,0,0,720,480,True)
        global proc
        proc = Process(target=shell_main,args=(DELAY,))
        proc.start()
        btn_start.config(command=act_stop,text = "Stop")
        lbl_status.config(text = "Started successfully.")
    else:
        lbl_status.config(text = "Timberman is not launched.")

def on_delay_change(e):
    global DELAY,multi
    if multi:
        DELAY = float(e)*10
    else:
        DELAY = float(e)

def set_delay_multi():
    global DELAY,multi
    if multi:
        btn_delay_multi.config(relief=tk.RAISED)
        multi = False
        DELAY /= 10
    else:
        btn_delay_multi.config(relief=tk.SUNKEN)
        multi = True
        DELAY *= 10
        
def act_stop():
    try:
        proc.terminate()
    except Exception:
        lbl_status.config(text = "Process termination error.")
    btn_start.config(command=act_start,text = "Start")
    lbl_status.config(text = "Stopped successfully.")
    proc.join()

def act_exit():
    try:
        proc.terminate()
    except Exception:
        lbl_status.config(text = "Process termination error.")
    root.destroy()

if __name__ == '__main__':
    freeze_support()



    root = tk.Tk()
    root.title('TimberBotty')
    #root.overrideredirect(True)
    root.resizable(0,0)
    root.geometry("+0+480")

    frm = ttk.Frame(root, padding=18)
    frm.grid()

    lbl_title = ttk.Label(frm,text="TimberBotty",font="Helvetica 18 bold")
    lbl_title.grid(row=0,column=0,columnspan=2)
    lbl_desc = tk.Label(frm,text=" To use the bot:\n・Launch Timberman\n・Make sure it's in Windowed mode\n・Press \"Singleplayer\", then \"Start\" here",font="Helvetica 14",justify="left")
    lbl_desc.grid(row=1,column=0,columnspan=2,pady=10)
    btn_start = tk.Button(frm, text="Start", command=act_start,font="sans 14 bold")
    btn_start.grid(row=3,column=0)
    lbl_title = ttk.Label(frm,text="Delay:",font="Helvetica 14")
    lbl_title.grid(row=2,column=0,pady=15)
    scl_delay = tk.Scale(frm, from_=0.01, to=0.25, resolution=0.01, orient=tk.HORIZONTAL,command=on_delay_change)
    scl_delay.set(DELAY)
    scl_delay.place(x=133,y=147)
    btn_delay_multi = tk.Button(frm, text="x10", command=set_delay_multi,font="sans 14")
    btn_delay_multi.place(x=235,y=155)
    btn_exit = tk.Button(frm, text="Exit", command=act_exit,font="sans 14 bold")
    btn_exit.grid(row=3,column=1,pady=6)
    lbl_status = ttk.Label(frm,text="Press \"Start\" if you're at the ▶ button.",font="Helvetica 14 italic",justify="left",background="lightgray")
    lbl_status.grid(row=4,column=0,columnspan=2,sticky = tk.W+tk.E,pady=4)
    
    root.mainloop()
    