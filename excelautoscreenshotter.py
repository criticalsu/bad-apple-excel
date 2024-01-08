import os
import time
import pyautogui

start = time.time()

#Do it one chunk at a time in case something goes wrong, this thing is slow asf
#Close all apps before running, else the excel loading is going to be slow
for file in os.listdir("C:/Users/ACER/Documents/GitHub/bad-apple-excel/frames")[:10]:
    #Automate taking screenshot, delays need to be adjusted
    os.startfile(f"C:/Users/ACER/Documents/GitHub/bad-apple-excel/frames/{file}")
    time.sleep(1.2) #Waiting for excel to fully open
    pyautogui.click(1342, 752) #click the zoom
    time.sleep(0.05)
    pyautogui.click(170, 678, clicks=2) #click the custom zoom digit
    time.sleep(0.05)
    pyautogui.press('backspace') #delete 1 zero (100->10)
    time.sleep(0.05)
    pyautogui.click(112, 713) #click ok
    time.sleep(0.05)
    pyautogui.keyDown("ctrl") #hide hud
    pyautogui.keyDown("shift")
    pyautogui.press("f1")
    time.sleep(0.1)
    pyautogui.screenshot(fr"screenshot/{file[:-5]}.png") #take screenshot
    time.sleep(0.05)
    pyautogui.press("f1") #reenable hud for the next file
    pyautogui.keyUp("ctrl")
    pyautogui.keyUp("shift")
    time.sleep(0.05)
    pyautogui.click(1342, 12) #click the close button, can't be bothered to use os.kill or whatever

print(f"Elapsed: {time.time() - start}")