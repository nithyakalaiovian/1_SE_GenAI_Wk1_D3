import pyautogui
import time
import pyscreeze

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

print("Step1 :Open the Browser")
time.sleep(1)  # Wait for 1 seconds to allow the user to open the browser

pyautogui.hotkey('win', 'r')  # Open a new tab
time.sleep(1)  # Wait for 1 second
pyautogui.typewrite('msedge')  # Type 'msedge' in the run dialog
time.sleep(1)  # Wait for 1 second
pyautogui.press('enter')  # Press Enter to open msedge
time.sleep(1)  # Wait for 1 second
pyautogui.hotkey('ctrl', 't')  # Open a new tab

pyautogui.typewrite('https://www.grtjewels.com')  # Type the URL
time.sleep(1)  # Wait for 1 second'
pyautogui.press('enter')  # Press Enter to open msedge
time.sleep(5)  # Wait for 5 seconds


pyautogui.screenshot().save("screenshot.png") #save the screenshot of the current screen
time.sleep(1)  # Wait for 5 seconds

pyautogui.hotkey('alt', 'f4')  # close current browser window'