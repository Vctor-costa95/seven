import pyautogui

x ,y = pyautogui.locateCenterOnScreen("CE.png",confidence=0.7)
pyautogui.moveTo(x,y,duration=0.5)
pyautogui.click()

x ,y = pyautogui.locateCenterOnScreen("9.png",confidence=0.7)
pyautogui.moveTo(x,y,duration=0.5)
pyautogui.click()

x ,y = pyautogui.locateCenterOnScreen("vezes.png",confidence=0.8)
pyautogui.moveTo(x,y,duration=0.5)
pyautogui.click()

x ,y = pyautogui.locateCenterOnScreen("9.png",confidence=0.7)
pyautogui.moveTo(x,y,duration=0.5)
pyautogui.click()

x ,y = pyautogui.locateCenterOnScreen("igual.png",confidence=0.7)
pyautogui.moveTo(x,y,duration=0.5)
pyautogui.click()