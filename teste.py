import pyautogui
import time

pyautogui.hotkey('win', 'e')
time.sleep(2)
pyautogui.hotkey('shift', 'tab')
time.sleep(2)
for _ in range(3):  
    pyautogui.press('down')
pyautogui.press('enter')
time.sleep(1)
# Pasta documents
pyautogui.hotkey('ctrl', 'e')
pyautogui.hotkey('shift', 'tab')
pyautogui.press('right')
pyautogui.hotkey('shift', 'down')
for _ in range(6):
    pyautogui.press('down')
pyautogui.press('enter')
time.sleep(1)
# Pasta processos
pyautogui.hotkey('ctrl', 'e')
pyautogui.hotkey('shift', 'tab')
for _ in range(2):
    pyautogui.press('right')                
for _ in range(5):
    pyautogui.press('down') 
pyautogui.press('enter')
time.sleep(1)
# Pasta envio_wpp
pyautogui.hotkey('ctrl', 'e')
time.sleep(1)
pyautogui.write('tabela_placas_ativas.png')
time.sleep(2)
for _ in range(4):
    pyautogui.press('tab')
time.sleep(1)
pyautogui.press('down')
time.sleep(1)
pyautogui.hotkey('ctrl', 'c')
time.sleep(1)
#whatsapp
pyautogui.press('win')
time.sleep(1)
pyautogui.write('whatsapp')
time.sleep(1)
pyautogui.press('enter')  
time.sleep(4)
pyautogui.write('61981109691')  
time.sleep(1)
pyautogui.press('down')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.hotkey('ctrl', 'v')
time.sleep(1)
for _ in range(4):
    pyautogui.press('tab')
time.sleep(1)
pyautogui.press('enter')

