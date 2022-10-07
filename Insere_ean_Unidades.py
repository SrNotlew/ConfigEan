from csv import excel
import time
from openpyxl import load_workbook
import pyautogui
import pyscreenshot as ImageGrab
import pytesseract
import cv2


wb = load_workbook(filename='./Planilhas/ENTRADAS.xlsx')

def sleepI():
    time.sleep(0.4)

def sleepB():
    time.sleep(0.7)

# DEFININDO PATH TESSERACT
pytesseract.pytesseract.tesseract_cmd = "C:\\Users\\igor.goncalves\\AppData\\Local\\Tesseract-OCR\\Tesseract.exe"

# LACO DE REPETICAO LISTA ACESSA PLANILHA
i = 2
sheet_ranges = wb['Planilha1']
for description in sheet_ranges:
    description = [
    sheet_ranges['C{}'.format(i)].value,     # 0    CODIGO PRODUTO
    sheet_ranges['G{}'.format(i)].value,     # 1    EAN PRODUTO
    ]
    #BUSCAR 
    pyautogui.moveTo(439, 684)
    pyautogui.click()
    sleepI()
    #COLOCA 
    pyautogui.moveTo(520, 475, 0.2)
    pyautogui.click()
    sleepI()
    pyautogui.typewrite(str(description[0]))
    sleepI()
    #localiza
    pyautogui.moveTo(375, 629)
    pyautogui.click()
    sleepI()
    pyautogui.press('SPACE')
    time.sleep(3.5)
    pyautogui.moveTo(346, 164, 0.3) #CLICA EM CARACTERISTICAS
    time.sleep(1.5)
    pyautogui.click()
    #move ate UNIDADES
    pyautogui.moveTo(1013, 367, 0.3)
    pyautogui.click()
    sleepB()

    # TIRAR PRINT
    im = ImageGrab.grab(bbox=(238, 438, 463, 491))
    im.save(".\Prints\\box.png")#salvar
    #LER A IMGAEM SALVA
    printAbrir = cv2.imread(".\Prints\\box.png")
    #Reconhecer IMAGEM
    resultado = pytesseract.image_to_string(printAbrir)

    print(resultado.strip())
    if str(resultado.strip()) == "‘UNIDADE":
        pass
    else:
        #move ate INCLUI NOVO CODIGO
        pyautogui.moveTo(672, 568, 0.3)
        pyautogui.click()
        time.sleep(1)
        pyautogui.press('TAB')
        pyautogui.press('space')

        #move ate UNIDADE DE VENDA
        pyautogui.moveTo(461, 451, 0.3)
        pyautogui.click()
        sleepB()
        pyautogui.click()
        pyautogui.press('DOWN',16)
        #TIPO DE BARRA
        pyautogui.press('TAB')
        sleepI()
        pyautogui.press('DOWN')
        sleepI()
        #codigo de barra
        pyautogui.press('TAB')
        sleepI()
        pyautogui.typewrite(str(description[1]))
        sleepI()
        #EFETIVA
        pyautogui.moveTo(1122, 607, 0.3)
        pyautogui.click()
        sleepI()
        pyautogui.press('LEFT')
        sleepI()
        pyautogui.press('ENTER')
    
    time.sleep(2)
    sleepI()
    #fecha UNIDADEs
    pyautogui.moveTo(1155, 119, 0.3)
    pyautogui.click()
    sleepI()
    pyautogui.click()
    sleepB()

    #ALTERA EM COMPLEMENTARES 
    pyautogui.moveTo(526, 162, 0.3) #CLICA COMPLEMTENTARES
    pyautogui.click()
    sleepI()
    pyautogui.click()
    # TIRAR PRINT
    im = ImageGrab.grab(bbox=(474, 258, 583, 284))
    im.save(".\Prints\\box.png")#salvar
    #LER A IMGAEM SALVA
    printAbrir = cv2.imread(".\Prints\\box.png")
    #Reconhecer IMAGEM
    resultado = pytesseract.image_to_string(printAbrir)

    print(resultado.strip())
    if str(resultado.strip()) == 'EAN-13':
        pyautogui.moveTo(976, 684, 0.3) 
        pyautogui.click()

    else:
        pyautogui.moveTo(383, 684, 0.3)
        pyautogui.click()
        time.sleep(2)
        pyautogui.moveTo(451, 275, 0.3) #CLICA TIP EAN
        pyautogui.click()
        sleepI()
        pyautogui.moveTo(361, 335, 0.3) #CLICA TIP EAN
        pyautogui.click()
        sleepI()
        pyautogui.press('enter')
        pyautogui.press('TAB')
        sleepI()
        pyautogui.typewrite(str(description[1]))
        sleepI()
        pyautogui.press('TAB')
        pyautogui.press('enter')
        sleepI()
        pyautogui.moveTo(544, 275, 0.3) #CLICA TIP EAN
        pyautogui.click()
        sleepI()
        pyautogui.moveTo(524, 335, 0.3) #CLICA TIP EAN
        pyautogui.click()
        sleepI()
        pyautogui.press('TAB')
        sleepI()
        pyautogui.typewrite(str(description[1]))
        sleepI()
        pyautogui.moveTo(1033, 685, 0.3) #EFETIVA
        pyautogui.click()
        # CLICA OK WMS
        time.sleep(1.5)
        pyautogui.moveTo(683, 407, 0.3)
        sleepI()
        pyautogui.click()
        sleepI()
        time.sleep(3)
        # CLICA FECHAR WMS
        pyautogui.moveTo(1111, 44, 0.3)
        sleepI()
        pyautogui.click()
        sleepI()
        time.sleep(6)


    sleepB()
    print(description)
    
    i += 1
