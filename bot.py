import pyautogui
import time
import os
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook
import pytesseract
from PIL import Image
from tqdm import tqdm

DELAY_TO_ACTIONS = 1
WAITING_TIME = 20

FILE_NAME = "Basket.xlsx"
SHEET_NAME = "Main"

screen_width, screen_height = pyautogui.size()
region_width = int(screen_width * 0.5)
region_height = int(screen_height * 0.5)
region_x = (screen_width - region_width) // 2
region_y = (screen_height - region_height) // 2

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def openFile(filename):
    dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(dir, filename)

    return pd.read_excel(file_path)

def update_excel(file_name, sheet_name, index, status):
    wb = load_workbook(file_name)
    sheet = wb[sheet_name]

    row_number = index + 2
    sheet[f"I{row_number}"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    sheet[f"J{row_number}"] = status

    wb.save(file_name)
    wb.close()

def check_for_errors():
    screenshot_path = "screenshot.png"
    pyautogui.screenshot(screenshot_path, region=(region_x, region_y, region_width, region_height))

    img = Image.open(screenshot_path)
    text = pytesseract.image_to_string(img)

    error_keywords = [
        "Você não atingiu o mínimo de 100 ações para compra!",
        "Cliente com Pendência CCL para operação!"
    ]

    for keyword in error_keywords:
        if keyword.lower() in text.lower():
            print(f"🚨 ERRO: {text.strip()}")
            return True

    return False

def push(sinacor, id, quantity):
    try:
        pyautogui.doubleClick(940, 410)

        for letter in sinacor:
          pyautogui.typewrite(letter)
          time.sleep(0.5)

          pyautogui.press('backspace')

        pyautogui.press('tab')

        for letter in id:
            pyautogui.typewrite(letter)
            time.sleep(0.5)

            pyautogui.press('backspace')

        pyautogui.press('tab', presses=3)

        for letter in quantity:
            pyautogui.typewrite(letter)
            time.sleep(0.5)

            pyautogui.press('backspace')

        # Button "Validar"
        pyautogui.click(2000, 555)
        time.sleep(WAITING_TIME)

        if check_for_errors():
            # Close Message Error 1
            pyautogui.click(1786, 873)
            time.sleep(DELAY_TO_ACTIONS)

            # Close Message Error 2
            pyautogui.click(1786, 873)
            time.sleep(DELAY_TO_ACTIONS)
            return "Erro ao enviar"

        # Button "Incluir"
        pyautogui.click(2000, 1265)
        time.sleep(DELAY_TO_ACTIONS)

        # Button "Enviar para Formalização", se existir
        pyautogui.click(1640, 1045)
        time.sleep(WAITING_TIME)

        # Close Success Message
        pyautogui.click(1786, 873)
        time.sleep(DELAY_TO_ACTIONS)

        print(f"Sinacor: {sinacor} | ID: {id} | Quantidade: {quantity}: Sucesso")
        return "Enviado"

    except Exception as e:
        return f"Erro: {str(e)}"

def main():
    try:
        data = openFile(FILE_NAME)

        rows = len(data)
        progress = 0

        for index, row in tqdm(data.iterrows(), total=rows, desc="Processando", unit="registros"):
            sinacor = str(row['Sinacor'])
            id = str(row['ID'])
            quantity = str(row['Quantidade tryd'])
            registered = str(row['Sent'])

            if registered == "nan":
              response = push(sinacor, id, quantity)

              progress += 1
              progress_percent = round((progress / rows) * 100, 2)
              print(f"📊 Progresso: {progress_percent:.2f}% concluído")

              update_excel(FILE_NAME, SHEET_NAME, index, response)

            time.sleep(DELAY_TO_ACTIONS)

        print("\n✅ Processamento concluído!")

    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    main()
