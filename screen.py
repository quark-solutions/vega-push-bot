import pyautogui
import time

try:
    pyautogui.FAILSAFE = True
except Exception as e:
    print(f"Erro ao configurar FAILSAFE: {e}")

print("Mova o mouse para a posição desejada e pressione Ctrl+C para sair...")

try:
    while True:
        x, y = pyautogui.position()

        print(f"X: {x} | Y: {y}   ", end="\r", flush=True)
        time.sleep(0.1)
except KeyboardInterrupt:
    print("\nCoordenadas salvas!")
