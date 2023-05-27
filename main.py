import pyautogui
import win32gui
import ctypes

KEY_CODE_ESCAPE = 0x1b

DNF_WINDOW_WIDTH = 1440
DNF_WINDOW_HEIGHT = 900

def initializeDnfWindow() -> None:
  # アラド戦記のウィンドウハンドルを取得する
  windowHandle = win32gui.FindWindow(None, 'アラド戦記')
  if windowHandle == 0:
    raise Exception('アラド戦記のウィンドウが見つかりませんでした')

  # 解像度チェック
  rect = win32gui.GetWindowRect(windowHandle)
  windowWidth = rect[2] - rect[0]
  windowHeight = rect[3] - rect[1]
  if windowWidth != DNF_WINDOW_WIDTH or windowHeight != DNF_WINDOW_HEIGHT:
    raise Exception(f'アラド戦記のウィンドウサイズが異常です\nアラド戦記のウィンドウサイズを{DNF_WINDOW_WIDTH}x{DNF_WINDOW_HEIGHT}に変更してください')

  # アラド戦記のウィンドウをデスクトップの左上に移動する
  win32gui.MoveWindow(windowHandle, 0, 0, DNF_WINDOW_WIDTH, DNF_WINDOW_HEIGHT, True)

  # アラド戦記のウィンドウをアクティブにする
  win32gui.SetForegroundWindow(windowHandle)

def click(x: int, y: int, button: str) -> None:
  pyautogui.mouseDown(x=x, y=y, button=button)
  pyautogui.mouseUp(x=x, y=y, button=button)

def drag(x1: int, y1: int, x2: int, y2: int, button: str) -> None:
  pyautogui.mouseDown(x=x1, y=y1, button=button)
  pyautogui.mouseUp(x=x2, y=y2, button=button)

def HIWORD(x: int) -> int:
  return x & 0x8000

def isKeyPressed(key: int) -> bool:
  return bool(HIWORD(ctypes.windll.user32.GetAsyncKeyState(key)))

def loop() -> None:
  # グローブを購入する（左クリックする）
  click(630, 340, 'left')
  # 購入確認ダイアログの「確認」を左クリックする
  click(630, 340, 'left')
  # 装備タブの左上のアイテムを売却する
  drag(850, 440, 750, 440, 'left')
  # 購入確認ダイアログの「確認」を左クリックする
  click(750, 440, 'left')

if __name__ == "__main__":
  initializeDnfWindow()

  while True:
    if isKeyPressed(KEY_CODE_ESCAPE):
      exit()
    else:
      loop()
