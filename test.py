import win32gui
import win32ui
from ctypes import windll
from PIL import Image

hwnd = win32gui.FindWindow(None, 'Calculator')

#left, top, right, bot = win32gui.GetClientRect(hwnd)

windll.user32.SetProcessDPIAware()
left, top, right, bot = win32gui.GetWindowRect(hwnd)
print(left, top, right, bot)
w = right - left
h = bot - top

hwndDC = win32gui.GetWindowDC(hwnd)
mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
saveDC = mfcDC.CreateCompatibleDC()

saveBitMap = win32ui.CreateBitmap()
saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

saveDC.SelectObject(saveBitMap)

#result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 1)

#the numbers behind matters
result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)
print(result)

bmpinfo = saveBitMap.GetInfo()
bmpstr = saveBitMap.GetBitmapBits(True)

im = Image.frombuffer(
    'RGB',
    (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
    bmpstr, 'raw', 'BGRX', 0, 1)

win32gui.DeleteObject(saveBitMap.GetHandle())
saveDC.DeleteDC()
mfcDC.DeleteDC()
win32gui.ReleaseDC(hwnd, hwndDC)

if result == 1:
    im.save("test.png")