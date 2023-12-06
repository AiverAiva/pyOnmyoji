import win32gui, win32con, win32api, random, time, os, win32ui, cv2
import numpy as np

DEBUG = False

class HwndError(Exception):
    pass

class Robot():

    def __init__(self):
        self.__hwnd = self.__get_window()
        self.thread_ter = False
        print(self.__hwnd)
        
    def click(self, x, y, sleepTime = 0):
        lParam = win32api.MAKELONG(x, y)
        win32gui.PostMessage(self.__hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lParam) #btndown
        time.sleep(random.uniform(0.1,0.5))
        win32gui.PostMessage(self.__hwnd, win32con.WM_LBUTTONUP, 0, lParam)                     #btnup
        time.sleep(sleepTime)
        
        # print("點",x,y,"坐標")
        # long_position = win32api.MAKELONG(x, y)  # 模拟鼠标指针 传送到指定坐标
        # win32api.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position)  # 模拟鼠标按下
        # win32api.SendMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position)  # 模拟鼠标弹起
    
    def drag(self, pos1, pos2):
        """
        後台鼠標拖拽
            :param pos1: (x,y) 起點坐標
            :param pos2: (x,y) 終點坐標
        """
        if 1:
            move_x = np.linspace(pos1[0], pos2[0], num=50, endpoint=True)[0:]
            move_y = np.linspace(pos1[1], pos2[1], num=50, endpoint=True)[0:]
            win32gui.SendMessage(self.__hwnd, win32con.WM_LBUTTONDOWN, 0, win32api.MAKELONG(pos1[0], pos1[1]))
            for i in range(50):
                x = int(round(move_x[i]+random.uniform(1,3)))
                y = int(round(move_y[i]+random.uniform(1,2)))
                win32gui.SendMessage(self.__hwnd, win32con.WM_MOUSEMOVE, 0, win32api.MAKELONG(x, y))
                time.sleep(0.01)
            win32gui.SendMessage(self.__hwnd, win32con.WM_LBUTTONUP, 0, win32api.MAKELONG(pos2[0], pos2[1]))

    
    
    
    # def get_color(self, x1, y1):
        ##前台
        # i_desktop_window_id = win32gui.GetDesktopWindow()
        # i_desktop_window_dc = win32gui.GetWindowDC(i_desktop_window_id)
        # x2, y2 = win32gui.ClientToScreen(self.__hwnd, (x1, y1))
        # long_colour = win32gui.GetPixel(i_desktop_window_dc, x2, y2)
        # i_colour = int(long_colour)
        # R, G, B = (i_colour & 0xff), ((i_colour >> 8) & 0xff), ((i_colour >> 16) & 0xff)
        # return (R, G, B)
    def get_color(self, x1, y1):
        #後台
        try:
            i_desktop_window_dc = win32gui.GetWindowDC(self.__hwnd)
            long_colour = win32gui.GetPixel(i_desktop_window_dc, x1, y1)
            i_colour = int(long_colour)
            R, G, B = (i_colour & 0xff), ((i_colour >> 8) & 0xff), ((i_colour >> 16) & 0xff)
            # print (R, G, B)
            return (R, G, B)
        except:
            raise HwndError("HwndError")
    
    def find_color(self, x1, y1, x2, y2, R, G, B, R2, G2, B2):
        i_desktop_window_dc = win32gui.GetWindowDC(self.__hwnd)
        for j in range(y1, y2, 5):
            for i in range(x1, x2, 5):
                long_colour = win32gui.GetPixel(i_desktop_window_dc, i, j)
                # long_colour2 = win32gui.GetPixel(i_desktop_window_dc, i+1, j+1)
                i_colour = int(long_colour)
                # i_colour2 = int(long_colour)
                # print(i, j)
                if( abs(R -(i_colour & 0xff)) < 20 and  abs(G - ((i_colour >> 8) & 0xff)) < 20 and abs(B - ((i_colour >> 16) & 0xff)) < 20):
                    for l in range(j, j+10):
                        for k in range(i, i+10):
                            long_colour2 = win32gui.GetPixel(i_desktop_window_dc, k, l)
                            i_colour2 = int(long_colour2)
                            if(abs(R2 -(i_colour2 & 0xff)) < 10 and  abs(G2 - ((i_colour2 >> 8) & 0xff)) < 10 and abs(B2 - ((i_colour2 >> 16) & 0xff)) < 10):
                                print (i, j)
                                print ((i_colour & 0xff), ((i_colour >> 8) & 0xff), (i_colour >> 16) & 0xff)
                                return i, j
        return -1, -1
        
    
    def find_color_thread(self, x1, y1, x2, y2, R, G, B, R2, G2, B2):
        i_desktop_window_dc = win32gui.GetWindowDC(self.__hwnd)
        for j in range(y1, y2, 5):
            for i in range(x1, x2, 5):
                if(self.thread_ter):
                    return -1, -1
                long_colour = win32gui.GetPixel(i_desktop_window_dc, i, j)
                # long_colour2 = win32gui.GetPixel(i_desktop_window_dc, i+1, j+1)
                i_colour = int(long_colour)
                # i_colour2 = int(long_colour)
                # print(i, j)
                if( abs(R -(i_colour & 0xff)) < 20 and  abs(G - ((i_colour >> 8) & 0xff)) < 20 and abs(B - ((i_colour >> 16) & 0xff)) < 20):
                    for l in range(j, j+10):
                        for k in range(i, i+10):
                            long_colour2 = win32gui.GetPixel(i_desktop_window_dc, k, l)
                            i_colour2 = int(long_colour2)
                            if(abs(R2 -(i_colour2 & 0xff)) < 10 and  abs(G2 - ((i_colour2 >> 8) & 0xff)) < 10 and abs(B2 - ((i_colour2 >> 16) & 0xff)) < 10):
                                print (i, j)
                                print ((i_colour & 0xff), ((i_colour >> 8) & 0xff), (i_colour >> 16) & 0xff)
                                self.thread_ter = True
                                return i, j
        return -1, -1
    
    def get_pos(self):
        
        #找座標
        x, y = win32api.GetCursorPos()
        x1, y1 = win32gui.ScreenToClient(self.__hwnd ,(x, y))
        
        #找色
        i_desktop_window_id = win32gui.GetDesktopWindow()
        i_desktop_window_dc = win32gui.GetWindowDC(i_desktop_window_id)
        x2, y2 = win32gui.ClientToScreen(self.__hwnd, (x1, y1))
        long_colour = win32gui.GetPixel(i_desktop_window_dc, x2, y2)
        i_colour = int(long_colour)
        R, G, B = (i_colour & 0xff), ((i_colour >> 8) & 0xff), ((i_colour >> 16) & 0xff)
        print (str(x1) + ", " + str(y1) + ", " + str(R) + ", " + str(G) +", "+ str(B))
        return (x1, y1, R, G, B)
    def window_capture(self, filename):
        # 根據窗口句柄獲取窗口的設備上下文DC（Divice Context）
        hwndDC = win32gui.GetWindowDC(self.__hwnd)
        # 根據窗口的DC獲取mfcDC
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        # mfcDC創建可兼容的DC
        saveDC = mfcDC.CreateCompatibleDC()
        # 創建bigmap準備保存圖片
        saveBitMap = win32ui.CreateBitmap()
        # 獲取監控器信息
        MoniterDev = win32api.EnumDisplayMonitors(None, None)
        w = MoniterDev[0][2][2]
        h = MoniterDev[0][2][3]
        # print w,h　　　#圖片大小
        # 為bitmap開辟空間
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        # 高度saveDC，將截圖保存到saveBitmap中
        saveDC.SelectObject(saveBitMap)
        # 截取從左上角（0，0）長寬為（w，h）的圖片
        saveDC.BitBlt((0, 0), (w, h), mfcDC, (0, 0), win32con.SRCCOPY)
        saveBitMap.SaveBitmapFile(saveDC, filename)
    
    def window_capture_region(self, x0, y0, x1, y1, filename):
        # 根據窗口句柄獲取窗口的設備上下文DC（Divice Context）
        hwndDC = win32gui.GetWindowDC(self.__hwnd)
        # 根據窗口的DC獲取mfcDC
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        # mfcDC創建可兼容的DC
        saveDC = mfcDC.CreateCompatibleDC()
        # 創建bigmap準備保存圖片
        saveBitMap = win32ui.CreateBitmap()
        # 獲取監控器信息
        MoniterDev = win32api.EnumDisplayMonitors(None, None)
        # w = MoniterDev[0][2][2]
        # h = MoniterDev[0][2][3]
        # print w,h　　　#圖片大小
        # 為bitmap開辟空間
        saveBitMap.CreateCompatibleBitmap(mfcDC, x1-x0, y1-y0)
        # 高度saveDC，將截圖保存到saveBitmap中
        saveDC.SelectObject(saveBitMap)
        # 截取從左上角（0，0）長寬為（w，h）的圖片
        saveDC.BitBlt((0, 0), (x1-x0, y1-y0), mfcDC, (x0, y0), win32con.SRCCOPY)
        saveBitMap.SaveBitmapFile(saveDC, "./PIC/" + filename)
    
    def imagesearch(self, image, precision=0.8):
        # im = cv2.imread("./PIC/!current.jpg")
        im = cv2.imread("./PIC/!current.jpg", cv2.IMREAD_GRAYSCALE)
        
        # template = cv2.imread("./PIC/" + image + ".jpg")
        template = cv2.imread("./PIC/" + image + ".jpg", cv2.IMREAD_GRAYSCALE)
        
        #獲取圖片大小 width, height, channel
        w, h, c = cv2.imread("./PIC/" + image + ".jpg").shape
        
        res = cv2.matchTemplate(im, template, cv2.TM_CCOEFF_NORMED)
        
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        # print(max_val, max_loc)
        if max_val < precision:
            return [-1,-1, -1, -1]
        return max_loc[0], max_loc[1], w, h
    
    def __get_window(self):  
    
        #父視窗
        parent = win32gui.FindWindow(None, "Task Manager")
        #print (parent)
        return parent
        
        #找尋子視窗
        #if not parent:
        #    return 0
        #hwndChildList = []     
        #win32gui.EnumChildWindows(parent, lambda hwnd, param: param.append(hwnd),  hwndChildList)
        
        #for i in hwndChildList:
        #    print (win32gui.GetWindowText(i))
        #    if(win32gui.GetWindowText(i) == "Onmyoji"):
        #        return i
                
        #return 0



if __name__ == '__main__':
    a = Robot()
    a.window_capture("test.jpg")

