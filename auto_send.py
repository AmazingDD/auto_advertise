import os
import win32gui #pywin32-221.win-amd64-py3.7.exe
import win32con
from ctypes import windll
import win32clipboard as w
import time
import datetime
from PIL import Image

#发送文字
def setText(info):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, info)
    w.CloseClipboard()

#发送图片
def setImage(imgpath):
    im = Image.open(imgpath)
    im.save('test.bmp')
    aString = windll.user32.LoadImageW(0, r"test.bmp", win32con.IMAGE_BITMAP, 0, 0, win32con.LR_LOADFROMFILE)
 
    print(aString)
    if aString != 0:  ## 由于图片编码问题  图片载入失败的话  aString 就等于0
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardData(win32con.CF_BITMAP, aString)
        w.CloseClipboard()  

#指定窗口（QQ昵称备注）
def sendByUser(uname):
    hwnd = win32gui.FindWindow('TXGuiFoundation', uname)
    #hwnd = win32gui.FindWindow('ChatWnd', uname)
    win32gui.SendMessage(hwnd, 258, 22, 2080193)
    win32gui.SendMessage(hwnd, 770, 0, 0)
    win32gui.SendMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32gui.SendMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

#获取无后缀的图片名称
def getNosuffixImgName(imgname):
    return os.path.splitext(imgname)[0]

# wrap
def send_img(qun_name):
    if os.path.exists('test.bmp'):
        os.remove('test.bmp')
    setImage(img_dir + img)
    sendByUser(qun_name)
    time.sleep(10)
    

# 特定时间段发送
lower_bound = datetime.datetime.strptime(str(datetime.datetime.now().date()) + '12:00', '%Y-%m-%d%H:%M')
upper_bound =  datetime.datetime.strptime(str(datetime.datetime.now().date()) + '17:00', '%Y-%m-%d%H:%M')

while 1:
    # 当前时间
    now_time = datetime.datetime.now()

    if now_time > lower_bound and now_time < upper_bound:
        # setText('低手续费接代卖，助您解放双手轻松售卖')
        # sendByUser('林婉儿')
        time.sleep(60)
        print('开始发送...')
        # 装备代卖
        img_dir = './imgs/equip/'
        imgs = os.listdir(img_dir)
        for img in imgs:
            send_img('新枫之谷墨雪交易群')
            send_img('新枫之谷橘子二群')
            send_img('新枫之谷交易二群')
            send_img('【金币】新枫之谷常威群')
        time.sleep(600)
        # 点装代卖
        img_dir = './imgs/gash/'
        imgs = os.listdir(img_dir)
        for img in imgs:
            send_img('新枫之谷墨雪交易群')
            send_img('新枫之谷橘子二群')
            send_img('新枫之谷交易二群')
            send_img('新枫之谷山猫时装外观交易群')
        time.sleep(600)
    else:
        break
