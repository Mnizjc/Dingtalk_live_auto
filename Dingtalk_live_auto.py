# -*- coding:utf-8 -*-
# 作者：伊斯卡尼达尔
# 注释：mnizjc
# 创建：2022-10-31
# 更新：2022-10-31
# 注释及fork：2024-7-3
# 用意：钉钉直播签到自动化程序

"""
我提供你们的是python源码, 运行程序前请下载相关的库, 
签到部分是通过点击坐标实现的, 每个人的电脑屏幕不同, 所以你得改一下坐标！
请尊重成果, 有疑问可以QQ：1457436639联系我.
"""

import os
import time
from typing import Any, List, NewType, Optional

import pyautogui
import win32api
import win32con
import win32gui
from PIL import ImageGrab

import pywintypes

# 句柄类型,用于类型标注,不全大写因为会与ctypes.wintypes.HWND冲突而后者是(void*)
Hwnd = NewType('Hwnd', int)

hwnd_title = {}  # 窗口标题字典


def get_all_hwnd(hwnd: Hwnd, mouse: Any):
    """该函数为传入EnumWindows的回调函数,其回调函数规范为

    BOOL CALLBACK EnumWindowsProc(
      _In_ HWND   hwnd,
      _In_ LPARAM lParam
    );

    第一个参数是窗口句柄,第二个参数是一个由主调函数传来的自定义指针,本函数未使用

    本函数作用为将主调函数给定的窗口句柄及句柄对应窗口标题保存在hwnd_title

    Args:
        hwnd (Hwnd): 窗口句柄
        mouse (Any): 自定义指针
    """
    if (win32gui.IsWindow(hwnd) and  # 句柄是否指向窗口,句柄在python win32gui中的类型是int
            win32gui.IsWindowEnabled(hwnd) and  # 窗口是否启用(?)
            win32gui.IsWindowVisible(hwnd)):  # 是否WS_VISIBLE,文档中说若某个窗口被其他窗口完全遮挡也有可能不为0
        # (涉及字符串所以分AW版本)获取窗口的标题或者控件的文本
        hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})


'''
win32gui对GetWindowText函数进行了极大程度上的简化,该函数原型为
int GetWindowTextA( //返回不包含中止null的以复制的字符串长度,0则失败,使用GetLastError函数查看错误信息
  [in]  HWND  hWnd, //句柄
  [out] LPSTR lpString, //缓冲区
  [in]  int   nMaxCount //缓冲区大小
);
而这里只需要给定一个句柄即可,且直接返回本应放在缓冲区的内容
不过这也有个好处,在python一切内容都是动态的,不需要为数组或者容器或者缓冲区之类的东西的大小考虑了.
'''


def get_all_child_window(parent: Hwnd) -> Optional[List[Hwnd]]:
    """获取parent句柄下的子窗口

    Args:
        parent (Hwnd): 指定父窗口

    Returns:
        Optional[List[Hwnd]]: 子窗口列表,可能会返回None
    """
    if not parent:  # 如果parent句柄null
        return  # 则返回null(或者这里我应该叫None)
    hwndChildList = []  # 子窗口列表
    win32gui.EnumChildWindows(  # 枚举parent窗口的子窗口直到最后一个字窗口或回调函数返回FALSE(此FALSE非彼False)
        parent, lambda hwnd, param: param.append(hwnd), hwndChildList)
    return hwndChildList


'''
经典的回调+自定义指针, EnumChildWindows原型为
BOOL EnumChildWindows( //不使用返回值
  [in, optional] HWND        hWndParent, //指定父窗口
  [in]           WNDENUMPROC lpEnumFunc, //回调函数
  [in]           LPARAM      lParam //传给回调函数的自定义指针
);
'''


def setforeground_window(window_handle: Hwnd):
    """将窗口至于最顶层

    Args:
        window_handle (Hwnd): 应激活并带到前台的窗口的句柄。(原文如此)
    """
    while True:
        try:
            win32gui.SetForegroundWindow(window_handle)  # 强制在最前端显示
            # Iskender()
            return
        except:
            time.sleep(0.1)


'''
SetForegroundWindow官方文档称其作用为:
将创建指定窗口的线程引入前台并激活窗口。
键盘输入将定向到窗口，并为用户更改各种视觉提示。
系统为创建前台窗口的线程分配的优先级略高于其他线程。

至于这玩意会引发啥异常,咱不是自己开发的,只是拜读大佬的代码,咱也不知道可能发生什么
'''


def close_analyse_window():
    """尝试关闭统计窗口
    """
    print("检测到直播结束...")
    # 关闭统计窗口
    print("尝试关闭统计窗口...")
    try:
        win32gui.EnumWindows(get_all_hwnd, 0)
        for h, t in hwnd_title.items():
            if win32gui.GetClassName(h) == "DingEAppWnd":
                setforeground_window(h)  # 使当前窗口在最前
                win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                print("成功关闭统计窗口")
                break
    except:
        print("关闭失败, 可能统计窗口已经被关闭了...")
    print("本轮检测结束, " + str(delay_time) + "s后进行下一轮检测")
    # return


'''
EnumWindows官方文档:
通过将句柄传递到每个窗口，进而将传递给应用程序定义的回调函数，枚举屏幕上的所有顶级窗口。
枚举窗口 将一直持续到最后一个顶级窗口被枚举或回调函数返回 FALSE。

原型为:
BOOL EnumWindows( //成功则非0,失败则0,使用GetLastError函数查看错误信息
  [in] WNDENUMPROC lpEnumFunc, //回调函数
  [in] LPARAM      lParam //自定义指针
);

win32gui应该也对这种行为做了简化,
这里EnumWindows实际返回的是None,
错误信息用异常捕获了

GetClassName官方文档:
检索指定窗口所属的类的名称。

原型为:
int GetClassName( //返回不包含中止null的以复制的字符串长度,0则失败,使用GetLastError函数查看错误信息
  [in]  HWND   hWnd, //窗口句柄
  [out] LPTSTR lpClassName, //缓冲区
  [in]  int    nMaxCount //缓冲区大小
);

简化,同样经典用缓冲区传出字符串

PostMessage官方文档:
Places (在与创建指定窗口的线程关联的消息队列中发布) 消息，
并在不等待线程处理消息的情况下返回消息。
(为什么这里翻译都不全?这里places应该是「放置」义)

原型:
BOOL PostMessageA( 成功则非0,失败则0
  [in, optional] HWND   hWnd, //窗口句柄,值(HWND)0xffff则广播,(实际还有一个特殊值null,但我没看懂干啥的,针对win32的开发经验还是太少了)
  [in]           UINT   Msg, //要传递的消息,参见文档
  [in]           WPARAM wParam, //其他的消息特定信息
  [in]           LPARAM lParam //其他的消息特定信息
);

这个函数行为和SendMessage函数作用很类似,都是发送消息,
而这二者的区别是
PostMessage函数将消息压入消息队列就返回,程序继续运行(换句话说就是异步)
SendMessage函数直接调用指定窗口的过程,并等待完成之后返回消息处理结果(即同步)

SendMessage的原型除hWnd未标记optional外其余与PostMessage一致,这里不再赘述

并且在最后一句加入了一句用处并不大的return,
up主可能经常写一些类c语言如c/c++,
其实get_all_hwnd函数中if语句使用括号的习惯也能看出一点
'''


def get_live_window_isopened(live_window_handle):
    '''
    能看懂这个函数干啥了,但是没看懂为啥要这个参数...
    '''
    while True:
        time.sleep(delay_time)
        # 查找所有窗口标题和句柄 StandardFrame
        isOpened_temp = False
        win32gui.EnumWindows(get_all_hwnd, 0)
        try:
            for h, t in hwnd_title.items():
                if t == '钉钉' and win32gui.GetClassName(h) == "StandardFrame":
                    isOpened_temp = True
        except:
            close_analyse_window()
            return
        if not isOpened_temp:
            close_analyse_window()
            return


'''
至于怎么获取到这个什么ClassName,
我这两天找到了两个能查出来窗口的东西
SendMessage(这次不是函数)和Window Detective
'''

delay_time = 60


def Iskender(live_wnd_handle: Hwnd):
    """最大化直播窗口并准备点签到

    Args:
        live_wnd_handle (Hwnd): 直播窗口句柄
    """
    var = 1
    s = 0
    # 最大化直播窗口,这样窗口就固定下来,否则每次打开窗口都会乱跑,固定位置点不上
    win32gui.ShowWindow(live_wnd_handle, win32con.SW_MAXIMIZE)
    while var == 1:  # 表达式永远为 true
        time.sleep(5)
        s = s+1
        # pyautogui.click(x=1082, y=615)  # 单击
        # 最大化后:1812,1171
        pyautogui.click(x=1812, y=1171)
        print("点击", s, "次")


# 查找所有窗口标题和句柄 StandardFrame_DingTalk
win32gui.EnumWindows(get_all_hwnd, 0)
for h, t in hwnd_title.items():
    if t != '钉钉':
        continue
    if win32gui.GetClassName(h) != "StandardFrame_DingTalk":  # 跳过其他窗口捕获主窗口
        continue
    print("获取到钉钉窗口句柄 " + str(h))
    ding_main_window_handle = h
    break

try:
    if ding_main_window_handle is None:
        exit(0)
except:
    print("请先打开钉钉窗口")
    os.system("pause")
    exit(0)

ding_child_list = get_all_child_window(ding_main_window_handle)

print(ding_child_list)

win32gui.ShowWindow(ding_main_window_handle, win32con.SW_MAXIMIZE)  # 最大化
setforeground_window(ding_main_window_handle)  # 强制在最前端显示
time.sleep(2)

for c in ding_child_list:
    if win32gui.GetWindowText(c) == "Chrome Legacy Window":
        ding_chrome_window = c
try:
    # 可能刚打开钉钉,还没打开群组聊天,此时找不到Chrome Legacy Window
    print("获取到聊天窗口 " + str(ding_chrome_window))
except NameError:
    print("未获取到聊天窗口,请确定所在群组并打开聊天窗口")
    win32gui.ShowWindow(ding_main_window_handle, win32con.SW_MINIMIZE)
    exit(0)
win32gui.ShowWindow(ding_main_window_handle, win32con.SW_MINIMIZE)  # 完成后最小化

print("准备就绪, 3s后开始检测")

while True:
    win32gui.ShowWindow(ding_main_window_handle, win32con.SW_MAXIMIZE)  # 最大化

    time.sleep(0.3)
    try:
        x_start, y_start, x_end, y_end = win32gui.GetWindowRect(
            ding_chrome_window)
        box = (x_start, y_start, x_end, y_end)
        image = ImageGrab.grab(box)
        win32gui.ShowWindow(ding_main_window_handle,
                            win32con.SW_MINIMIZE)  # 截图完成后最小化
        if image.getpixel((5, 5)) == (224, 237, 254):
            live_window_handle_temp = ''
            print("检测到有直播可进入, 准备检测是否已启动直播页面...")
            # 查找所有窗口标题和句柄 StandardFrame
            win32gui.EnumWindows(get_all_hwnd, 0)
            isOpened = False
            for h, t in hwnd_title.items():
                if t != '钉钉':
                    continue
                if win32gui.GetClassName(h) != "StandardFrame":  # 跳过其他窗口捕获直播窗口
                    continue
                isOpened = True
                live_window_handle_temp = h
                break
            if isOpened:
                print("直播窗口已打开, 开始监控直播窗口变化")
                Iskender(live_window_handle_temp)
                get_live_window_isopened(live_window_handle_temp)
            else:
                print("直播窗口未打开, 尝试打开")
                win32gui.ShowWindow(ding_main_window_handle,
                                    win32con.SW_MAXIMIZE)  # 最大化
                setforeground_window(ding_main_window_handle)  # 强制在最前端显示
                time.sleep(0.5)
                left, top, right, bottom = win32gui.GetWindowRect(
                    ding_chrome_window)
                move_x = left + 5
                move_y = top + 5
                win32api.SetCursorPos((move_x, move_y))  # 鼠标挪到点击处
                win32api.mouse_event(
                    win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)  # 鼠标左键按下
                win32api.mouse_event(
                    win32con.MOUSEEVENTF_LEFTUP, 0, 0)  # 鼠标左键抬起
                print("启动完成, 等待直播进入...")
                time.sleep(8)
                print("开始获取直播窗口...")
                # 查找所有窗口标题和句柄 StandardFrame
                win32gui.EnumWindows(get_all_hwnd, 0)
                isOpened = False
                for h, t in hwnd_title.items():
                    if t != '钉钉':
                        continue
                    if win32gui.GetClassName(h) != "StandardFrame":  # 跳过其他窗口捕获直播窗口
                        continue
                    isOpened = True
                    live_window_handle_temp = h
                    break
                if isOpened:
                    print("直播窗口已打开, 开始监控直播窗口变化")
                    Iskender(live_window_handle_temp)
                    get_live_window_isopened(live_window_handle_temp)

                else:
                    print("打开失败, " + str(delay_time) + "s后再次尝试...")
        else:
            print("未检测到直播, 本轮检测完毕")
        image.save('temp.jpg')
        time.sleep(delay_time)
    except pywintypes.error as e:
        print("本轮尝试执行过程中发现如下异常：%s\t可能是因为钉钉进程已被销毁, 请检查重试, 退出中..." % str(e))
        os.system("pause")
