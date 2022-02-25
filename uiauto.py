# -*- coding: utf-8 -*-
import subprocess
import uiautomation as auto
import time


def test():
    print(auto.GetRootControl())
    subprocess.Popen('C:\\Program Files\\Google\\Chrome\\Application\\Chrome.exe')
    googleWindow = auto.PaneControl(searchDepth=1, Name='新标签页 - Google Chrome')
    print(googleWindow.Name)
    googleWindow.SetTopmost(True)
    # 查找notepadWindow所有子孙控件中的第一个EditControl，因为EditControl是第一个子控件，可以不指定深度
    search = googleWindow.ComboBoxControl(searchDepth=5, Name='在 Google 上搜索，或者输入一个网址')
    print(search.Name)
    try:
        # 获取EditControl支持的ValuePattern，并用Pattern设置控件文本为"Hello"
        search.GetValuePattern().SetValue('https://zh.moegirl.org.cn/index.php?title=%E5%88%BB%E6%99%B4&variant=zh-hans'
                                          '&mobileaction=toggle_view_desktop')  # or edit.GetPattern(
    except auto.comtypes.COMError as ex:
        # 如果遇到COMError, 一般是没有以管理员权限运行Python, 或者这个控件没有实现pattern的方法(如果是这种情况，基本没有解决方法)
        # 大多数情况不需要捕捉COMError，如果遇到了就加到try block
        pass
    auto.SendKeys('{Enter}')  # 在文本末尾打字
    print('current text:', search.GetValuePattern().Value)  # 获取当前文本
    subprocess.Popen('C:\\Program Files (x86)\\DingDing\\DingtalkLauncher.exe')
    dingWindow = auto.WindowControl(searchDepth=1, Name='钉钉')
    dingWindow.SetTopmost()
    edit = dingWindow.EditControl(searchDepth=7, Name='请输入消息')
    times = 2
    for i in range(0, times):
        try:
            edit.GetValuePattern().SetValue('喂')
        except auto.comtypes.COMError as ex:
            edit.Click()
            edit.SendKeys("喂？")
            pass
        #dingWindow.ButtonControl(searchDepth=9, Name='发送').Click()
        time.sleep(0.1)
    dingWindow.ButtonControl(searchDepth=3, Name='关闭').Click()
    print(1)

    # 然后从TitleBarControl的子孙控件中找第二个ButtonControl, 即最大化按钮，并点击按钮
    # notepadWindow.TitleBarControl(Depth=1).ButtonControl(foundIndex=2).Click()
    # 从notepadWindow前两层子孙控件中查找Name为'关闭'的按钮并点击按钮
    # notepadWindow.ButtonControl(searchDepth=2, Name='关闭').Click()


if __name__ == '__main__':
    test()
