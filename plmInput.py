from pywinauto import Application
from pywinauto import findwindows
from pywinauto import mouse
from pywinauto import uia_defines as uia

import pyautogui
import time
import psutil
from pywinauto.findbestmatch import MatchError
import configparser
import sys

def is_app_running(app_name):
    # 遍历所有正在运行的进程
    for proc in psutil.process_iter(['pid', 'name', 'status']):
        try:
            # 获取进程信息
            pinfo = proc.as_dict(attrs=['pid', 'name', 'status'])
            # 判断进程名是否匹配应用程序名，并且进程状态为“running”或“sleeping”
            if app_name in pinfo['name'].lower() and pinfo['status'] in ['running', 'sleeping']:
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
#        print(pinfo['name'])
    return False

if is_app_running('caxavault.exe'):
    print('应用打开')
    app = Application(backend='uia').connect(path="C:\\Program Files\\Common Files\\CAXA shared\\CAXA EAP\\1.0\\bin\\CaxaVault.exe")
else:
    print("应用未打开")
    # 启动目标应用
    try:
        app = Application().start('C:\\Program Files\\Common Files\\CAXA shared\\CAXA EAP\\1.0\\bin\\CaxaVault.exe')
    except :
        pyautogui.alert("程序的路径不正确")



# 获取应用程序主窗口
try:
    #main_window = app.window(title_re="CAXA PLM 协同管理2021 - 图文档 - 查询")
    main_window = app.window(title_re='CAXA PLM 协同管理.*', visible_only=True)
    # 获取输入框控件
except:
    pyautogui.alert("plm找不到，且找不到输出框")
    sys.exit()

# 将输入框设置为焦点
#print(main_window.print_control_identifiers())
gongshi = main_window.child_window(title="工时汇报", control_type="Hyperlink")
# 等待控件加载完成并且处于可用状态
gongshi.wait('exists enabled', timeout=10)
gongshi.set_focus()
# 执行超链接所代表的操作
gongshi.click_input()

# 获取应用程序中的所有窗口对象

gWindow = main_window.child_window(title='工时汇报',control_type='Window')
#gWindow.print_control_identifiers()
#project = gWindow.child_window(title="小时", control_type="Edit")
project = gWindow.child_window(auto_id="20155", control_type="ComboBox")
p1 = project.child_window(auto_id="DropDown", control_type="Button")
# 点击ComboBox控件，打开下拉列表
p1.click_input()

# 输入向下箭头键，选择第4个选项
for i in range(4):
    pyautogui.press('down')

# 输入回车键，确定选择
pyautogui.press('enter')

combo_box = gWindow.child_window(auto_id="20156", control_type="ComboBox")
#edit = combo_box.child_window(auto_id="1001", control_type="Edit")
#edit = gWindow.child_window(auto_id="1001", control_type="Edit")
# 等待控件加载完成并且处于可用状态
edit = combo_box.child_window(auto_id="DropDown", control_type="Button")

edit.click_input()

# 输入向下箭头键，选择第4个选项
for i in range(1):
    pyautogui.press('down')

# 输入回车键，确定选择
pyautogui.press('enter')

#日期
date_button = gWindow.child_window( auto_id="69273889", control_type="Button")
date_button.set_focus()
date_button_rect = date_button.rectangle()
pyautogui.click(date_button_rect.right-30, date_button_rect.top+10)
pyautogui.typewrite('17')

content = gWindow.child_window(auto_id="20159", control_type="Edit")
content.set_focus()
# 输入要搜索的项目名称
content.type_keys('富士康的打印机进行打印确认')

#工作时间按
content = gWindow.child_window(auto_id="20160", control_type="Edit")
content.set_focus()
# 输入要搜索的项目名称
content.type_keys('8')

resume = gWindow.child_window(title="提交", auto_id="1", control_type="Button")
resume.set_focus()
resume.click()


time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.press('enter')
#resume = gWindow.child_window(title="取消",  control_type="Button")
#resume.set_focus()
#resume.click()
sys.exit()

#input_box.type_keys("adf")
#pyautogui.keyDown('enter')

