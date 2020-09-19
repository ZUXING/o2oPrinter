import os
import win32api
import win32event
import win32process
import win32con
import win32print

#Version 0.77
#ChenGuanglin Software Studio & ZUXING

# 预定义内容 ================================================================================================
projectName = 'o2oPrinter'
hostCurrUser = os.getenv('USERNAME') # 当前用户名
defaultDir = 'C:\\Users\\' + hostCurrUser + '\\' # 默认文件夹
iniPath = defaultDir + projectName + '.ini' # ini文件路径

# 读取ini文件 ===============================================================================================
import configparser
config = configparser.ConfigParser()

ini = os.path.exists(iniPath) # 定义ini文件位置
if not (ini): # ini文件不存在
        # 第一次运行时要设置的内容 ==========================================================================
        import tkinter # 导入tkinter库
        from tkinter import filedialog

        # 指定文件接收的文件夹
        setWindow = tkinter.Tk() # 创建一个Tkinter.Tk()实例
        setWindow.withdraw() # 将Tkinter.Tk()实例隐藏
        

        # 手动指定文件夹
        recvDir = filedialog.askdirectory(title=u'接收文件的文件夹', initialdir=(os.path.expanduser(defaultDir)))
        if len(recvDir) == 0:
                recvDir = defaultDir
        print(recvDir)

        # 保存为ini文件 =====================================================================================

        config.add_section('path') # ini节
        config.set('path','recvDir',recvDir)  # 写入接收文件路径
        
        config.add_section('fileProcess') # ini节
        config.set('fileProcess','delRecvFile','0') # 写入接收后是否删除
        
        config.write(open(iniPath,'a'))            #保存数据
        
else: # ini文件存在
        config.read(iniPath)
        recvDir = config.get('path', 'recvDir')
        print(recvDir)
        
# 开始核心程序 ==============================================================================================

print("在线打印机开始工作")

from win32com.shell.shell import ShellExecuteEx
from win32com.shell import shellcon


while 1:
	dirs = os.listdir(recvDir)
	for file in dirs:
		if os.path.splitext(file)[1] == ".docx":
			print(file)
			process_info = ShellExecuteEx(nShow=win32con.SW_SHOW,
				fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
				lpVerb='print',
				lpFile=recvDir + '\\' + file,
				lpParameters='/d:"%s"' % win32print.GetDefaultPrinter ())
			win32event.WaitForSingleObject(process_info['hProcess'], -1)
			print("打印完成")
			os.remove(recvDir + file)
os.system("pause")
