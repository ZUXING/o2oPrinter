import os
import win32api
import win32event
import win32process
import win32con
import win32print

#Version 0.76
#ChenGuanglin Software Studio & ZUXING
print("Codename:o2oPrinter")
print("在线打印机开始工作")

from win32com.shell.shell import ShellExecuteEx
from win32com.shell import shellcon

recvFilePath = 'o2oPrint\\'
while 1:
	dirs = os.listdir(recvFilePath)
	for file in dirs:
		if os.path.splitext(file)[1] == ".docx":
			print(file)
			process_info = ShellExecuteEx(nShow=win32con.SW_SHOW,
				fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
				lpVerb='print',
				lpFile=recvFilePath + file,
				lpParameters='/d:"%s"' % win32print.GetDefaultPrinter ())
			win32event.WaitForSingleObject(process_info['hProcess'], -1)
			print("打印完成")
			os.remove(recvFilePath + file)
os.system("pause")
