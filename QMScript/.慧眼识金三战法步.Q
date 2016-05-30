[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=b4c0ae4c-3387-4be8-afb8-b908711fb214
Description=慧眼识金三步战法
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
LogStart "d:\gupiao\log\慧眼识金.log"
/*
Hwnd = Lib.同花顺.getThsHwnd()
Call Lib.公用函数.InputKeyboardCode("300310")
Delay 1000
返回值 = Lib.框架函数1.processStockFunction("300310","换手率增长",0)
MessageBox 返回值
*/

Hwnd = Lib.指南针全赢.getHmjHwnd()
If Hwnd < 0 Then 
	Plugin.Msg.Tips "指南针软件启动失败！"
	EndScript
End If
status = Lib.指南针全赢.setTradeDataStyle()
If status = false Then 
	Plugin.Msg.Tips "不能设置分笔数据！"
	EndScript 
End If
//Call Lib.公用函数.删除文件("d:\gupiao\data\huiyan.json")
//Call Lib.框架函数1.loopProcessStockFile("d:\gupiao\data\basic\正常股票.txt", "慧眼识金", 0)
If Plugin.File.IsFileExist("d:\gupiao\data\huiyan.json") Then 
	Call Lib.公用函数.删除文件("d:\gupiao\data\hsl.json")
	Call Lib.框架函数1.loopProcessTHSStockFile("d:\gupiao\data\huiyan.json","换手率增长",0)
End If

stopStatus = Lib.指南针全赢.stopHmj()
If stopStatus < 0 Then 
	TracePrint "关闭红马甲软件失败"
End If

EndScript 
LogStop
