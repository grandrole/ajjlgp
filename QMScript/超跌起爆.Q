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
MacroID=3a37f379-eb7a-4983-908b-480d42a3fc18
Description=超跌起爆
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
LogStart "d:\gupiao\log\超跌起爆.log"

Call Lib.公用函数.删除文件("d:\gupiao\data\超跌起爆.json")

Hwnd1 = Lib.公用函数.getHwnd("指南针全赢博弈")
If Hwnd1 > 0 Then 
	Call Lib.框架函数1.loopProcessStockFile("d:\gupiao\data\basic\正常股票.txt","超跌起爆",0)
End If
Delay 1000
If Plugin.File.IsFileExist("d:\gupiao\data\超跌起爆.json") Then 
	Call Lib.框架函数1.loopProcessTHSStockFile("d:\gupiao\data\超跌起爆.json","ENE下轨",0)
End If

LogStop
