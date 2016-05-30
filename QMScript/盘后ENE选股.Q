[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=2
RunOnce=1
EnableWindow=
MacroID=19395024-6cb5-421d-8553-23d2df367edf
Description=盘后ENE选股
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
LogStart "d:\gupiao\log\ENE下轨.log"
num = Plugin.File.ReadINI("Ths", "loopNum", "d:\gupiao\data\Config.ini")
no = CInt(num)
If no = 0 Then 
	If Plugin.File.IsFileExist("d:\gupiao\data\ENE下轨.json") Then 
		Call Lib.公用函数.删除文件("d:\gupiao\data\ENE下轨.json")
	End If
End If

Call Lib.框架函数1.loopProcessTHSStockFile("d:\gupiao\data\huiyan.json","到达ENE下轨", no)

LogStop
