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
Description=������
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
LogStart "d:\gupiao\log\������.log"

Call Lib.���ú���.ɾ���ļ�("d:\gupiao\data\������.json")

Hwnd1 = Lib.���ú���.getHwnd("ָ����ȫӮ����")
If Hwnd1 > 0 Then 
	Call Lib.��ܺ���1.loopProcessStockFile("d:\gupiao\data\basic\������Ʊ.txt","������",0)
End If
Delay 1000
If Plugin.File.IsFileExist("d:\gupiao\data\������.json") Then 
	Call Lib.��ܺ���1.loopProcessTHSStockFile("d:\gupiao\data\������.json","ENE�¹�",0)
End If

LogStop
