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
MacroID=d2945f87-1d6f-4e10-88a8-4ed5aa37bff1
Description=盘中盯盘ENE
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
LogStart "d:\gupiao\log\盘中ENE盯盘.log"

moningStartTime = TimeValue("9:30:00")
moningStopTime = TimeValue("11:30:00")
afternoonStartTime = TimeValue("13:00:00")
afternoonStopTime = TimeValue("15:00:00")

Dim dt

/*
	程序运行状态： 
		0:  ready 启动程序
		1:  start 开始
		2:  Stop 结束
*/


status = 1
Do
	dt = Time
	ret1 = ret2 = ret3 = ret4 = 0
	ret1 = DateDiff("s", dt, moningStartTime)
	Delay 1000
	If ret1 < 0 Then 
	 	//上午已经开市
		ret3 = DateDiff("s", dt, moningStopTime)
		Delay 1000
		If ret3 < 0 Then 
			//上午已经休市
			ret2 = DateDiff("s", dt, afternoonStartTime)
			Delay 1000
			If ret2 < 0 Then 
				//下午已开市
				ret4 = DateDiff("s", dt, afternoonStopTime)
				Delay 1000
				If ret4 < 0 Then 
					//下午已休市
					status = 2
				Else 
					//下午正在交易
					status = 1
				End If
			Else 
				//下午未开市
				status = 0
			End If
		Else 
			//上午正在交易
			status = 1
		End If
	Else 
		//上午未开市
		status = 0
	End If
	
	//在非交易时间测试程序使用	
	status = 1
	Select Case status
		Case 0
			If ret1 > 0 Then 
				Plugin.Msg.Tips "离上午开市时间还差" & ret1 \ 60 & "分钟"
				Delay 60000
			Else
				If ret2 > 0 Then 
					Plugin.Msg.Tips "离下午开市时间还差" & ret2 \ 60 & "分钟"
					Delay 60000
				End If
			End If
		Case 1
			//初始化系统
			currTime = Minute(Now)
			thsHwnd = Lib.同花顺.getThsHwnd()
			If thsHwnd < 0 Then 
				MessageBox "启动同花顺程序失败，请退出系统，重新运行"
				EndScript
			End If

			//盘中盯盘

			//查找持仓股票的卖点
			Call Lib.公用函数.删除文件("d:\gupiao\data\卖出股票.json")
			Call Lib.公用函数.删除文件("d:\gupiao\data\买入股票.json")
			//持仓股票处理
			Call chiCangStock()
			//自选股票找买点
			Call zixuanguStock()
			
			If currTime mod 30 < 3 Then 
				Call getStockData("30分钟")
				//获取30分钟数据
				If ret2 > 0 Then 
					//获取上午60分钟数据
					Call getStockData("60分钟")
				End If
			End If
			
			If currTime mod 60 < 4 Then 
				//获取60分钟数据
				Call getStockData("60分钟")
			End If
			
			//间隔5分钟
			interVal = Minute(Now) mod 5
			If interVal >= 0 Then 
				interVal = 5 - InterVal
				Plugin.Msg.Tips "等待" & interVal & "分钟"
				Delay interVal * 60 * 1000
			End If
			
			
		Case 2
			//获取日线数据
			Call getStockData("day")
			//获取周线数据
			Call getStockData("week")
			//获取月线数据
			Call getStockData("month")
			//查找自选股票
			
			Call Lib.公用函数.删除文件("d:\gupiao\data\删除股票.json")
			Call Lib.公用函数.删除文件("d:\gupiao\data\自选股票.json")
			Call Plugin.Sys.ExitOS(1)
	End Select
Loop

LogStop 
Plugin.Msg.Tips "股市结束，正在关机 ......"
//Call Plugin.Sys.ExitOS(1)
EndScript

Function getStockData(period)
	返回值 = Lib.同花顺.setPeriod(period)
	Call Lib.框架函数1.loopProcessTHSStockFile("d:\gupiao\data\basic\持仓.txt","获取股票数据",0)	
	Call Lib.框架函数1.loopProcessTHSStockFile("d:\gupiao\data\basic\自选股票.txt","获取股票数据",0)	
End Function

Function chiCangStock
	Plugin.Msg.Tips "查找持仓股票的卖点......"
	Call Lib.公用函数.删除文件("d:\gupiao\data\同花顺卖出.json")
	返回值 = Lib.同花顺.setKlineStyle()
	返回值 = Lib.同花顺.setFivePicStyle()
	返回值 = Lib.同花顺.setPeriod("5分钟")

	Call Lib.公用函数.删除文件("d:\gupiao\data\高抛.json")
	Call Lib.公用函数.删除文件("d:\gupiao\data\低吸.json")


	Call Lib.框架函数1.loopProcessTHSStockFile("d:\gupiao\data\basic\持仓.txt","T0高抛低吸",0)	
	ret2 = Plugin.File.IsFileExist("d:\gupiao\data\高抛.json")
	If ret2 = True Then 
		Call Lib.框架函数1.loopProcessTHSStockFile("d:\gupiao\data\高抛.json","卖出股票",0)	
	End If
	ret3 = Plugin.File.IsFileExist("d:\gupiao\data\低吸.json")
	If ret3 = true Then 
		Call Lib.框架函数1.loopProcessTHSStockFile("d:\gupiao\data\低吸.json", "买入股票",0)
	End If
End Function

Function zhaoMaiDian

End Function
