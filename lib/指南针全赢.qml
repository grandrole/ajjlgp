[General]
SyntaxVersion=2
MacroID=795072df-f671-4227-8ccc-219b2d9e888e
[Comment]

[Script]
Sub setCPJCStyle()
	Call Lib.公用函数.InputKeyboardCode("000901")
	Delay 3000
	Call Lib.公用函数.InputKeyboardCode("cpjc")
	FindPic 0,0,1366,768,"D:\gupiao\images\指标叠加.bmp",0.9,intX,intY
	If intX> 0 And intY> 0 Then
		KeyPress "N", 1
	End If
	Delay 3000
End Sub

Function findHuiYan(stockCode)
If isStockImage(stockCode) = False Then 
		//不是指定股票图形
		findHuiYan = -1 
		Exit Function
	End If
	
	If isCurrentDate() = False Then 
		//股票日期不对
		findHuiYan = -2
		Exit Function
	End If
	FindPic 992,118,1038,349,"D:\gupiao\images\huiyan.bmp",0.9,intX1,intY1
	If intX1 > 0 And intY1 > 0 Then 
		//发现慧眼
		findHuiYan = 1
	Else 
		FindColor 992,118,1038,349,"FF8000",intX2,intY2
		If intX2 > 0 And intY2 > 0 Then 
			//发现蓝色-超跌慧眼
			findHuiYan = 2
		Else 
			FindColor 992, 118, 1038, 349, "FF00FF", intX3, intY3
			If intX3 > 0 and intY3 > 0 Then 
				//发现紫色-高抛低吸慧眼
				findHuiYan = 3
			Else 
				FindColor 992, 118, 1038, 349, "0000FF", intX4, intY4
				If intX4 > 0 and intY4 > 0 Then 
				 	//发现红色-追涨慧眼
					findHuiYan = 4
				Else 
					findHuiYan = -1
				End If
			End If
		End If	
	End If
End Function


	
Function findBuyImage(stockCode)
	/*
	Call Lib.公用函数.InputKeyboardCode(stockCode)
	Delay 1000
	*/
	If isStockImage(stockCode) = False Then 
		//不是指定股票图形
		findBuyImage = -1 
		Exit Function
	End If
	
	If isCurrentDate() = False Then 
		//股票日期不对
		findBuyImage = -2
		Exit Function
	End If
	FindPic 1003,122,1023,336,"D:\gupiao\images\B.bmp",0.9,intX1,intY1
	//MessageBox "intX1=" & intX1 & ",intY1=" & intY1
	If intX1 > 0 And intY1 > 0 Then 
		findBuyImage = 1
	Else 
		findBuyImage = 0
	End If
End Function

Function findSellImage(stockCode)
	Call Lib.公用函数.InputKeyboardCode(stockCode)
	Delay 1000
	If isStockImage(stockCode) = False Then 
		//不是指定股票图形
		findSellImage = -1 
		Exit Function
	End If
	
	If isCurrentDate() = False Then 
		//股票日期不对
		findSellImage = -2
		Exit Function
	End If
	FindPic 992,118,1038,349,"D:\gupiao\images\S.bmp",0.9,intX1,intY1
	If intX1 > 0 And intY1 > 0 Then 
		fundsStr = getFundsData(stockCode)
		funds = Split(fundsStr, "|")
		Dim fundsData
		If UBound(funds)>=0 then
        	For i = 0 to UBound(funds)
            	fundsData(i) = CDbl(funds(i))
        	Next
        
        	If fundsData(1) > 0 Then 
       	 		//多空资金为流入状态
       			findSellImage = -3 
       			Plugin.Msg.Tips stockCode & "多空资金为流入状态，不能买出:" & fundsStr
       			TracePrint stockCode & "多空资金为流入状态，不能买出:" & fundsStr
       			Exit Function
       		End If

        	findSellImage = 1
			str = stockCode & " 找到买出点,请快速卖出:" & fundsStr
			TracePrint str
			Plugin.Msg.Tips str
			Call Lib.公用函数.addStockInFilename("d:\gupiao\data\高抛.json", stockCode)
    	End If
	End If
End Function

Sub clearZiXuanGu
	//点击“A股行情"
	Call Lib.公用函数.MouseClick(695, 48)
	Delay 1000
	Call Lib.公用函数.MouseClick(377, 676)
	Delay 1000
	KeyPress "Delete", 500
End Sub

Sub addStockItem(stockCode)
	SayString ","
End Sub


Sub closeCPJCStyle
	
	Call Lib.公用函数.InputKeyboardCode("600007")
	Delay 3000
	Call Lib.公用函数.InputKeyboardCode("cpjc")
	Delay 1000
	KeyPress "Enter", 1
End Sub

Function startHmj
	startHmj = -1 
	//启动红马甲软件
	RunApp "D:\Fix08\WavMain\WavMain.exe"
	//等待6秒
	Delay 3 * 1000
	//查找红马甲标题图片
	Hwnd = Lib.公用函数.getHwnd("指南针红马甲博弈")
	If Hwnd > 0 Then 
		//关闭学习园地窗口
		Call Lib.公用函数.MouseClick(1350,35)
		Delay 1000
		startHmj = Hwnd
	Else 
		startHmj = -1
	End If
End Function

Function stopHmj
	Hwnd = Lib.公用函数.getHwnd("指南针红马甲博弈")
	If Hwnd > 0 Then 
		//关闭红马甲窗口
		Call Lib.公用函数.MouseClick(1347,14)
		Delay 1000
		//按Y键，退出系统
		Call Lib.公用函数.InputKeyboardCodeNoEnter("Y")
		stopHmj = 1
	Else 
		TracePrint "没有找到红马甲窗口"
		stopHmj = -1
	End If
End Function

Function getHmjHwnd
	Hwnd = Lib.公用函数.getHwnd("指南针红马甲博弈")
	If Hwnd > 0 Then 
		getHmjHwnd = Hwnd
	Else 
	 	Hwnd = startHmj()
	 	If Hwnd > 0 Then 
	 		getHmjHwnd = Hwnd
	 	Else 
	 		//启动红马甲软件失败
	 		getHmjHwnd = - 1 
	 	End If
	End If
End Function


Function startZnz
	startZnz = -1

	//Call Lib.API.运行程序("C:\Users\Lenovo\Desktop\5-15指南针私享家\指南针助手.exe")
	RunApp "C:\Users\Lenovo\Desktop\5-15指南针私享家\指南针助手.exe"


	Delay 3*6000
	FindPic 0,0,1024,768,"Attachment:\znz_start_1.bmp",0.9,intX,intY
	If intX > 0 And intY > 0 Then 
		//点击”登录“按钮
		Call Lib.公用函数.MouseClick(648, 432)
		Delay 5000
		FindPic 0, 0, 1024, 768, "Attachment:\znz_start_2.bmp", 0.9, intX, intY
		If intX > 0 and intY > 0 Then 
			//点击”确定“按钮
			Call Lib.公用函数.MouseClick(800, 470)
			Delay 6000
			FindPic 0, 0, 1024, 768, "Attachment:\znz_start_3.bmp", 0.9, intX, intY
			If intX > 0 and intY > 0 Then 
				//点击”确定“按钮
				Call Lib.公用函数.MouseClick(1347,41)
				Delay 1000
				startZnz = 1
			Else 
				//没有显示指南针住界面
				startZnz = - 3 
			End if				
		Else 
			//没有出现指南针验证通过界面
			startZnz = -2
		End If
	Else 
		//没有显示指南针登录验证界面
		startZnz = -1
	End If

End Function


Function stopZnz
	Hwnd = Lib.公用函数.getHwnd("指南针全赢博弈")
	If Hwnd < 0 Then 
	 	//指南针没有启动
	 	stopZnz = - 1 
	 	Exit Function
	End If
	//点击关闭窗口
	ret1 = Lib.公用函数.verifyMouse(1349, 16, 0, 0, 1024, 768, "d:\gupiao\images\znz_stop_1.bmp")
	If ret1 > 0 Then 
		//模仿“Alt+Y" 退出
		KeyDown 18, 1
		KeyPress 89, 1
		KeyUp 18, 1
		Delay 1000
	Else 
		//没有出现退出窗口
		stopZnz = - 2 
	End if
End Function

Function isCurrentDate()

	Call Lib.公用函数.MouseClick(1087, 160)
	Delay 1000
	
	s = getDataItem(971,619,1059,641,"ffffff-000000")
	//Plugin.Msg.Tips "截至日期s=" & s
	//MessageBox "截至日期=" & s
	If s <> "" Then 
		znzDate = CDate(s)
		diffDay = DateDiff("d", Now, znzDate)
		//Plugin.Msg.Tips "时间间隔:" & diffDay
		If diffDay > 0 Then 
			isCurrentDate = true
		Else 
			TracePrint "日期不匹配： " & s & "间隔：" & diffDay
			isCurrentDate = False
		End If
	Else 
		TracePrint "不能识别日期，请检查大漠插件是否安装正确!"
		Plugin.Msg.Tips "不能识别日期，请检查大漠插件是否安装正确!"
	End If	
End Function

Function isStockImage(stockCode)
	s = getDataItem(1069,68,1130,90,"40e0d0-000000")
	//MessageBox "s=" & s
	If s = stockCode Then 
		isStockImage = true
	Else 
		isStockImage = False
	End If

//isStockImage = true

End Function

//获得资金数据
Function getFundsData(stockCode)
	duokong = getDataItem(146,420,234,442, "ff5050-000000|40e0d0-000000")
	gansidui = getDataItem(151, 471, 239, 493, "ff5050-000000|40e0d0-000000")
	zhuli = getDataItem(146,521,225,542,"ff5050-000000|40e0d0-000000")
	zlCangWei = getDataItem(222,568,362,589,"ffff00-000000|ff00ff-000000")
	getFundsData = duokong & "|" & gansidui & "|" & zhuli & "|" & zlCangWei
End Function

//使用大漠插件，获得数据
Function getDataItem(x1, y1, x2, y2, color)
	Set dm = createobject("dm.dmsoft")
	dm_ret = dm.SetDict(0, "d:\gupiao\data\font\znz.txt")
	str = dm.Ocr(x1, y1, x2, y2, color, 1.0)
	//MessageBox "getDataItem str=" & str
	pos = InstrRev(str, ":", - 1 , 1)
	//MessageBox "getDataItem pos=" & pos
	retStr = Mid(str, pos + 1)
	//messageBox("getDataItem retStr=" & retStr)
	getDataItem = retStr
End Function

//获得均线数据
Function getJunXianData(stockCode)
	ma5 = getDataItem(175, 91, 263, 113, "ffff00-000000")
	ma13 = getDataItem(250, 91, 338, 113, "ff00ff-000000")
	ma34 = getDataItem(327, 91, 415, 113, "00ff00-000000")
	maMax = getDataItem(400, 90, 488, 112, "00ffff-000000")
	getJunXianData = ma5 & "|" & ma13 & "|" & ma34 & "|" & maMax
End Function

//获得交易数据
Function getTransData(stockCode)
	Set dm = createobject("dm.dmsoft")
	dm_ret = dm.SetDict(0, "d:\gupiao\data\font\znz.txt")
	//当前价格
	current = getDataItem(1105,209,1219,231, "ff3232-000000|00e600-000000|c0c0c0-000000")
	highPrice = getDataItem(1107,350,1221,372,"ff3232-000000|00e600-000000|c0c0c0-000000")
	lowPrice = getDataItem(1254,351,1368,373, "ff3232-000000|00e600-000000|c0c0c0-000000")
	openPrice = getDataItem(1106,372,1220,394, "ff3232-000000|00e600-000000|c0c0c0-000000")
	huanShou = getDataItem(1106,416,1202,438, "c0c0c0-000000")
	getTransData = current & "|" & highPrice & "|" & lowPrice & "|" & openPrice & "|" & huanShou
	//MessageBox "getTransData=" & getTransData
End Function

Function chaoDieQiBao(stockCode)
	If isStockImage(stockCode) = False Then 
		//不是指定股票图形
		chaoDieQiBao = -1 
		Exit Function
	End If
	
	If isCurrentDate() = False Then 
		//股票日期不对
		chaoDieQiBao = -2
		Exit Function
	End If
	FindPic 985,534,1031,629,"D:\gupiao\images\cdqb.bmp",0.9,intX1,intY1
	If intX1 > 0 And intY1 > 0 Then 
		chaoDieQiBao = 1
	Else 
		FindPic 985,534,1031,629, "D:\gupiao\images\cdqb_2.bmp", 0.9, intX1, intY1
		If intX2 > 0 and intY2 > 0 Then 
			chaoDieQiBao = 1
		Else 
			chaoDieQiBao = -3
		End If
	End If
End Function

Function isTradeDataStyle
	FindPic 1092,634,1164,684,"D:\gupiao\images\分笔.bmp",0.9,intX,intY
	If intX > 0 And intY > 0 Then 
		isTradeDataStyle = true
	Else 
		isTradeDataStyle = false
	End If
End Function

Function setTradeDataStyle
	Call Lib.公用函数.InputKeyboardCode("000901")
	Delay 1000
	status = false
	For i = 1 To 3
		If isTradeDataStyle = True Then 
			setTradeDataStyle = true
			Exit Function
		Else 
			Call Lib.公用函数.MouseClick(1138,675)
			Delay 1000
		End If	
	Next
	Plugin.Msg.Tips "分笔数据设置失败！"
End Function