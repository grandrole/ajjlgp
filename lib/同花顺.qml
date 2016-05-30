[General]
SyntaxVersion=2
MacroID=5c2bac73-1cb5-41d9-a4df-26938f944b49
[Comment]

[Script]
//请在下面写上您的子程序或函数
//写完保存后，在任一命令库上点击右键并选择“刷新”即可

//测试启动/停止同花顺软件
Function testStartStopThs
	hwnd = getThsHwnd()
	If hwnd < 0 Then 
		MessageBox "启动同花顺失败！"
	Else 
		Plugin.Window.Active (hwnd)
		Call Lib.公用函数.MouseClick(500,300)
		Delay 1000
	End If
	
	MessageBox hwnd
End Function



//同花顺.lib

Function startThs
	TracePrint "启动同花顺股票软件.......启动"
	RunApp "D:\同花顺软件\同花顺\hexin.exe"
	Delay 3000
	TracePrint "启动同花顺股票软件 ........ 结束"
	startThs = 1
End Function


Function stopThs
	ret = Lib.公用函数.getHwnd("同花顺(")
	If ret < 0 Then 
		stopThs = -1 
		Exit Function
	End If
	//Alt + F4 关闭窗口
	KeyDown 18, 1
	KeyPress 115, 1
	KeyUp 18, 1
	delay 1000
	
	FindPic 788,501,933,568,"D:\gupiao\images\ths_stop_1.bmp",0.9,intX,intY
	If intX > 0 And intY > 0 Then 
		//点击“退出”按钮
		Call Lib.公用函数.MouseClick(878,531)
		Delay 1000
		stopThs = 1
	Else 
		stopThs = -2
	End If
End Function

/*
获取同花顺Hwnd 
参数：无 
返回值： 
	-20		启动同花顺成功,没有找到同花顺Hwnd
	>0		返回同花顺Hwnd
*/

Function getThsHwnd
	hwnd = Lib.公用函数.getHwnd("同花顺(")
	If hwnd < 0 Then 
		call startThs()
		Delay 5 * 1000
		subHwnd = Lib.公用函数.getHwnd("同花顺(")
		If subHwnd < 0 Then 
			getThsHwnd = -20 
		Else 
			getThsHwnd = subHwnd
			Call Plugin.Window.Active(subHwnd)
		End If
	Else 
		//返回同花顺Hwnd
		getThsHwnd = hwnd
		Call Plugin.Window.Active(hwnd)
		Delay 1000
	End If
End Function

Function getMainPicStyle()
	If 1 = isKLineStyle() Then 
		getMainPicStyle = "KLineStyle"
	ElseIf 1 = isENEStyle() Then
	 	getMainPicStyle = "ENEStyle"
	 ElseIf 1 = isFenShiStyle() Then
	 	getMainPicStyle = "FenShiStyle"
	 Else 
	 	getMainPicStyle = "Other"
	 End if
End Function

Function setMainPicStyle(style)
	For i = 1 To 3
		tmpStyle = getMainPicStyle()
		If tmpStyle = "Other" Then 
			Call Lib.公用函数.InputKeyboardCode("002460")
			Delay 500		
			tmpStyle = getMainPicStyle()
		End If
	
		If tmpStyle = "FenShiStyle" Then 
			KeyPress "Enter", 1
			Delay 500
			tmpStyle = getMainPicStyle()
		End If
	
		If style = tmpStyle Then 
			TracePrint "目前已经为" & style & "模式，无须修改"
			Exit function
		End If
	
		If style = "ENEStyle" and tmpStyle = "KLineStyle" Then 
			Call setENEStyle()
			Delay 1000
		End If
	
		If style = "KLineStyle" and tmpStyle = "ENEStyle" Then 
			KeyPress "Tab", 1
			Delay 1000
		End If
		
	Next

End Function

Function setENEStyle()
	//设置ENE方式
	setENEStyle = 0
	For i = 1 To 3
		If 1 = isENEStyle() Then 
		 	setENEStyle = 1
			Exit function
		else
			Call Lib.公用函数.InputKeyboardCode("300310")
			Call Lib.公用函数.MouseClick(495,136)
			Call Lib.公用函数.InputKeyboardCode("ene")
			Delay 1000
		End if
	Next
End Function

Function setKLineStyle()
	setKLineStyle = 0
	Call Lib.公用函数.InputKeyboardCode("300310")
	For i = 1 To 3
		If 1 = isKLineStyle() Then 
			setKLineStyle = 1
			Exit Function
		Else 
			KeyPress Enter, 1
			Delay 1000
		End If
	next
End Function

Function isKLineStyle()
	FindPic 37,54,223,96,"D:\gupiao\images\ma5.bmp",0.9,intX,intY
	If intX > 0 And intY > 0 Then 
		isKLineStyle = 1
	Else 
		isKLineStyle = 0
	End If
End function

Function isENEStyle()
	FindPic 66,37,262,147,"D:\gupiao\images\eneImage.bmp",0.9,intX,intY
	If intX > 0 And intY > 0 Then 
		isENEStyle = 1
	Else 
		isENEStyle = 0
	End If
End Function

Function isFenShiStyle()
	FindPic 90,45,234,115,"D:\gupiao\images\分时.bmp",0.9,intX,intY
	If intX > 0 And intY > 0 Then 
		isFenShiStyle = 1
	Else 
		isFenShiStyle = 0
	End If
End Function

Function setAddtionalPicStyle
		Call setMainPicStyle("KLineStyle")
		Delay 1000
		//点击“窗“
		FindPic 929,41,1072,109,"D:\gupiao\images\窗.bmp",0.9,intX,intY
		If intX > 0 And intY > 0 Then 
			//MessageBox "intX=" & intX & ",intY=" & intY
			Call Lib.公用函数.MouseClick(intX, intY)
			Delay 500
		Else 
			MessageBox "没有找到“窗”字样"
			Exit function
		End If

		//点击”6图组合“
		Call Lib.公用函数.MouseClick(1047,200)
		Delay 1000
		//点击3图
		Call Lib.公用函数.MouseClick(272,396)
		//设置cci
		Call Lib.公用函数.InputKeyboardCode("cci")
		Delay 1000
		//点击4图
		Call Lib.公用函数.MouseClick(572,482)
		//设置kdj
		Call Lib.公用函数.InputKeyboardCode("boll")
		Delay 1000
		//点击5图
		Call Lib.公用函数.MouseClick(519,530)
		//设置macd
		Call Lib.公用函数.InputKeyboardCode("kdj")
		Delay 1000
		//点击6图
		Call Lib.公用函数.MouseClick(440,588)
		//设置macd
		Call Lib.公用函数.InputKeyboardCode("macd")
		Delay 1000	
End Function


Function thsStopImage
	FindPic 456,186,906,459,"D:\gupiao\images\同花顺停止.bmp",0.9,intX,intY
	If intX > 0 And intY > 0 Then 
		thsStopImage = 1
		
		KeyDown 18, 1
		KeyPress 67, 1
		KeyUp 18, 1

		Call stopThs()
		MessageBox "出现同花顺停止界面"
	Else 
		thsStopImage = 0
	End If
End Function

Function setPeriod(interval)
	Select Case interval
		Case "5分钟"
			value = "32"
			imgSrc = "5分钟.bmp"
		Case "30分钟"
			value = "34"
			imgSrc = "30分钟.bmp"
		Case "60分钟"
			value = "35"
			imgSrc = "60分钟.bmp"
		Case "day"
			value = "36"
			imgSrc = "day.bmp"
		Case "week"
			value = "37"
			imgSrc = "week.bmp"
		Case "month"
			value = "38"
			imgSrc = "month.bmp"
	End Select
	cmpImgSrc = "d:\gupiao\images\" & imgSrc
	
	//MessageBox "value=" & value & ",cmpImgSrc=" & cmpImgSrc
	ret = -1
	For i = 0 To 5
		ret = Lib.公用函数.verifyInputCode(value, 0, 0, 63, 90, cmpImgSrc)
		TracePrint "判断imgSrc是否正确：" & ret
		If ret > 0 Then 
			Exit For
		End If
	Next
	If ret < 0 Then 
		setPeriod = -1
		Exit Function
	Else 
		setPeriod = 1
	End If
End Function

Function getPeriod
	Dim cmpImg
	
	cmpImg = Array("month", "week", "day", "60分钟", "30分钟","5分钟")
	For i = 0 To 5
		cmpImgSrc = "d:\gupiao\images\" & cmpImg(i) & ".bmp"
		//MessageBox "getPeriod img=" & cmpImgSrc
		intX = - 1 
		intY = - 1 
		FindPic 19,64,77,98,cmpImgSrc,1,intX,intY
		If intX> 0 And intY> 0 Then
			getPeriod = cmpImg(i)
			Exit Function
		End If
	Next
	getPeriod = -1
End Function


Sub addSelfStock(stockCode)
	Call Lib.公用函数.InputKeyboardCode(stockCode)
	Delay 1000
	KeyPress "Insert", 1
End Sub



Sub removeAllSelfStock
	TracePrint "删除同花顺自选股票"
	Call Lib.公用函数.InputKeyboardCode("06")
	//MessageBox "zxg"
	Delay 1000
	KeyPress "Delete", 200
	Delay 3000
End Sub


Function isUpArrow(x1,y1,x2,y2)
	FindPic x1, y1, x2, y2, "D:\gupiao\images\up.bmp", 0.9, intX, intY
	//MessageBox x1 & "," & y1 & "," & x2 & "," & y2 & ",pos=" & intX & "," & intY
	If intX> 0 And intY> 0 Then
		isUpArrow = "1"
	Else 
		isUpArrow = "0"
	End If	
	
End Function


Function getKLineData
		 //MessageBox "get KLine Data"
		ma5 = getThsDataItem(145, 69, 255, 86, "ffffff-000000")//0
		ma10 = getThsDataItem(197,69,329,86,"ffff0b-000000") //1
		ma20 = getThsDataItem(293,69,425,86,"ff80ff-000000") //2
		ma30 = getThsDataItem(368,69,500,86,"00e600-000000") //3
		ma40 = getThsDataItem(485,70,614,88,"02e2f4-000000") //4
		ma5Up = isUpArrow(145, 69, 255, 86)
		ma10Up = isUpArrow(197,69,329,86)
		ma20Up = isUpArrow(293,69,425,86)
		ma30Up = isUpArrow(368,69,500,86)
		ma40Up=isUpArrow(485,70,614,88)
		getKLineData = ma5 & "|" & ma5Up & "|" & ma10 & "|" & ma10Up & "|"  & ma20 & "|" & ma20Up & "|" & ma30 & "|" & ma30Up & "|" &ma40 & "|" & ma40Up
End Function

Function getEneData
	upperEne = getThsDataItem(228, 69, 391, 88, "ffffff-000000")
	upperEneUp = isUpArrow(228, 69, 391, 88)
	lowerEne = getThsDataItem(352, 68, 469, 87, "ffff0b-000000")
	lowerEneUp = isUpArrow(352, 68, 469, 87)
	ene = getThsDataItem(441,69,558,88,"008040-000000")
	eneUp = isUpArrow(441, 69, 558, 88)
	getEneData = upperEne & "|" & upperEneUp & "|" & lowerEne & "|" & lowerEneUp & "|" & ene & "|" & eneUp
End Function

Function getTradeData
	newPrice = getThsDataItem(1163,331,1249,349,"00e600-000000|ff3232-000000|ffffff-000000")	//6
	openPrice = getThsDataItem(1283,331,1368,346,"00e600-000000|ff3232-000000|ffffff-000000")	//7
	highPrice = getThsDataItem(1284,351,1369,366,"00e600-000000|ff3232-000000|ffffff-000000")	//8
	lowerPrice = getThsDataItem(1282, 368, 1367, 383, "00e600-000000|ff3232-000000|ffffff-000000")//9
	weibi = getThsDataItem(1161, 96, 1267, 113, "ff3232-000000|00e600-000000")
	
	weicha = getThsDataItem(1261,96,1367,113, "ff3232-000000|00e600-000000")
	waipan = getThsDataItem(1178,466,1246,484, "ff3232-000000|00e600-000000")
	neipan = getThsDataItem(1296,466,1364,484, "ff3232-000000|00e600-000000")
	liangbi = getThsDataItem(1298,381,1366,399, "ff3232-000000|00e600-000000")		
	getTradeData =  newPrice & "|" & openPrice & "|" & highPrice & "|" & lowerPrice & "|" & weibi & "|" & weicha & "|" & waipan & "|" & neipan & "|" & liangbi
End Function

Function getCCIData()
	cciValue = getThsDataItem(85,392,185,410,"ffffff-000000")
	cciUp = isUpArrow(85,392,185,410)
	getCCIData = cciValue & "|" & cciUp
End Function

Function getKDJData
	kValue = getThsDataItem(87,529,173,547, "f0f888-000000")
	dValue = getThsDataItem(177,528,263,546,"54fcfc-000000")
	jValue = getThsDataItem(244,529,330,547, "ff80ff-000000")
	kUp = isUpArrow(87,529,173,547)
	dUp = isUpArrow(177,528,263,546)
	jUp = isUpArrow(244,529,330,547)
	getKDJData = kValue & "|" & kUp & "|" & dValue & "|" & dUp & "|" & jValue & "|" & jUp
End Function

Function getBollData
	midBoll = getThsDataItem(97,470,183,488, "02e2f4-000000")
	midBollUp = isUpArrow(97,470,183,488)
	upperBoll = getThsDataItem(177,469,263,487, "ff3232-000000")
	upperBollUp = isUpArrow(177,469,263,487)
	lowerBoll = getThsDataItem(279,470,379,488, "00e600-000000")
	lowerBollUp = isUpArrow(279,470,379,488)
	getBollData = midBoll & "|" & midBollUp & "|"  & upperBoll & "|" & upperBollUp & "|"  & lowerBoll & "|" & lowerBollUp
End Function

Function getCJLData
	cjl = getThsDataItem(53,325,139,343, "ff3232-000000|54fcfc-000000")
	cjlUp = isUpArrow(53,325,139,343)
	cjl5 = getThsDataItem(122,326,208,344, "ffff0b-000000")
	cjl5Up = isUpArrow(122,326,208,344)
	cjl10 = getThsDataItem(208,325,294,343, "ffffff-000000")
	cjl10Up = isUpArrow(208,325,294,343)
	getCJLData = cjl & "|" & cjlUp &"|"  & cjl5 & "|" & cjl5Up & "|" & cjl10 & "|" & cjl10Up	
End Function

Function getMacdData
	macd = getThsDataItem(134,579,220,597,"ff3232-000000|00e600-000000")
	macdUp = isUpArrow(134,579,220,597)
	diff = getThsDataItem(231,580,317,598,"ffffff-000000")
	diffUp = isUpArrow(231,580,317,598)
	dea = getThsDataItem(306,579,392,597,"ffff0b-000000")
	deaUp = isUpArrow(306,579,392,597)
	getMacdData = macd & "|" & macdUp &"|"  & diff & "|" & diffUp & "|" & dea & "|" & deaUp
End Function


Function getPageDataStr(stockCode)

	//MessageBox "mStyle=" & mStyle
	tradeStr = getTradeData()
	cciStr = getCCIData()
	kdjStr = getKDJData()
	bollStr = getBollData()
	cjlStr = getCJLData()
	macdStr = getMacdData()
	tempStr = tradeStr & "|" & cciStr & "|" & kdjStr & "|" & bollStr & "|" & cjlStr & "|" & macdStr
	
	mStyle = getMainPicStyle()
	If "KLineStyle" = mStyle Then 
		kLineStr = getKLineData()
		getPageDataStr = kLineStr & "|" & tempStr
	End If
	
	If mStyle = "ENEStyle" Then 
		eneStr = getEneData()	
		getPageDataStr = eneStr & "|" & tempStr
	End If
	
	/*
	dataStr = "{" & vbCrLf
	dataStr = dataStr & vbTab & "newPrice=" & newPrice & "," & vbCrLf
	dataStr = dataStr & vbTab & "ma5=" & ma5 & "," & vbCrLf
	dataStr = dataStr & vbTab & "ma10=" & ma10 & "," & vbCrLf
	dataStr = dataStr & vbTab & "ma20=" & ma20 & "," & vbCrLf
	dataStr = dataStr & vbTab & "ma30=" & ma30 & "," & vbCrLf
	dataStr = dataStr & vbTab & "ma40=" & ma40 & "," & vbCrLf
	dataStr = dataStr & vbTab & "ma50=" & ma50 & "," & vbCrLf
	dataStr = dataStr & vbTab & "openPrice=" & openPrice & "," & vbCrLf
	dataStr = dataStr & vbTab & "highPrice=" & highPrice & "," & vbCrLf
	dataStr = dataStr & vbTab & "lowerPrice=" & lowerPrice & "," & vbCrLf
	dataStr = dataStr & vbTab & "weibi=" & weibi & "," & vbCrLf
	dataStr = dataStr & vbTab & "weicha=" & weicha & "," & vbCrLf
	dataStr = dataStr & vbTab & "neipan=" & neipan & "," & vbCrLf
	dataStr = dataStr & vbTab & "waipan=" & waipan & "," & vbCrLf
	dataStr = dataStr & vbTab & "liangbi=" & liangbi & "," & vbCrLf
	dataStr = dataStr & vbTab & "cci=" & cciValue & "," & vbCrLf
	dataStr = dataStr & vbTab & "k=" & kValue & "," & vbCrLf
	dataStr = dataStr & vbTab & "d=" & dValue & "," & vbCrLf
	dataStr = dataStr & vbTab & "j=" & jValue & "," & vbCrLf
	dataStr = dataStr & vbTab & "macd=" & macd & "," & vbCrLf
	dataStr = dataStr & vbTab & "diff=" & diff & "," & vbCrLf
	dataStr = dataStr & vbTab & "dea=" & dea & "," & vbCrLf
	dataStr = dataStr & vbTab & "hsl=" & hsl & "," & vbCrLf
	dataStr = dataStr & vbTab & "hsl5=" & hsl5 & "," & vbCrLf
	dataStr = dataStr & vbTab & "hsl10=" & hsl10 & vbCrLf
	dataStr = dataStr & "}"
	
	getPageDataStr = dataStr
	*/
	
End Function

Function getThsDataItem(x1, y1, x2, y2, color)
set dm = createobject("dm.dmsoft")
//TracePrint dm
dm_ret = dm.SetDict(0,"d:\gupiao\data\font\ths.txt")
s = dm.Ocr(x1, y1, x2, y2, color, 1.0)
//MessageBox "s=" & s
If s <> "" Then 
	getThsDataItem = Lib.公用函数.trimString(s)
Else 
	getThsDataItem = -1
End If
End Function



