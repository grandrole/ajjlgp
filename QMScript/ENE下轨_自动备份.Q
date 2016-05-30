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
MacroID=96818159-d2b7-4123-abfd-1fd2a7ed6f9d
Description=ENE下轨_自动备份
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
返回值 = Lib.同花顺.getThsHwnd()
//MessageBox "返回值=" & 返回值
ret = setENEStyle()
MessageBox "ret=" & ret

Call Lib.公用函数.删除文件("d:\gupiao\data\ENE下轨.json")
call loopProcessTHSStockFile("d:\gupiao\data\basic\正常股票.txt", "ENE下轨", 0)



Sub loopProcessTHSStockFile(filename,flowStr, position)

	TracePrint "同花顺【" & filename & "】文件循环" & flowStr & "...... 开始"

    thsHwnd = Lib.同花顺.getThsHwnd()
	If thsHwnd < 0 Then 
		msgStr = "不能正常启动同花顺股票软件，请手工启动！"
		showErrorMsg(msgStr)
		Exit Sub
	End If
		
    stockFile = Plugin.File.OpenFile(filename)
    If stockFile = - 1  Then 
    	msgStr = "打开" & filename & "文件失败"
		showErrorMsg(msgStr)
        Exit Sub
    End If
   
   stockCode = Plugin.File.ReadLine(stockFile)
   If position > 1 Then 
    	num = position
    	While num > 0
    		stockCode = Plugin.File.ReadLine(stockFile)
    		num = num - 1
    	Wend   	
   End If

	Call Lib.同花顺.showSingleStockPage()

	If InStr(flowStr, "ENE") Then 
		setENEStyle()
		setPeriod ("day")
	Else 
		setKLineStyle()
	End If

    no = 1
    globalNo = position
    
    While stockCode <> ""
        If no > 5000 Then 
        	Call Plugin.File.WriteINI("ths", "loopNum", globalNo, "d:\gupiao\data\Config.ini")
        	Delay 1000
        	thsHwnd1 = Lib.同花顺.getThsHwnd()

			If thsHwnd1 > 0 Then 
            	Call Lib.同花顺.showSingleStockPage()
            	If flowStr = "均线纠结" or flowStr = "T0高抛低吸" or flowStr = "均线多头排列" Then 
    				//返回值 = Lib.同花顺.setFivePicStyle()
    			End If
				no = 1
            End if
        Else 
            no = no + 1
        End If
        globalNo = globalNo + 1
        Call Lib.公用函数.InputKeyboardCode(stockCode)
        //光标移动到最新数据区
        Call Lib.公用函数.MouseClick(1173,267)

		Delay 1000
		
		TracePrint "处理" & stockCode & ",操作:" & flowStr
        If toEneDown(stockCode) = true Then 
        	Call Lib.公用函数.addStockInFilename("d:\gupiao\data\ENE下轨.json",stockCode)
        End If
        str = Plugin.File.ReadLine(stockFile)
        stockCode = Mid(str, 2)
    Wend
    Call Plugin.File.CloseFile(stockFile)
    Call Plugin.File.WriteINI("ths", "loopNum", "0", "d:\gupiao\data\Config.ini")
	TracePrint "同花顺【" & filename & "】文件循环" & flowStr & "...... 结束"
    
End Sub


Function toEneDown(stockCode)
	pageDataStr = getEneTradeData(stockCode)
	//MessageBox "dataStr=" & dataStr
	pageStr = Split(pageDataStr, "|")
	Dim pageData(50)
	If UBound(pageStr) >= 0 Then 
		For i = 0 To UBound(pageStr)
			If pageStr(i) = "-1" Then 
				pageData(i) = - 1 
				msgStr = "数据采集错误：" & i & "数据=" & pageDataStr
				showErrorMsg(msgError)
			else
				pageData(i) = Cdbl(pageStr(i))
			End if
		Next
	End If
	upperEne = pageData(0)
	upperEneUp = pageData(1)
	lowerEne = pageData(2)
	lowerEneUp = pageData(3)
	ene = pageData(4)
	eneUp = pageData(5)
	newPrice = pageData(6)
	openPrice = pageData(7)
	highPrice = pageData(8)
	lowerPrice = pageData(9)
	
	If lowerPrice <= lowerEne Then 
		toENEDown = true
	Else 
		toENEDown = false
	End If
End Function


Function getEneTradeData(stockCode)
	eneStr = getEneData()
	tradeStr = getTradeData()
	getEneTradeData = eneStr & "|" & tradeStr
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

Function isUpArrow(x1, y1, x2, y2)

	FindPic x1, y1, x2, y2, "D:\gupiao\images\up.bmp", 0.9, intX, intY
	//MessageBox x1 & "," & y1 & "," & x2 & "," & y2 & ",pos=" & intX & "," & intY
	If intX> 0 And intY> 0 Then
		isUpArrow = "1"
	Else 
		isUpArrow = "0"
	End If	
	
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


//公共函数
Sub showErrorMsg(msgStr)
	Plugin.Msg.Tips msgStr
    TracePrint msgStr	
End Sub






