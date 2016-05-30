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
MacroID=9087859d-af51-4e9c-99ae-9e8044305f59
Description=我的脚本2_自动备份
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
LogStart "d:\gupiao\log\T0.log"
thsHwnd = Lib.同花顺.getThsHwnd()
Delay 1000
//MessageBox "thsHwnd=" & thsHwnd
If thsHwnd > 0 Then 
	call Lib.同花顺.setFivePicStyle()
	Delay 1000
	Call Lib.公用函数.InputKeyboardCode("300310")
	Delay 1000
	Call T0GaoPaoDiXi("300310")
End If
LogStop


Function T0GaoPaoDiXi(stockCode)
	pageDataStr = Lib.同花顺.getPageData(stockCode)
	//MessageBox "返回数据=" & pageDataStr
	pageStr = Split(pageDataStr, "|")
	Dim pageData(25)
	If UBound(pageStr) >= 0 Then 
		For i = 0 To UBound(pageStr)
			pageData(i) = Cdbl(pageStr(i))
		Next
	End If
	ma5 = pageData(0)
	ma10 = pageData(1)
	ma20 = pageData(2)
	ma30 = pageData(3)
	ma40 = pageData(4)
	ma50 = pageData(5)
	newPrice = pageData(6)
	openPrice = pageData(7)
	highPrice = pageData(8)
	lowPrice = pageData(9)
	weibi = pageData(10)
	weicha = pageData(11)
	waipan = pageData(12)
	neipan = pageData(13)
	liangbi = pageData(14)
	cciValue = pageData(15)
	k = pageData(16)
	d = pageData(17)
	j = pageData(18)
	macd = pageData(19)
	diff = pageData(20)
	dea = pageData(21)
	hsl = pageData(22)
	hsl5 = pageData(23)
	hsl10=pageData(24)
	
	dataStr = "newPrice=" & newPrice & ",ma5=" & ma5& ",ma10=" & ma10& ",ma20=" & ma20& ",ma30=" & ma30	& ",ma40=" & ma40& ",ma50=" & ma50	& ",ma60=" & ma60& ",openPrice=" & openPrice& ",highPrice=" & highPrice	& ",lowPrice=" & lowPrice& ",weibi=" & weibi& ",weicha=" & weicha& ",waipan=" & waipan& ",neipan=" & neipan	& ",liangbi=" & liangbi	& ",cci=" & cci	& ",k=" & k	& ",d=" & d	& ",j=" & j	& "macd=" & macd& ",diff=" & diff& ",dea=" & dea & ",hsl=" &hsl & ",hsl5=" & hsl5 & ",hsl10=" & hsl10
	//MessageBox dataStr
	If cciValue > - 150 and cciValue < -50 Then 
		//CCI进入低吸
		If k < 25 and d < 25 and j < 30 Then 
			diffValue = abs(newPrice - lowPrice) / lowPrice
			If diffValue < 0.01 Then 
				MessageBox "低吸:" & dataStr
			Else 
				TracePrint stockCode & "diffValue=" & diffValue & "没有在最低点" & dataStr
			End If
			MessageBox "低吸：" & dataStr
		Else 
			TracePrint stockCode & "CCI=" & cciValue & ",KDJ没有在高位区域" & dataStr
		End If
	Else 
		TracePrint stockCode & " CCI Value：" & cciValue & " 小于-100" & dataStr
	End If
	
	If cciValue < 150 and cciValue > 50 Then 
		//高抛
		If k > 80 and d > 80 and j > 80 Then 
		 	diffValue = abs(newPrice - highPrice) / highPrice
		 	If diffValue < 0.01 Then 
		 		MessageBox "高抛:" & dataStr
		 	Else 
		 		TracePrint stockCode & "diffValue=" & diffValue & "没有在最高点" & dataStr
		 	End If
		Else 
			TracePrint stockCode & "KDJ在低位" & dataStr
		End If
	End If

End Function
Function T0DiXi(stockCode)
	T0DiXi = -1
	pageDataStr = Lib.同花顺.getPageData(stockCode)
	//MessageBox "返回数据=" & pageDataStr
	pageStr = Split(pageDataStr, "|")
	Dim pageData(25)
	If UBound(pageStr) >= 0 Then 
		For i = 0 To UBound(pageStr)
			pageData(i) = Cdbl(pageStr(i))
		Next
	End If
	ma5 = pageData(0)
	ma10 = pageData(1)
	ma20 = pageData(2)
	ma30 = pageData(3)
	ma40 = pageData(4)
	ma50 = pageData(5)
	newPrice = pageData(7)
	openPrice = pageData(7)
	highPrice = pageData(8)
	lowPrice = pageData(9)
	weibi = pageData(10)
	weicha = pageData(11)
	waipan = pageData(12)
	neipan = pageData(13)
	liangbi = pageData(14)
	cciValue = pageData(15)
	k = pageData(16)
	d = pageData(17)
	j = pageData(18)
	macd = pageData(19)
	diff = pageData(20)
	dea = pageData(21)
	hsl = pageData(22)
	hsl5 = pageData(23)
	hsl10=pageData(24)
	
	dataStr = "newPrice=" & newPrice & ",ma5=" & ma5& ",ma10=" & ma10& ",ma20=" & ma20& ",ma30=" & ma30	& ",ma40=" & ma40& ",ma50=" & ma50	& ",ma60=" & ma60& ",openPrice=" & openPrice& ",highPrice=" & highPrice	& ",lowPrice=" & lowPrice& ",weibi=" & weibi& ",weicha=" & weicha& ",waipan=" & waipan& ",neipan=" & neipan	& ",liangbi=" & liangbi	& ",cci=" & cci	& ",k=" & k	& ",d=" & d	& ",j=" & j	& "macd=" & macd& ",diff=" & diff& ",dea=" & dea & ",hsl=" &hsl & ",hsl5=" & hsl5 & ",hsl10=" & hsl10
		
	//MessageBox dataStr
	MessageBox "before cciValue=" & cciValue
	If cciValue > -200 and cciValue < -50 Then 
		MessageBox "cciValue=" & cciValue
		//cci进入低吸区域
		cciStatus = Lib.同花顺.isUpArrow("CCI")
		If cciStatus = true Then 
			//CCI进入上升趋势
			If k < 25 and d < 25 and j < 25 Then 
				macdStatus = Lib.同花顺.isUpArrow("MACD")
				//If macdStatus = true Then 
					//macd进入上升通道
					diffValue = Abs(newPrice - lowPrice) / lowPrice
					If diffValue < 0.01 Then 
						T0diXi = newPrice
					Else 
						T0DiXi = -4		//价格不在低位区域
					End If
				//End if
			Else 
				T0DiXi = -3			//KDJ没有进入超卖区域
			End If
		Else 
			T0DiXi = -2		//CCI方向向下
		End If
	Else 
		T0DiXi = -1		//CCI值> -100
	End If
End Function

Function T0GaoPao(stockCode)
	T0GaoPao = -1
	pageDataStr = Lib.同花顺.getPageData(stockCode)
	//MessageBox "返回数据=" & pageDataStr
	pageStr = Split(pageDataStr, "|")
	Dim pageData(25)
	If UBound(pageStr) >= 0 Then 
		For i = 0 To UBound(pageStr)
			pageData(i) = pageStr(i)
		Next
	End If
	ma5 = pageData(0)
	ma10 = pageData(1)
	ma20 = pageData(2)
	ma30 = pageData(3)
	ma40 = pageData(4)
	ma50 = pageData(5)
	newPrice = pageData(7)
	openPrice = pageData(7)
	highPrice = pageData(8)
	lowPrice = pageData(9)
	weibi = pageData(10)
	weicha = pageData(11)
	waipan = pageData(12)
	neipan = pageData(13)
	liangbi = pageData(14)
	cci = pageData(15)
	k = pageData(16)
	d = pageData(17)
	j = pageData(18)
	macd = pageData(19)
	diff = pageData(20)
	dea = pageData(21)
	hsl = pageData(22)
	hsl5 = pageData(23)
	hsl10=pageData(24)
	
	dataStr = "newPrice=" & newPrice & ",ma5=" & ma5& ",ma10=" & ma10& ",ma20=" & ma20& ",ma30=" & ma30	& ",ma40=" & ma40& ",ma50=" & ma50	& ",ma60=" & ma60& ",openPrice=" & openPrice& ",highPrice=" & highPrice	& ",lowPrice=" & lowPrice& ",weibi=" & weibi& ",weicha=" & weicha& ",waipan=" & waipan& ",neipan=" & neipan	& ",liangbi=" & liangbi	& ",cci=" & cci	& ",k=" & k	& ",d=" & d	& ",j=" & j	& "macd=" & macd& ",diff=" & diff& ",dea=" & dea & ",hsl=" &hsl & ",hsl5=" & hsl5 & ",hsl10=" & hsl10
	If newPrice < ma5 Then 
		T0GaoPao = newPrice
	End If
End Function

Function T0Sell
	filename = "d:\gupiao\data\t0stock.xlsx"
	status = Plugin.File.IsFileExist(filename)
	//MessageBox "T0Stock exist:" & status
	If status = False Then 
		T0Sell = - 1 
		Exit Function
	End If
	Call Plugin.Office.OpenXls(filename)
	row = 1
	stockCode = Plugin.Office.ReadXls(1, row, 1)
	While stockCode <> ""
		//MessageBox "stockCode=" & stockCode
		Call Lib.公用函数.InputKeyboardCode(stockCode)
		Delay 1000
		status = Plugin.Office.ReadXls(1, row, 2)
		//MessageBox "status=" & status
		Select Case status
			Case 1
				//找买点
				buyStatus = T0DiXi(stockCode)
				MessageBox "buyStatus = " & buyStatus
			Case 2
				//买入
				//Plugin.Office.WriteXls (1,row,5, Now)			
				//plugin.Office.WriteXls(1,row,2,"3")
			Case 3
				//找卖点
				sellStatus = T0GaoPao(stockCode)
				/*
				If sellStatus > 0 Then 
					//Plugin.Office.WriteXls (1,row,7,sellStatus)
					//Plugin.Office.WriteXls (1,row,2, "4")
				End If
				*/
			Case 4
				//卖出
				//Plugin.Office.WriteXls(1,row, 8, Now)
				//plugin.Office.WriteXls(1,row,2, "5")
		End Select

		row = row + 1
		stockCode = Plugin.Office.ReadXls(1,row,1)
	Wend
	Call Plugin.Office.CloseXls()
End Function


Function addT0BuyDataItem(stockCode, buyTime, price, count)
	
End Function
