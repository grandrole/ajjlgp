[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=0
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=3a89eb5c-2766-4244-8ab1-c63d4a789d62
Description=高抛低吸测试
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
Delay 5000
//MessageBox "thsHwnd=" & thsHwnd
If thsHwnd > 0 Then 
	Delay 1000
	Call Lib.同花顺.setFivePicStyle()
	Delay 1000

	Call testGaoPaoDiXi("300365")
	
End If
LogStop 

Function testGaoPaoDiXi(stockCode)
	Call Lib.公用函数.InputKeyboardCode(stockCode)
	Delay 1000

	//点击从左侧开始
	Call Lib.公用函数.MouseClick(130,493)

	For i = 0 To 150
		KeyDown "Right", 1
		Delay 1000
		
		ret = T0GaoPaoDiXi(stockCode)
		If ret = 1 Then 
			MessageBox "买入"
			
		End If

	Next

End Function

Function T0GaoPaoDiXi(stockCode)
	pageDataStr = Lib.同花顺.getPageDataStr(stockCode)
	TracePrint "pageDataStr=" & pageDataStr
	MessageBox "pageDataStr=" & pageDataStr
	periodNo = Lib.同花顺.getPeriod()
	//MessageBox "pageDataStr=" & pageDataStr
	If inStr(pageDataStr, "-1|||||-1|-1|||0.00|") Then 
		//MessageBox "停牌的股票"
		Call Lib.公用函数.addStockInFilename("d:\gupiao\data\删除股票.json",stockCode)
		T0GaoPaoDiXi = - 11 
		Exit Function
	End If
	pageStr = Split(pageDataStr, "|")
	Dim pageData(40)
	If UBound(pageStr) >= 0 Then 
		For i = 0 To UBound(pageStr)
			If pageStr(i) = "-1" Then 
				If mid(stockCode, 1, 1) = "2" Then 
					If i = 16 or i = 17 Then 
						pageData(i) = -1
					End If
				ElseIf periodNo = "5分钟" or periodNo = "30分钟" or periodNo = "60分钟" Then
					If i = 8 Then 
						pageData(i) = -1
					else
						pageData(i) = - 1 
						Plugin.Msg.Tips "数据采集错误：" & i & "数据=" & pageDataStr
						TracePrint "数据采集错误：" & i & "数据=" & pageDataStr
					End if
				End if
			else
				pageData(i) = Cdbl(pageStr(i))
			End if
		Next
	End If
	
	ma5 = abs(pageData(0))
	ma5Up = pageData(1)
	ma11 = abs(pageData(2))
	ma11Up = pageData(3)
	ma31 = abs(pageData(4))
	ma31Up = pageData(5)
	ma63 = abs(pageData(6))
	ma63Up = pageData(7)
	ma250 = abs(pageData(8))
	ma250Up = pageData(9)
	newPrice = abs(pageData(10))
	openPrice = abs(pageData(11))
	highPrice = abs(pageData(12))
	lowerPrice = pageData(13)
	weibi = pageData(14)
	weicha = pageData(15)
	waipan = abs(pageData(16))
	neipan = abs(pageData(17))
	liangbi = pageData(18)
	cciValue = pageData(19)
	cciUp = pageData(20)
	kValue = abs(pageData(21))
	kUp = pageData(22)
	dValue = abs(pageData(23))
	dUp = pageData(24)
	jValue = abs(pageData(25))
	jUp = pageData(26)
	midBoll = abs(pageData(27))
	midBollUp = pageData(28)
	upperBoll = abs(pageData(29))
	upperBollUp = pageData(30)
	lowerBoll = abs(pageData(31))
	lowerBollUp = pageData(32)
	cjl = abs(pageData(33))
	cjlUp = pageData(34)
	cjl5 = abs(pageData(35))
	cjl5Up = pageData(36)
	cjl10 = abs(pageData(37))
	cjl10Up = pageData(38)
	
	/*
	testPageStr = getTestPrice()//测试专用
	testStr = Split(testPageStr, "|")
	Dim testData(25)
	If UBound(testStr) >= 0 Then 
		For i = 0 To UBound(testStr)
			If testStr(i) = "-1"  Then 
				testData(i) = - 1 
				Plugin.Msg.Tips "临时数据采集错误：" & i& "数据=" & testPageStr
			else
				testData(i) = Cdbl(testStr(i))
			End if
		Next
	End If
	//MessageBox "testStr=" & testStr
	
	newPrice = testData(0)
	openPrice = testData(1)
	highPrice = testData(2)
	lowerPrice = testData(3)
	cjl = testData(4)
	*/
	dataStr = "" 
	dataStr = dataStr & ma5 & "|"
	dataStr = dataStr & ma5Up & "|"
	dataStr = dataStr & ma11 & "|"
	dataStr = dataStr & ma11Up & "|"
	dataStr = dataStr & ma31 & "|"
	dataStr = dataStr & ma31Up & "|"
	dataStr = dataStr & ma63 & "|"
	dataStr = dataStr & ma63Up & "|"
	dataStr = dataStr & ma250 & "|"
	dataStr = dataStr & ma250Up & "|"
	dataStr = dataStr & newPrice & "|"
	dataStr = dataStr & openPrice & "|" 
	dataStr = dataStr & highPrice & "|" 
	dataStr = dataStr & lowerPrice & "|" 
	dataStr = dataStr & weibi & "|" 
	dataStr = dataStr & weicha & "|" 
	dataStr = dataStr & neipan & "|" 
	dataStr = dataStr & waipan & "|" 
	dataStr = dataStr & liangbi & "|" 
	dataStr = dataStr & cciValue & "|"
	dataStr = dataStr & cciUp & "|"
	dataStr = dataStr & kValue & "|"
	dataStr = dataStr & kUp & "|"
	dataStr = dataStr & dValue & "|"
	dataStr = dataStr & dUp & "|"
	dataStr = dataStr & jValue & "|"
	dataStr = dataStr & jUp & "|"
	dataStr = dataStr & midBoll & "|"
	dataStr = dataStr & midBollUp & "|"
	dataStr = dataStr & upperBoll & "|"
	dataStr = dataStr & upperBollUp & "|"
	dataStr = dataStr & lowerBoll & "|"
	dataStr = dataStr & lowerBollUp & "|" 
	dataStr = dataStr & cjl & "|"
	dataStr = dataStr & cjlUp & "|"
	dataStr = dataStr & cjl5 & "|"
	dataStr = dataStr & cjl5Up & "|"
	dataStr = dataStr & cjl10 & "|"
	dataStr = dataStr & cjl10Up

	//MessageBox "dataStr=" & dataStr
	cfgDataKey = "data_" & periodNo
	cfgStatusKey = "status_" & periodNo
		
	oldPageDataStr = Plugin.File.ReadINI(stockCode, cfgDataKey, "d:\gupiao\data\Config.ini")
	//MessageBox "oldPageDataStr=" & oldPageDataStr
	If oldPageDataStr <> "" Then 
		oldPageStr = Split(oldPageDataStr, "|")
		Dim oldPageData(40)
		If UBound(oldPageStr) >= 0 Then 
			For i = 0 To UBound(oldPageStr)
				If oldPageStr(i) = "-1" Then 
					If mid(stockCode, 1, 1) = "2" Then 
						If i = 16 or i = 17  Then 
							oldPageData(i) = - 1 
						ElseIf periodNo = "5分钟" or periodNo = "30分钟" or periodNo = "60分钟" Then
							If i = 8 Then 
								oldPageData(i) = - 1 
							Else 
								oldPageData(i) = - 1 
								Plugin.Msg.Tips "上次数据采集错误：" & i & "数据=" &oldPageDataStr
							End If
						End If
					End If
				else
					oldPageData(i) = Cdbl(oldPageStr(i))
				End If
			Next
		End If
		oldma5 = oldPageData(0)
		oldma5Up = oldPageData(1)
		oldma11 = oldPageData(2)
		oldma11Up = oldPageData(3)
		oldma31 = oldPageData(4)
		oldma31Up = oldPageData(5)
		oldma63 = oldPageData(6)
		oldma63Up = oldPageData(7)
		oldma250 = oldPageData(8)
		oldma250Up = oldPageData(9)
		oldNewPrice = oldPageData(10)
		oldOpenPrice = oldPageData(11)
		oldHighPrice = oldPageData(12)
		oldLowerPrice = oldPageData(13)
		oldWeibi = oldPageData(14)
		oldWeicha = oldPageData(15)
		oldWaipan = oldPageData(16)
		oldNeipan = oldPageData(17)
		oldLiangbi = oldPageData(18)
		oldCciValue = oldPageData(19)
		oldCciUp = oldPageData(20)
		oldKvalue = oldPageData(21)
		oldKUp = oldPageData(22)
		oldDvalue = oldPageData(23)
		oldDUp = oldPageData(24)
		oldJvalue = oldPageData(25)
		oldJUp = oldPageData(26)
		oldMidBoll = oldPageData(27)
		oldMidBollUp = oldPageData(28)
		oldUpperBoll = oldPageData(29)
		oldUpperBollUp = oldPageData(30)
		oldLowerBoll = oldPageData(31)
		oldLowerBollUp = oldPageData(32)
		oldCJL = oldPageData(33)
		oldCJLUp = oldPageData(34)
		oldCJL5 = oldPageData(35)
		oldCJL5Up = oldPageData(36)
		oldCJL10 = oldPageData(37)
		oldCJL10Up = oldPageData(38)
	End If
	
	/* 获得股票当前位置
		1   底部区域
		2	KDJ-J=0
		3   寻找买点
		4 	买入股票
		4	寻找卖点
		5	卖出股票
	*/

	stockPositionArea = Plugin.File.ReadINI(stockCode, "positionArea", "d:\gupiao\data\Config.ini")
	Select Case stockPositionArea
		Case 1
			//进入底部区域
			MessageBox "进入底部区域"

		Case 2
			//kdj进入底部区域
			If dataJ = 0 Then 
				stockPositionArea = 3
			End If
		Case 3
			//寻找买点
			If dataCCI > - 100  Then 
				If dataNewPrice > dataMa5 Then 
					stockPositionArea = 4
				Else 
					diffValue = abs(dataNewPrice - dataMa5) / dataNewPrice
					If diffValue <= 0.001 Then 
						stockPositionArea = 4
					Else 
						stockPositionArea = -1
					End If
				End If
			End If
		Case 4
			//买入股票
			MessageBox "买入" & newPrice & now
			Call Plugin.File.WriteINI(stockCode, "buyPrice", newPrice, "d:\gupiao\data\Config.ini")
			Call Plugin.File.WriteINI(stockCode, "buyTime",Now , "d:\gupiao\data\Config.ini")
			Call Plugin.File.WriteINI(stockCode, "buyCount", 100, "d:\gupiao\data\Config.ini")
				
		Case 5
			//持股
			If dataNewPrice >= dataMa5 Then 
				stockPositionArea = 5
			Else
				diffValue = abs(dataNewPrice - dataMa5) / dataNewPrice
				If diffValue <= 0.001 Then 
					stockPositionArea = 5
				Else 
					stockPositionArea = 6
				End If
			End If
		Case 6
			//卖出股票
			Call sellStock(stockCode, dataNewPrice, 100)
			stockPositionArea = -1
			
		Case Else
			
			//寻找底部区域
			If cciValue < - 200 Then 
				stockPositionArea = 1
			End If
	End Select
	Call Plugin.File.WriteINI(stockCode, "positionArea", stockPositionArea, "d:\gupiao\data\Config.ini")

	/*
	newPrice = getNewPrice()
		
	old_cci = Plugin.File.ReadINI(stockCode, "cci", "d:\gupiao\data\Config.ini")
	old_cciDirection = Plugin.File.ReadINI(stockCode, "cciDirection", "d:\gupiao\data\Config.ini")
	cciBuyStatus = - 1 
	cciSellStatus = -1
	If old_cci <> "" and cciValue <> "" Then 
		o_cci = CDbl(old_cci)
		n_cci = CDbl(cciValue)
		If o_cci <= - 100  and n_cci > - 100  Then 
			cciBuyStatus = 1
		End If
		If o_cci <= 100 and n_cci > 100 Then 
			cciBuyStatus = 2
		End If
		If o_cci > - 100  and n_cci < - 100  Then 
			cciSellStatus = 1
		End If
		If o_cci > 100 and n_cci < 100 Then 
			cciSellStatus = 2
		End If
	End If
	
	maBuyStatus = - 1 
	maSellStatus = -1
	If ma5 <> - 1  and ma10 <> - 1  and ma20 <> - 1  Then 
		ma5_data = CDbl(ma5)
		ma10_data = CDbl(ma10)
		ma20_data = CDbl(ma20)
		
		If ma5_data >= ma10_data and ma10_data >= ma20_data Then 
			maBuyStatus = 1
		End If
		
		If ma5_data < ma10_data and ma10_data < ma20_data Then 
			maSellStatus = 1
		End If
	End If
	
	old_k = Plugin.File.ReadINI(stockCode, "k", "d:\gupiao\data\Config.ini")
	old_d = Plugin.File.ReadINI(stockCode, "d", "d:\gupiao\data\Config.ini")
	old_j = Plugin.File.ReadINI(stockCode, "j", "d:\gupiao\data\Config.ini")
	n_k = CDbl(kValue)
	n_d = CDbl(dValue)
	n_j = CDbl(jValue)
	o_k = CDbl(old_k)
	o_d = CDbl(old_d)
	o_j = CDbl(old_j)
	//MessageBox "newKdj=" & n_k & "," & n_d & "," & n_j & ",oldKdj=" & o_k & "," & o_d & "," & o_j
	kdjBuyStatus = - 1 
	kdjSellStatus = -1
	If o_k < o_d and n_k > n_d Then 
		kdjBuyStatus = 1
	End If
	If o_k > o_d and n_k < n_d Then 
		kdjSellStatus = 1
	End If
	
	If cciSellStatus > 0 and kdjSellStatus > 0 Then 
		MessageBox "cci卖出" & newPrice
		Call sellStock(stockCode, newPrice, 100)
	End If

	

	If maBuyStatus > 0 Then 
		old_maBuyStatus = Plugin.File.ReadINI(stockCode, "maBuyStatus", "d:\gupiao\data\Config.ini")
		//MessageBox "old_maBuyStatus=" & old_maBuyStatus
		If old_maBuyStatus = "-1"  Then 
			MessageBox "买入" & newPrice
			Call buyStock(stockCode, newPrice, 100)
		End If
	End If
	*/
	
	Call Plugin.File.WriteINI(stockCode, "pageData", pageDataStr, "d:\gupiao\data\Config.ini")
End Function


Function buyStock(stockCode, price, count)
	yongjin = price * count * 3 / 2000
	If yongjin < 5 Then 
		yongjin = 5.00
	End If
	nPrice = (price * count + yongjin) / count
	
	Call Plugin.File.WriteINI(stockCode, "buyPrice", nPrice, "d:\gupiao\data\Config.ini")
	Call Plugin.File.WriteINI(stockCode, "buyTime",Now , "d:\gupiao\data\Config.ini")
	Call Plugin.File.WriteINI(stockCode, "buyCount", buyCount, "d:\gupiao\data\Config.ini")
End Function


Function sellStock(stockCode, price, count)
	buyPrice = Plugin.File.ReadINI(stockCode, "buyPrice", "d:\gupiao\data\Config.ini")
	buyTime = Plugin.File.ReadINI(stockCode, "buyTime", "d:\gupiao\data\Config.ini")
	
	yongjin = price * count * 3 / 2000
	If yongjin < 5 Then 
		yongjin = 5.00
	End If
	yinhuashui = price * count / 1000
	nPrice = (price * count - yongjin - yinhuashui) / count
	lirun = (nPrice - buyPrice) * count
	MessageBox stockCode & "利润=" & lirun
	Call Plugin.File.WriteINI(stockCode, "sellPrice", nPrice, "d:\gupiao\data\Config.ini")
	Call Plugin.File.WriteINI(stockCode, "sellTime",Now , "d:\gupiao\data\Config.ini")
	Call Plugin.File.WriteINI(stockCode, "sellCount", buyCount, "d:\gupiao\data\Config.ini")
	
End Function

Function getNewPrice()
	priceStr = Lib.同花顺.getThsDataItem(24,280,84,303,"00e600-000000|ff3232-000000|ffffff-000000")
	getNewPrice = priceStr
End Function


/*
功能：判断两条线是否金叉 
 返回：true 成功金叉
       False 失败金叉
*/
Function glodCross(olda, a, oldb, b)
	If a >= olda and b <= oldb Then 
		//形成金叉
		
	End If
	
End Function

/* 
	功能：判断2条均线是否死叉 
	返回：true 成功死叉
	      False 失败死叉
*/
Function dieCross(olda, a, oldb, b)
	If olda >= a and b >= oldb Then 
		//形成死叉
		
	End If
End Function

Function keyKLineData
	
End Function
