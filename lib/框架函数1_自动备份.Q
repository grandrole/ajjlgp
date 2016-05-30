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
MacroID=ef71fa7c-4e63-435a-8497-dc5cb8f5b46d
Description=框架函数1_自动备份
Enable=1
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
//请在下面写上您的子程序或函数
//写完保存后，在任一命令库上点击右键并选择“刷新”即可

//同花顺每次执行90次
Sub loopProcessTHSStockFile(filename, flowStr, position)

	TracePrint "同花顺【" & filename & "】文件循环" & flowStr & "...... 开始"
	
    stockFile = Plugin.File.OpenFile(filename)
    If stockFile = -1  Then 
        MsgBox "打开" & filename & "文件失败"
        Exit Sub
    End If
   
   str = Plugin.File.ReadLine(stockFile)
    num = position
    While num > 0
    	str = Plugin.File.ReadLine(stockFile)
    	num = num - 1
    Wend
    stockCode = str
    no = 1
    globalNo = position
    
    thsHwnd = Lib.同花顺.getThsHwnd()
	If thsHwnd < 0 Then 
		TracePrint "不能正常启动同花顺股票软件，请手工启动！"
		Exit Sub
	End If
	
	Call Lib.同花顺.showSingleStockPage()
    If flowStr = "均线纠结" or flowStr = "T0高抛低吸" or flowStr = "均线多头排列" Then 
    	返回值 = Lib.同花顺.setFivePicStyle()
    End If
    
    While stockCode <> ""
        If no > 90 Then 
        	Call Plugin.File.WriteINI("ths", "loopNum", globalNo, "d:\gupiao\data\Config.ini")
        	Delay 1000
            //MessageBox no
            返回值 = Lib.同花顺.stopThs()
            Delay 6000
            
            KeyPress "Enter", 1
            
            Hwnd = -1 
            While Hwnd < 0
            	call Lib.同花顺.startThs()
				Delay 10 * 1000
				Hwnd = Lib.公用函数.getHwnd("同花顺(")
			Wend
			If Hwnd > 0 Then 
            	Call Lib.同花顺.showSingleStockPage()
            	If flowStr = "均线纠结" or flowStr = "T0高抛低吸" or flowStr = "均线多头排列" Then 
    				返回值 = Lib.同花顺.setFivePicStyle()
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
        Call processStockFunction(stockCode, flowStr,globalNo)	
        str = Plugin.File.ReadLine(stockFile)
        stockCode = Mid(str, 2)
    Wend
    Call Plugin.File.CloseFile(stockFile)
    Call Plugin.File.WriteINI("ths", "loopNum", "0", "d:\gupiao\data\Config.ini")
	TracePrint "同花顺【" & filename & "】文件循环" & flowStr & "...... 结束"
    
End Sub

//执行指定的处理流程
Function processStockFunction(stockCode, flowStr, lineNo)
	Select Case flowStr
		 Case "增加同花顺自选股票"
          	  	Call Lib.同花顺.addSelfStock(stockCode)
          Case "卖出股票"
          		//MessageBox "卖出：" & stockCode
                ret1 = Lib.长城证券智达信.sellStock(stockCode)
                Delay 1000
                If ret1 > 0 Then 
                	//提交
                	ret = Lib.公用函数.deleteStockFromFile("d:\gupiao\data\basic\持仓.txt", stockCode)
                    Call Lib.文件.插入文本内容到指定行("d:\gupiao\data\basic\自选股.txt",stockCode,1)
                Else 
                    TracePrint "卖出" & stockCode & " 错误：" & ret
               	End If
           Case "买入股票"
           		ret2 = Lib.长城证券智达信.buyStock(stockCode)
                Delay 1000
				If ret2 > 0 Then 
					ret = Lib.公用函数.deleteStockFromFile("d:\gupiao\data\basic\自选股.txt",stockCode)
					Call Lib.文件.插入文本内容到指定行("d:\gupiao\data\basic\持仓.txt",stockCode,1)
				Else 
					TracePrint "买入" & stockCode & "错误:" & ret
				End If
			Case "停牌"
				Call Lib.文件.插入文本内容到指定行("d:\gupiao\data\basic\停牌股票.txt", stockCode, 1)
				返回值 = Lib.公用函数.deleteStockFromFile("d:\gupiao\data\basic\正常股票.txt", stockCode)
			Case "是否复牌"
				
			Case "复牌"
				Call Lib.文件.插入文本内容到指定行("d:\gupiao\data\basic\正常股票.txt", stockCode, 1)
				返回值 = Lib.公用函数.deleteStockFromFile("d:\gupiao\data\basic\停牌股票.txt", stockCode)
			Case "T0高抛低吸"
				status = T0GaoPaoDiXi(stockCode)
				If status = 1 Then 
					//低吸
					Call Lib.公用函数.addStockInFilename("d:\gupiao\data\低吸.json",stockCode)
				Else 
					If status = 2 Then 
						Call Lib.公用函数.addStockInFilename("d:\gupiao\data\高抛.json",stockCode)
					End If
				End If
	End Select
End Function

Function T0GaoPaoDiXi(stockCode)
	pageDataStr = Lib.同花顺.getPageDataStr(stockCode)
	periodNo = Lib.同花顺.getPeriod()
	MessageBox "pageDataStr=" & pageDataStr
	If inStr(pageDataStr, "-1|||||-1|-1|||0.00|") Then 
		//MessageBox "停牌的股票"
		Call Lib.公用函数.addStockInFilename("d:\gupiao\data\删除股票.json",stockCode)
		T0GaoPaoDiXi = - 11 
		Exit Function
	End If
	pageStr = Split(pageDataStr, "|")
	Dim pageData(50)
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
	macd = pageData(39)
	diff = pageData(40)
	dea = pageData(41)
	
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
	dataStr = dataStr & cjl10Up & "|"
	dataStr = dataStr & macd & "|"
	dataStr = dataStr & macdUp & "|"
	dataStr = dataStr & diff & "|"
	dataStr = dataStr & diffUp & "|"
	dataStr = dataStr & dea & "|"
	dataStr = dataStr & deaUp
		

	MessageBox "dataStr=" & dataStr
	cfgDataKey = "data_" & periodNo
	cfgStatusKey = "status_" & periodNo
		
	oldPageDataStr = Plugin.File.ReadINI(stockCode, cfgDataKey, "d:\gupiao\data\Config.ini")
	//MessageBox "oldPageDataStr=" & oldPageDataStr
	If oldPageDataStr <> "" Then 
		oldPageStr = Split(oldPageDataStr, "|")
		Dim oldPageData(50)
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
		oldCdma = oldPageData(39)
		oldCdmaUp = oldPageData(40)
		oldDiff = oldPageData(41)
		oldDiffUp = oldPageData(42)
		oldDea = oldPageData(43)
		oldDeaUp = oldPageData(44)
	End If
	
	
	
	T0GaoPaoDiXi = ""
	Select Case periodNo
		Case "day"
			If ma5Up = "1" Then 
				T0GaoPaoDiXi=T0GaoPaoDiXi&"|5MA向上"
			End If
			If ma31Up = "1" Then 	
				T0GaoPaoDiXi=T0GaoPaoDiXi&"|31MA向上"
			End If
	End Select

	
	minPriceStr = Plugin.File.ReadINI(stockCode, "最低点", "d:\gupiao\data\Config.ini")
	If minPriceStr <> "" Then 
		minPrice = CDbl(minPriceStr)
	Else 
		minPrice = 0
	End If
	maxPriceStr = Plugin.File.ReadINI(stockCode, "最高点", "d:\gupiao\data\Config.ini")
	If maxPriceStr <> "" Then 
		maxPrice = CDbl(maxPriceStr)
	Else 
		maxPrice = 0
	End If
	
	boduanDiStr = Plugin.File.ReadINI(stockCode, "波段底", "d:\gupiao\data\Config.ini")
	If boduanDiStr <> "" Then 
		boduanDi = CDbl(boduanDiStr)
	Else 
		boduanDi = 0
	End If
	boduanDingStr = Plugin.File.ReadINI(stockCode, "波段顶", "d:\gupiao\data\Config.ini")
	If boduanDingStr <> "" Then 
		boduanDing = CDbl(boduanDingStr)
	Else 
		boduanDing = 0
	End If
	
	If newPrice < boduanDi Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "破前低"
		Call Plugin.File.WriteINI(stockCode, "波段底", newPrice, "d:\gupiao\data\Config.ini")		
	End If
	If newPrice < minPrice Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "创新低"
		Call Plugin.File.WriteINI(stockCode, "最低点", newPrice, "d:\gupiao\data\Config.ini")
	End If
	If newPrice > maxPrice Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "破前高"
		Call Plugin.File.WriteINI(stockCode, "坡段顶", newPrice, "d:\gupiao\data\Config.ini")
	End If
	If newPrice > maxPrice Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "创新高"
		Call Plugin.File.WriteINI(stockCode, "最高点", newPrice, "d:\gupiao\data\Config.ini")
	End If
	
	If newPrice > ma31 Then 
		//找买点
		If oldKvalue < oldDvalue and kValue > dValue Then 
			//KDJ金叉
	 		If kUp = "1" and dUp = "1" Then 
	 		 	If newPrice >= ma5 Then 
	 		 		//站上5均线日
	 		 		If cjl > cjl5 Then 
						T0GaoPaoDiXi = T0GaoPaoDiXi & "|买入股票"
	 		 		End If
	 		 	End If
	 		End If
		End If 
	Else 
		//找卖点
		If oldKvalue > oldDvalue and kValue < dValue Then 
			If kUp = "0" and dUp = "0" Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi & "|卖出股票"
	 		End If
		End If	
	End If
	
	//MessageBox "dataStr" & dataStr & "," & oldPageDataStr & "," & oldDataStr
	
	jxYaLiStr = Plugin.File.ReadINI(stockCode, "均线压力", "d:\gupiao\data\Config.ini")
	If jxYaLiStr <> "" Then 
		jxYaLi = CDbl(jxYaLiStr)
		If openPrice <= jxYaLi Then 
			If newPrice >= jxYaLi Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi & "|站上均线压力"
			ElseIf highPrice >= jxYaLi Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi & "|遇到均线压力"
			End If
			
		End If
	End If
	
	jxZhiChengStr = Plugin.File.ReadINI(stockCode, "均线支撑", "d:\gupiao\data\Config.ini")
	If jxZhiChengStr <> "" Then 
		jxZhiCheng = CDbl(jxZhiCheng)
		If openPrice >= jxZhiCheng Then 
			If newPrice <= jxZhiCheng Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi & "|跌破均线支撑"
			elseif lowerPrice <= jxZhiCheng Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi& "|遇到均线支撑"
			End If
		End If
	End If
	
	//MessageBox "均线据：" & ma5 & "," & ma11 & "," & ma31 & "," & ma63 & "," & ma250
	
	//均线支撑和压力
	jxZhiCheng = 0
	jxYaLi = 0
	Dim ma(5)
	ma(0) = ma5
	ma(1) = ma11
	ma(2) = ma31
	ma(3) = ma63
	ma(4) = ma250
	
	For i = 0 To 4
		For j = i To 4
			If ma(i) >= ma(j) Then 
				temp = ma(j)
				ma(j) = ma(i)
				ma(i) = temp
			End If
		Next
	Next
		
	//MessageBox ma(0) & "<" & ma(1) & "<" & ma(2) & "<" & ma(3) & "<" & ma(4)
	dim maIndex(5)
	For i = 0 To 4
		If ma(i) = ma5 Then 
			maIndex(i) = "ma5"
		End If
		If ma(i) = ma11 Then 
			maIndex(i) = "ma11"
		End If
		If ma(i) = ma31 Then 
			maIndex(i) = "ma31"
		End If
		If ma(i) = ma63 Then 
			maIndex(i) = "ma63"
		End If
		If ma(i) = ma250 Then 
			maIndex(i) = "ma250"
		End If
	Next
	maIndexStr = maIndex(0) & "<" & maIndex(1) & "<" & maIndex(2) & "<" & maIndex(3) & "<" & maIndex(4)
	//MessageBox maIndex(0) & "<" & maIndex(1) & "<" & maIndex(2) & "<" & maIndex(3) & "<" & maIndex(4)
	
	If newPrice > 0 Then 
		diffValue1 = abs(ma5 - ma11)/newPrice
		diffValue2 = abs(ma11 - ma31)/newPrice
		diffValue3 = abs(ma5 - ma31) / newPrice
		If diffValue1 < 0.01 and diffValue2 < 0.01 and diffValue3 < 0.01 Then 
			T0GaoPaoDiXi = T0GaoPaoDiXi & "均线纠结"
		End If
	End If
		
	If newPrice <= ma(0) Then 
		//空头排列
		priceStr = "newPrice" & "<" & maIndexStr
		jxYaLi = ma(0)
	ElseIf newPrice >= ma(4) Then 
		//多头排列
		priceStr = maIndexStr & "<" & newPrice
		jxZhiCheng = ma(4)
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|均线真多头排列"
	Else 
		For num = 1 To 4
			If newPrice >= ma(num - 1) and newPrice <= ma(num) Then 
				Exit for
			End If
		Next
		jxYali = ma(num)
		jxZhiCheng = ma(num - 1)
		
		priceStr = maIndex(0) & "<"
		For i = 1 To num - 1
			priceStr = priceStr & maIndex(i) & "<"
		Next
		priceStr = priceStr & "newPrice" & "<"
		For i = num To 4
			priceStr = priceStr & maIndex(i) & "<"
		Next
	End If
	
	//MessageBox "priceStr=" & priceStr
	Call Plugin.File.WriteINI(stockCode, "均线压力", jxYaLi, "d:\gupiao\data\Config.ini")
	Call Plugin.File.WriteINI(stockCode, "均线支撑", jxYaLi, "d:\gupiao\data\Config.ini")
	T0GaoPaoDiXi = T0GaoPaoDiXi &"|" & priceStr
	
	If ma5 >= ma11 and ma11 >= ma31 and ma31 >= ma63 and ma63 >= ma250 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|均线真多头排列"
	End If
	
	If oldNewPrice < oldMa5 and newPrice >= ma5 Then 
		T0GaoPaoDiXi = "站上5日均线"
	End if
	If oldNewPrice >= oldMa5 and newPrice < ma5 Then 
		If Lib.同花顺.isUpArrow("MA5") Then 
			T0GaoPaoDiXi = T0GaoPaoDiXi & "|跌破5日均线"
		End If
	End If
	
	If oldCciValue < 100 and cciValue >= 100 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|CCI站上100"
	End If
	
	If oldCciValue >= 100 and cciValue < 100 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|CCI跌破100"
	End If
	
	cci_100Status = false
	If oldCciValue < - 100  and cciValue >= - 100  Then 
	 	T0GaoPaoDiXi = T0GaoPaoDiXi & "|CCI站上-100"
	End If
	
	If oldCciValue >= - 100  and cciValue < - 100  Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|CCI跌破-100"
	End If
	
	If oldKvalue < oldDvalue and kValue > dValue Then 
	 	If kUp = "1" and dUp = "1" Then 
			T0GaoPaoDiXi = T0GaoPaoDiXi & "|KDJ金叉"
	 	End If
	End If
	If kValue >= 80 and dValue >= 80 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|KD处于高位"
	End If
	
	If kValue <= 20 and dValue <= 20 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|KD处于低位"
	End If
	
	//MessageBox "kValue=" & kValue & ",dValue=" & dValue & ",oldKvalue=" & oldKvalue & ",oldDvalue=" & oldDvalue
	If kValue >= oldKvalue and dValue >= oldDvalue Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|KD趋势向上"
	End If
	If kValue <= oldKvalue and dValue <= oldDvalue Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|KD趋势向下"
	End If
	If oldKvalue > oldDvalue and kValue < dValue Then 
		If kUp = "0" and dUp = "0" Then 
	 		T0GaoPaoDiXi = T0GaoPaoDiXi & "|KDJ死叉"	
		End If
	End If
	
	If newPrice <= lowerBoll and oldNewPrice >= oldLowerBoll Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|跌破BOLL下轨"
	End If
	
	If oldNewPrice <= oldLowerBoll and newPrice >= lowerBoll Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|站上BOLL下轨"
	End If
	
	
	If newPrice >= upperBoll and oldNewPrice <= oldUpperBoll Then 
		//MessageBox "newPrice=" & newPrice & ",upperBoll=" & upperBoll & ",oldNewPrice=" & oldNewPrice & ",oldUpperBoll=" & oldUpperBoll
		T0GaoPaoDiXi = T0GaoPaoDiXi & "| 站上BOLL上轨"
	End If
		
	If oldNewPrice >= oldUpperBoll and newPrice <= upperBoll Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|跌破BOLL上轨"
	End If
	
	If oldNewPrice <= oldLowerBoll and newPrice >= lowerBoll Then 
			T0GaoPaoDiXi = T0GaoPaoDiXi & "|站上BOLL下轨"
	End If
		
	If oldNewPrice <= oldMidBoll and newPrice >= midBoll Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|站上BOLL中轨"
	End If
	
	If oldNewPrice >= oldMidBoll and newPrice <= midBoll Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|跌破BOLL中轨"
		//MessageBox "跌破boll中轨：" & dataStr & oldDataStr
	End If
	

	If cjl5 > cjl10 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|量能多头排列"
	End If
	If cjl > cjl5 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|大于5日成交量"
	End If
	
	If cjl > cjl10 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|大于10日成交量"
	End If
	Call Plugin.File.WriteINI(stockCode, cfgDataKey , dataStr, "d:\gupiao\data\Config.ini")
	Call Plugin.File.WriteINI(stockCode, cfgStatusKey, T0GaoPaoDiXi, "d:\gupiao\data\Config.ini")
End Function

Function getTestPrice()
	Delay 300
	newPrice = Lib.同花顺.getThsDataItem(24,280,84,303,"00e600-000000|ff3232-000000|ffffff-000000")
	Delay 300
	openPrice = Lib.同花顺.getThsDataItem(30,187,73,204,"00e600-000000|ff3232-000000|ffffff-000000")
	Delay 300
	highPrice = Lib.同花顺.getThsDataItem(30,220,77,235,"00e600-000000|ff3232-000000|ffffff-000000")
	Delay 300
	lowerPrice = Lib.同花顺.getThsDataItem(32,250,76,269,"00e600-000000|ff3232-000000|ffffff-000000")
	Delay 300
	cjl = Lib.同花顺.getThsDataItem(28, 383, 81, 400, "02e2f4-000000")
	//MessageBox "cjl=" & cjl
	getTestPrice = newPrice & "|" & openPrice & "|" & highPrice & "|" & lowerPrice & "|" & cjl
End Function