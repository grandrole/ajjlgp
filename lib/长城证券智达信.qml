[General]
SyntaxVersion=2
MacroID=8fd0565d-fbbf-44b5-9833-0348c9d26475
[Comment]

[Script]
//请在下面写上您的子程序或函数
//写完保存后，在任一命令库上点击右键并选择“刷新”即可
/*
买入股票 
返回值： 
	1		提交成功
	-1		没有启动股票软件
	-2		没有取得最大购买数量
	-3		最大购买数量为空
	-4		购买数量小于1手
	-5		没有取得股票购买价格
*/
Function buyStock(stockCode)
	T1 = 1000
	For n= 1 to 3
		tdzHwnd = Lib.公用函数.getHwnd("通达信网上交易")
		If tdzHwnd > 0 Then 
			Exit For
		End If
		Plugin.Msg.Tips "请启动股票交易软件"
		Delay 3000
	Next
	
	If tdxHwnd < 0 Then 
		buyStock = -1 
		//点击持仓
		Call Lib.公用函数.MouseClick(415,50)
		Exit Function
	End If
	
	//点击买入菜单
	Call Lib.公用函数.MouseClick(241,57)
	Delay T1
	
	MoveTo 291, 110
	HwndStockCode = Plugin.Window.MousePoint()
	Call Plugin.Window.SendString(HwndStockCode, stockCode)	//输入股票代码
	Delay T1
	//MessageBox "input stock count:" & stockCode & "," & Hwnd_B1
	
	//获得最大可买股票的数量
	MoveTo 348,175
	HwndMaxCount = Plugin.Window.MousePoint()
	txtMaxCount = Plugin.Window.GetTextEx(HwndMaxCount, 1)
	If txtMaxCount = "" Then 
		buyStock = -2 
		//点击持仓
		Call Lib.公用函数.MouseClick(415,50)
		Exit Function
	End If
	
	If txtMaxCount = "" Then 
		buyStock = -3 
		//点击持仓
		Call Lib.公用函数.MouseClick(415,50)
		Exit Function
	End If

	maxCount = CDbl(txtMaxCount)
	//MessageBox "MaxCount=" & maxCount	
	If maxCount < 100 Then 
		//点击持仓
		Call Lib.公用函数.MouseClick(415,50)
		buyStock = -4 
		Exit Function
	End If

	
	//获得买入的价格
	MoveTo 321, 135
	HwndPrice = Plugin.Window.MousePoint()
	txtPrice = Plugin.Window.GetTextEx(HwndPrice, 1)
	//MessageBox "价格:" & txtPrice

	//计算买入数量
	If txtPrice = "" Then 
		buyStock = -5
		//点击持仓
		Call Lib.公用函数.MouseClick(415,50)
		Exit Function
	End If
	
	price = CDbl(txtPrice)
	count = int(5000 / price / 100)
	If count <= 0 Then 
		count = 1
	End If
	count = count * 100
	
	//输入购买数量
	Call Lib.公用函数.MouseClick(307, 205)
	Delay T1
	HwndCount = Plugin.Window.MousePoint()
	//MessageBox "数量指针：" & HwndCount
	Call Plugin.Window.SendString(HwndCount, count)
	Delay T1
	
	//点击“买入下单"
	Call Lib.公用函数.MouseClick(373,232)
	Delay T1

	// 交易确定--------//  可以取消确定步骤，在交易软件内设置！！！
	Hwnd_B4 = Plugin.Window.Find("#32770", "买入交易确认")
	//Hwnd_B41 = Plugin.Window.FindEx(Hwnd_B4, 0, "Button", "取消")
	Hwnd_B41 = Plugin.Window.FindEx(Hwnd_B4, 0, "Button", "买入确认")
	Delay T1
	Call Plugin.Bkgnd.LeftClick(Hwnd_B41, 40, 10)
	// 交易确定------
		
	Delay T1
	Hwnd_B7 = Plugin.Window.Find("#32770", "提示") 
	Hwnd_B71 = Plugin.Window.FindEx(Hwnd_B7, 0, "Button", "确认")
	//Delay T1
	Call Plugin.Bkgnd.LeftClick(Hwnd_B71, 10, 10) //确认提示信息
	Delay T1
	
	buyStock = 1
	dispStr = "成功购买" & stockCode & " 数量：" & count & " 价格：" & txtPrice 
	Plugin.Msg.Tips dispStr
	TracePrint dispStr
	
	//点击持仓
	Call Lib.公用函数.MouseClick(415,50)
End Function


/*
卖出股票 
返回值： 
	1		提交成功
	-1		没有启动股票软件
	-2		没有取得卖出价格
	-3		没有获得最大开卖数量
	-4		最大可卖数量小于1手		
*/
Function sellStock(stockCode)
	T1 = 1000
	tdxHwnd = getTdxHwnd()
	
	If tdxHwnd < 0 Then 
		sellStock = -1 
		//点击持仓
		Call Lib.公用函数.MouseClick(415,50)
		Exit Function
	End If
	
	//点击买出菜单
	Call Lib.公用函数.MouseClick(281,58)
	Delay T1
	
	MoveTo 291, 110
	HwndStockCode = Plugin.Window.MousePoint()
	Call Plugin.Window.SendString(HwndStockCode, stockCode)	//输入股票代码
	Delay T1
	//MessageBox "input stock count:" & stockCode & "," & Hwnd_B1

	//获得卖出的价格
	MoveTo 306,136
	HwndPrice = Plugin.Window.MousePoint()
	txtPrice = Plugin.Window.GetTextEx(HwndPrice, 1)
	//Plugin.Msg.Tips "价格:" & txtPrice

	//获得卖出价格为空
	If txtPrice = "" Then 
		sellStock = -2
		//点击持仓
		Call Lib.公用函数.MouseClick(415,50)
		Exit Function
	End If
	
	//获得最大可卖的数量
	MoveTo 308,156
	HwndMaxCount = Plugin.Window.MousePoint()
	txtMaxCount = Plugin.Window.GetTextEx(HwndMaxCount, 1)
	If txtMaxCount = "" Then 
		sellStock = -3 
		//点击持仓
		Call Lib.公用函数.MouseClick(415,50)
		Exit Function
	End If
	//Plugin.Msg.Tips "最大可卖数量:" & txtMaxCount
	
	count = CDbl(txtMaxCount)
	//最大可卖数量小于1手
	If count < 100 Then 
		sellStock = -4 
		//点击持仓
		Call Lib.公用函数.MouseClick(415,50)
		Exit Function
	End If
	
	//输入购买数量
	Call Lib.公用函数.MouseClick(312,192)
	Delay T1
	HwndCount = Plugin.Window.MousePoint()
	//MessageBox "数量指针：" & HwndCount
	Call Plugin.Window.SendString(HwndCount, count)
	Delay T1
	
	//点击“卖出下单"
	Call Lib.公用函数.MouseClick(376,220)
	Delay T1
	
	// 交易确定--------//  可以取消确定步骤，在交易软件内设置！！！
	Hwnd_B4 = Plugin.Window.Find("#32770", "卖出交易确认")
	//Hwnd_B41 = Plugin.Window.FindEx(Hwnd_B4, 0, "Button", "取消")
	Hwnd_B41 = Plugin.Window.FindEx(Hwnd_B4, 0, "Button", "卖出确认")
	//Delay T1
	Call Plugin.Bkgnd.LeftClick(Hwnd_B41, 40, 10)
	// 交易确定------
		
	Delay T1
	Hwnd_B7 = Plugin.Window.Find("#32770", "提示") 
	Hwnd_B71 = Plugin.Window.FindEx(Hwnd_B7, 0, "Button", "确认")
	//Delay T1
	Call Plugin.Bkgnd.LeftClick(Hwnd_B71, 40, 10) //确认提示信息
	Delay T1

	sellStock = 1
	dispStr = "成功卖出" & stockCode & " 数量：" & count & " 价格：" & txtPrice 
	Plugin.Msg.Tips dispStr
	TracePrint dispStr
	
	//点击持仓
	Call Lib.公用函数.MouseClick(415,50)
End Function

Function exportToChiCang()
	Call clickChiCangButton()
	//点击输出
	Call Lib.公用函数.MouseClick(1329,143)
	Delay 1000
	//输出到文件d:\gupiao\data\weituo.xls
	Call outputToExcel("d:\gupiao\data\chiCang.xls")
	Delay 3000
	Call Lib.公用函数.MouseClick(600,426)
	Delay 3000

End Function

Function genChiCangTxt
 	chiCangFile = "d:\gupiao\data\chiCang.xls"
	Call Lib.公用函数.删除文件(chiCangFile)
	Call exportToChiCang()
	If Plugin.File.IsFileExist(chiCangFile) Then 
		Call Plugin.Office.OpenXls(chiCangFile)
		cdFd = Plugin.File.OpenFile("d:\gupiao\data\basic\持仓.txt")
		If cdFd < 0 Then 
			MessageBox "持仓文件创建失败"
			Exit Function
		End If
		row = 6
		code = Plugin.Office.ReadXls(1, row, 1)
		While code <> ""
			name = Plugin.Office.ReadXls(1, row, 2)
			buyPrice = Plugin.Office.ReadXls(1, row, 6)
			buyCountStr = Plugin.Office.ReadXls(1, row, 5)
			tmpStr = code & "|" & name & "|" & buyPrice & "|" & buyCount
			If buyCountStr <> "0" Then 
				Call Plugin.File.WriteLine(cdFd,code)
			End If
			row = row + 1
			code = Plugin.Office.ReadXls(1, row, 1)
		Wend
		Call Plugin.File.CloseFile(cdFd)
		Call Plugin.Office.CloseXls()
		getChiCang = 1
	End If
	
End Function

Sub clickChiCangButton
	T1 = 1000
	//点击持仓
	Call Lib.公用函数.MouseClick(363,43)
	Delay T1
End Sub

Function getTdxHwnd
	tdzHwnd = -1
	For n= 1 to 3
		tdzHwnd = Lib.公用函数.getHwnd("通达信网上交易")
		If tdzHwnd > 0 Then 
			Exit For
		End If
		Plugin.Msg.Tips "请启动股票交易软件"
		Delay 3000
	Next
	getTdxHwnd = tdzHwnd
End Function

Function exportJiaoGeDan(startDateStr,stopDateStr)
	Hwnd = getTdxHwnd()
	If Hwnd > 0 Then 
	 	If isDate(startDateStr) Then 
			startDate = Cdate(startDateStr)
		Else 
			startDate = Date
	 	End If
	 	If isDate(stopDateStr) Then 
			stopDate = Cdate(stopDateStr)
		Else 
			stopDate = Date
	 	End If
	 	
	 	If startDate > stopDate Then 
	 		tmpDate = startDate
	 		startDate = stopDate
	 		stopDate = tmpDate
	 	End If
	 	
	 	period = DateDiff("d", startDate, stopDate)
	 	If period >= 60 Then 
	 		startDate = DateAdd("d", -59, stopDate)
	 	End If
	 	
		//点击交割单
		Call Lib.公用函数.MouseClick(63,297)
		Delay 1000
		
		//输入开始日期
		//MoveTo 289, 86
		Call inputDateItem(292,67, Year(startDate))
		Call inputDateItem(317,68, Month(startDate))
		Call inputDateItem(332,68, Day(startDate))
		//输入结束日期
		Call inputDateItem(435,67, Year(stopDate))
		Call inputDateItem(466,69, Month(stopDate))
		Call inputDateItem(486,66, Day(stopDate))
		
		//点击查询
		Call Lib.公用函数.MouseClick(1273,69)
		Delay 1000
		
		//点击输出
		Call Lib.公用函数.MouseClick(1325,68)
		Delay 1000
		
		Call outputToExcel("d:\gupiao\data\jiaogedan.xls")

	Else 
		//长城证券交易软件未启动
		exportJiaoGeDan = -1
	End If

End Function

Function exportToWeiTuo(filename)
	Hwnd = getTdxHwnd()
	If Hwnd < 0 Then 
		//没有运行股票交易软件
		exportToWeiTuo = -1 
		Exit function
	End If
	
	//点击当日委托
	Call Lib.公用函数.MouseClick(74,197)
	Delay 1000
	
	//点击输出
	Call Lib.公用函数.MouseClick(1369,66)
	Delay 1000
	
	//输出到文件d:\gupiao\data\weituo.xls
	Call outputToExcel(filename)	
End Function

Function getWeiTuoStatus(stockCode, type)
 	weituoFile = "d:\gupiao\data\weituo.xls"
	Call Lib.公用函数.删除文件(weituoFile)
	Call exportToWeiTuo(weituoFile)
	If Plugin.File.IsFileExist(weituoFile) Then 
		Call Plugin.Office.OpenXls(weituoFile)
		row = 2
		TradeDate = Plugin.Office.ReadXls(1, row, 1)
		While TradeDate <> ""
			weituoTime = Plugin.Office.ReadXls(1, row, 2)
			code = Plugin.Office.ReadXls(1, row, 3)
			name = Plugin.Office.ReadXls(1, row, 4)
			TradeType = Plugin.Office.ReadXls(1, row, 5)
			price = Plugin.Office.ReadXls(1, row, 6)
			count = Plugin.Office.ReadXls(1, row, 7)
			status = Plugin.Office.ReadXls(1, row, 12)
			MessageBox TradeDate & "," & TradeType & "," & code & "," & name & "," & price & "," & count & "," & status
			TracePrint TradeDate & "," & TradeType & "," & code & "," & name & "," & price & "," & count & "," & status
			If code = stockCode Then 
				getWeiTuoStatus = status
				Exit function
			End If
			row = row + 1
			TradeDate = Plugin.Office.ReadXls(1, row, 1)
		Wend
		Call Plugin.Office.CloseXls()
		getWeiTuoStatus = -2
	End If

End Function

Function outputToExcel(filename)
	
	//选择输出到excel
	Call Lib.公用函数.MouseClick(625, 413)
	Delay 1000

	//输入文件名
	Call Lib.公用函数.MouseClick(900, 434)
	KeyPress "BackSpace", 50
	Delay 1000
	sayString(filename)
	Delay 1000
		
	//点击确定
	Call Lib.公用函数.MouseClick(849,518)

End Function

Function inputDateItem(x, y, num)
	Call Lib.公用函数.MouseClick(x,y)
	str = CInt(num)
	If num < 10 Then 
		str = "0" & str
	End If
	Call Lib.公用函数.InputKeyboardCodeNoEnter(str)
	Delay 500
End Function

Function getDateStr(DateVal)
	monthStr = Month(DateVal)
	If monthStr < 10 Then 
		monthStr = "0" & monthStr
	End If
	dayStr = Day(DateVal)
	If dayStr < 10 Then 
		dayStr = "0" & dayStr
	End If
	
	getDateStr = Year(DateVal) & "-" & monthStr & "-" & dayStr
	MessageBox getDateStr
End Function

Sub readJiaoGeDan
	Call Plugin.Office.OpenXls("D:\gupiao\data\jgd.xls")
	row = 2
	TradeDate = Plugin.Office.ReadXls(1, row, 1)
	While TradeDate <> ""
		TradeType = Plugin.Office.ReadXls(1, row, 2)
		code = Plugin.Office.ReadXls(1, row, 3)
		name = Plugin.Office.ReadXls(1, row, 4)
		price = Plugin.Office.ReadXls(1, row, 5)
		count = Plugin.Office.ReadXls(1, row, 6)
		yongjin = Plugin.Office.ReadXls(1, row, 11)
		yinhuashui = Plugin.Office.ReadXls(1, row, 12)
		guohufei = Plugin.Office.ReadXls(1, row, 13)
		jiesuanfei = Plugin.Office.ReadXls(1, row, 14)
		fujiafei = Plugin.Office.ReadXls(1, row, 15)
		fee = yongjin + yinhuashui + guohufei + jiesuanfei + fujiafei
		MessageBox TradeDate & "," & TradeType & "," & code & "," & name & "," & price & "," & count & "," & fee
		TracePrint TradeDate & "," & TradeType & "," & code & "," & name & "," & price & "," & count & "," & fee
		row = row + 1
	Wend
	Call Plugin.Office.CloseXls()
End Sub

Function buyNewStock
	tdxHwnd = getTdxHwnd()
	If tdxHwnd < 0 Then 
		TracePrint "通达信软件没有启动，请手工启动"
		buyNewStock = - 1 
		Exit Function
	End If
	//点击“新股申购”
	Call Lib.公用函数.MouseClick(78,672)
	Delay 1000
	
	//双击“第一支新股坐标”
	/*ToDo: 判断是否存在申购的新股
	*/
	Call Lib.公用函数.MouseClick(542,123)
	LeftDoubleClick 1
	Delay 3000
	
	//获得最大购买数量
	buyCount = Lib.长城证券智达信.getTextValue(359,204)
	Delay 1000
	//设置最大购买数量
	返回值 = Lib.长城证券智达信.setTextValue(331,223,buyCount)
	//点击确定
	Call Lib.公用函数.MouseClick(320, 252)
	Delay 1000	
End Function

Function getTextValue(x, y)
	MoveTo x, y
	txtHwnd = Plugin.Window.MousePoint()
	Delay 500
	txt = Plugin.window.GetTextEx(txtHwnd, 1)
	getTextValue = txt
End Function

Function setTextValue(x, y, value)
	Call Lib.公用函数.MouseClick(x,y)
	Delay 500
	txtHwnd = Plugin.Window.MousePoint()
	Call Plugin.Window.SendString(txtHwnd, value)
End Function