[General]
SyntaxVersion=2
MacroID=8fd0565d-fbbf-44b5-9833-0348c9d26475
[Comment]

[Script]
//��������д�������ӳ������
//д�걣�������һ������ϵ���Ҽ���ѡ��ˢ�¡�����
/*
�����Ʊ 
����ֵ�� 
	1		�ύ�ɹ�
	-1		û��������Ʊ���
	-2		û��ȡ�����������
	-3		���������Ϊ��
	-4		��������С��1��
	-5		û��ȡ�ù�Ʊ����۸�
*/
Function buyStock(stockCode)
	T1 = 1000
	For n= 1 to 3
		tdzHwnd = Lib.���ú���.getHwnd("ͨ�������Ͻ���")
		If tdzHwnd > 0 Then 
			Exit For
		End If
		Plugin.Msg.Tips "��������Ʊ�������"
		Delay 3000
	Next
	
	If tdxHwnd < 0 Then 
		buyStock = -1 
		//����ֲ�
		Call Lib.���ú���.MouseClick(415,50)
		Exit Function
	End If
	
	//�������˵�
	Call Lib.���ú���.MouseClick(241,57)
	Delay T1
	
	MoveTo 291, 110
	HwndStockCode = Plugin.Window.MousePoint()
	Call Plugin.Window.SendString(HwndStockCode, stockCode)	//�����Ʊ����
	Delay T1
	//MessageBox "input stock count:" & stockCode & "," & Hwnd_B1
	
	//����������Ʊ������
	MoveTo 348,175
	HwndMaxCount = Plugin.Window.MousePoint()
	txtMaxCount = Plugin.Window.GetTextEx(HwndMaxCount, 1)
	If txtMaxCount = "" Then 
		buyStock = -2 
		//����ֲ�
		Call Lib.���ú���.MouseClick(415,50)
		Exit Function
	End If
	
	If txtMaxCount = "" Then 
		buyStock = -3 
		//����ֲ�
		Call Lib.���ú���.MouseClick(415,50)
		Exit Function
	End If

	maxCount = CDbl(txtMaxCount)
	//MessageBox "MaxCount=" & maxCount	
	If maxCount < 100 Then 
		//����ֲ�
		Call Lib.���ú���.MouseClick(415,50)
		buyStock = -4 
		Exit Function
	End If

	
	//�������ļ۸�
	MoveTo 321, 135
	HwndPrice = Plugin.Window.MousePoint()
	txtPrice = Plugin.Window.GetTextEx(HwndPrice, 1)
	//MessageBox "�۸�:" & txtPrice

	//������������
	If txtPrice = "" Then 
		buyStock = -5
		//����ֲ�
		Call Lib.���ú���.MouseClick(415,50)
		Exit Function
	End If
	
	price = CDbl(txtPrice)
	count = int(5000 / price / 100)
	If count <= 0 Then 
		count = 1
	End If
	count = count * 100
	
	//���빺������
	Call Lib.���ú���.MouseClick(307, 205)
	Delay T1
	HwndCount = Plugin.Window.MousePoint()
	//MessageBox "����ָ�룺" & HwndCount
	Call Plugin.Window.SendString(HwndCount, count)
	Delay T1
	
	//����������µ�"
	Call Lib.���ú���.MouseClick(373,232)
	Delay T1

	// ����ȷ��--------//  ����ȡ��ȷ�����裬�ڽ�����������ã�����
	Hwnd_B4 = Plugin.Window.Find("#32770", "���뽻��ȷ��")
	//Hwnd_B41 = Plugin.Window.FindEx(Hwnd_B4, 0, "Button", "ȡ��")
	Hwnd_B41 = Plugin.Window.FindEx(Hwnd_B4, 0, "Button", "����ȷ��")
	Delay T1
	Call Plugin.Bkgnd.LeftClick(Hwnd_B41, 40, 10)
	// ����ȷ��------
		
	Delay T1
	Hwnd_B7 = Plugin.Window.Find("#32770", "��ʾ") 
	Hwnd_B71 = Plugin.Window.FindEx(Hwnd_B7, 0, "Button", "ȷ��")
	//Delay T1
	Call Plugin.Bkgnd.LeftClick(Hwnd_B71, 10, 10) //ȷ����ʾ��Ϣ
	Delay T1
	
	buyStock = 1
	dispStr = "�ɹ�����" & stockCode & " ������" & count & " �۸�" & txtPrice 
	Plugin.Msg.Tips dispStr
	TracePrint dispStr
	
	//����ֲ�
	Call Lib.���ú���.MouseClick(415,50)
End Function


/*
������Ʊ 
����ֵ�� 
	1		�ύ�ɹ�
	-1		û��������Ʊ���
	-2		û��ȡ�������۸�
	-3		û�л�����������
	-4		����������С��1��		
*/
Function sellStock(stockCode)
	T1 = 1000
	tdxHwnd = getTdxHwnd()
	
	If tdxHwnd < 0 Then 
		sellStock = -1 
		//����ֲ�
		Call Lib.���ú���.MouseClick(415,50)
		Exit Function
	End If
	
	//�������˵�
	Call Lib.���ú���.MouseClick(281,58)
	Delay T1
	
	MoveTo 291, 110
	HwndStockCode = Plugin.Window.MousePoint()
	Call Plugin.Window.SendString(HwndStockCode, stockCode)	//�����Ʊ����
	Delay T1
	//MessageBox "input stock count:" & stockCode & "," & Hwnd_B1

	//��������ļ۸�
	MoveTo 306,136
	HwndPrice = Plugin.Window.MousePoint()
	txtPrice = Plugin.Window.GetTextEx(HwndPrice, 1)
	//Plugin.Msg.Tips "�۸�:" & txtPrice

	//��������۸�Ϊ��
	If txtPrice = "" Then 
		sellStock = -2
		//����ֲ�
		Call Lib.���ú���.MouseClick(415,50)
		Exit Function
	End If
	
	//���������������
	MoveTo 308,156
	HwndMaxCount = Plugin.Window.MousePoint()
	txtMaxCount = Plugin.Window.GetTextEx(HwndMaxCount, 1)
	If txtMaxCount = "" Then 
		sellStock = -3 
		//����ֲ�
		Call Lib.���ú���.MouseClick(415,50)
		Exit Function
	End If
	//Plugin.Msg.Tips "����������:" & txtMaxCount
	
	count = CDbl(txtMaxCount)
	//����������С��1��
	If count < 100 Then 
		sellStock = -4 
		//����ֲ�
		Call Lib.���ú���.MouseClick(415,50)
		Exit Function
	End If
	
	//���빺������
	Call Lib.���ú���.MouseClick(312,192)
	Delay T1
	HwndCount = Plugin.Window.MousePoint()
	//MessageBox "����ָ�룺" & HwndCount
	Call Plugin.Window.SendString(HwndCount, count)
	Delay T1
	
	//����������µ�"
	Call Lib.���ú���.MouseClick(376,220)
	Delay T1
	
	// ����ȷ��--------//  ����ȡ��ȷ�����裬�ڽ�����������ã�����
	Hwnd_B4 = Plugin.Window.Find("#32770", "��������ȷ��")
	//Hwnd_B41 = Plugin.Window.FindEx(Hwnd_B4, 0, "Button", "ȡ��")
	Hwnd_B41 = Plugin.Window.FindEx(Hwnd_B4, 0, "Button", "����ȷ��")
	//Delay T1
	Call Plugin.Bkgnd.LeftClick(Hwnd_B41, 40, 10)
	// ����ȷ��------
		
	Delay T1
	Hwnd_B7 = Plugin.Window.Find("#32770", "��ʾ") 
	Hwnd_B71 = Plugin.Window.FindEx(Hwnd_B7, 0, "Button", "ȷ��")
	//Delay T1
	Call Plugin.Bkgnd.LeftClick(Hwnd_B71, 40, 10) //ȷ����ʾ��Ϣ
	Delay T1

	sellStock = 1
	dispStr = "�ɹ�����" & stockCode & " ������" & count & " �۸�" & txtPrice 
	Plugin.Msg.Tips dispStr
	TracePrint dispStr
	
	//����ֲ�
	Call Lib.���ú���.MouseClick(415,50)
End Function

Function exportToChiCang()
	Call clickChiCangButton()
	//������
	Call Lib.���ú���.MouseClick(1329,143)
	Delay 1000
	//������ļ�d:\gupiao\data\weituo.xls
	Call outputToExcel("d:\gupiao\data\chiCang.xls")
	Delay 3000
	Call Lib.���ú���.MouseClick(600,426)
	Delay 3000

End Function

Function genChiCangTxt
 	chiCangFile = "d:\gupiao\data\chiCang.xls"
	Call Lib.���ú���.ɾ���ļ�(chiCangFile)
	Call exportToChiCang()
	If Plugin.File.IsFileExist(chiCangFile) Then 
		Call Plugin.Office.OpenXls(chiCangFile)
		cdFd = Plugin.File.OpenFile("d:\gupiao\data\basic\�ֲ�.txt")
		If cdFd < 0 Then 
			MessageBox "�ֲ��ļ�����ʧ��"
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
	//����ֲ�
	Call Lib.���ú���.MouseClick(363,43)
	Delay T1
End Sub

Function getTdxHwnd
	tdzHwnd = -1
	For n= 1 to 3
		tdzHwnd = Lib.���ú���.getHwnd("ͨ�������Ͻ���")
		If tdzHwnd > 0 Then 
			Exit For
		End If
		Plugin.Msg.Tips "��������Ʊ�������"
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
	 	
		//������
		Call Lib.���ú���.MouseClick(63,297)
		Delay 1000
		
		//���뿪ʼ����
		//MoveTo 289, 86
		Call inputDateItem(292,67, Year(startDate))
		Call inputDateItem(317,68, Month(startDate))
		Call inputDateItem(332,68, Day(startDate))
		//�����������
		Call inputDateItem(435,67, Year(stopDate))
		Call inputDateItem(466,69, Month(stopDate))
		Call inputDateItem(486,66, Day(stopDate))
		
		//�����ѯ
		Call Lib.���ú���.MouseClick(1273,69)
		Delay 1000
		
		//������
		Call Lib.���ú���.MouseClick(1325,68)
		Delay 1000
		
		Call outputToExcel("d:\gupiao\data\jiaogedan.xls")

	Else 
		//����֤ȯ�������δ����
		exportJiaoGeDan = -1
	End If

End Function

Function exportToWeiTuo(filename)
	Hwnd = getTdxHwnd()
	If Hwnd < 0 Then 
		//û�����й�Ʊ�������
		exportToWeiTuo = -1 
		Exit function
	End If
	
	//�������ί��
	Call Lib.���ú���.MouseClick(74,197)
	Delay 1000
	
	//������
	Call Lib.���ú���.MouseClick(1369,66)
	Delay 1000
	
	//������ļ�d:\gupiao\data\weituo.xls
	Call outputToExcel(filename)	
End Function

Function getWeiTuoStatus(stockCode, type)
 	weituoFile = "d:\gupiao\data\weituo.xls"
	Call Lib.���ú���.ɾ���ļ�(weituoFile)
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
	
	//ѡ�������excel
	Call Lib.���ú���.MouseClick(625, 413)
	Delay 1000

	//�����ļ���
	Call Lib.���ú���.MouseClick(900, 434)
	KeyPress "BackSpace", 50
	Delay 1000
	sayString(filename)
	Delay 1000
		
	//���ȷ��
	Call Lib.���ú���.MouseClick(849,518)

End Function

Function inputDateItem(x, y, num)
	Call Lib.���ú���.MouseClick(x,y)
	str = CInt(num)
	If num < 10 Then 
		str = "0" & str
	End If
	Call Lib.���ú���.InputKeyboardCodeNoEnter(str)
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
		TracePrint "ͨ�������û�����������ֹ�����"
		buyNewStock = - 1 
		Exit Function
	End If
	//������¹��깺��
	Call Lib.���ú���.MouseClick(78,672)
	Delay 1000
	
	//˫������һ֧�¹����ꡱ
	/*ToDo: �ж��Ƿ�����깺���¹�
	*/
	Call Lib.���ú���.MouseClick(542,123)
	LeftDoubleClick 1
	Delay 3000
	
	//������������
	buyCount = Lib.����֤ȯ�Ǵ���.getTextValue(359,204)
	Delay 1000
	//�������������
	����ֵ = Lib.����֤ȯ�Ǵ���.setTextValue(331,223,buyCount)
	//���ȷ��
	Call Lib.���ú���.MouseClick(320, 252)
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
	Call Lib.���ú���.MouseClick(x,y)
	Delay 500
	txtHwnd = Plugin.Window.MousePoint()
	Call Plugin.Window.SendString(txtHwnd, value)
End Function