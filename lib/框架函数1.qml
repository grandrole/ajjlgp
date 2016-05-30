[General]
SyntaxVersion=2
MacroID=ef71fa7c-4e63-435a-8497-dc5cb8f5b46d
[Comment]

[Script]
//��������д�������ӳ������
//д�걣�������һ������ϵ���Ҽ���ѡ��ˢ�¡�����

//ͬ��˳ÿ��ִ��90��
Sub loopProcessTHSStockFile(filename,flowStr, periodNo)	
	maxTimesPerLoop = 100
	
	TracePrint "ͬ��˳��" & filename & "���ļ�ѭ��" & flowStr & "...... ��ʼ"

    thsHwnd = getThsHwnd()
	If thsHwnd < 0 Then 
		msgStr = "������������ͬ��˳��Ʊ��������ֹ�������"
		showErrorMsg(msgStr)
		Exit Sub
	End If
		
    stockFile = Plugin.File.OpenFile(filename)
    If stockFile = - 1  Then 
    	msgStr = "��" & filename & "�ļ�ʧ��"
		showErrorMsg(msgStr)
        Exit Sub
    End If
   
   	stockCode = Plugin.File.ReadLine(stockFile)
   	positionStr = Plugin.File.ReadINI("ths", "loopNum", "d:\gupiao\data\Config.ini")
	If positionStr <> "" Then 
		position = Cdbl(positionStr)
   		If position > 1 Then 
    		num = position
    		While num > 0
    			stockCode = Plugin.File.ReadLine(stockFile)
    			num = num - 1
    		Wend   	
   		End If
	End If
	//MessageBox "position=" & position
	
  	Call setAddtionalPicStyle()
	If InStr(flowStr, "ENE") Then 
   		setMainPicStyle ("ENEStyle")
   		//MessageBox "setENE Style"
   	Else 
   		setMainPicStyle("KLineStyle")
  	End If
  	����ֵ = Lib.ͬ��˳.setPeriod(periodNo)

  	
    no = 1
    globalNo = position
    
    While stockCode <> ""
    //MessageBox "stockCode=" & stockCode
        If no > maxTimesPerLoop Then 
        //MessageBox "loop maxTimesPerLoop "
        	Call Plugin.File.WriteINI("ths", "loopNum", globalNo, "d:\gupiao\data\Config.ini")
        	Call stopThs()
        	Delay 1000
			Call thsStopImage()
        	Call startThs()
        	thsHwnd1 = getThsHwnd()
        	//MessageBox "thsHwnd1=" & thsHwnd1
			If thsHwnd1 > 0 Then 
   				Call setAddtionalPicStyle()
   				If InStr(flowStr, "ENE") Then 
   					setMainPicStyle ("ENEStyle")
   				Else 
   					setMainPicStyle("KLineStyle")
   				End If
   				����ֵ = Lib.ͬ��˳.setPeriod(periodNo)

				no = 1
			Else 
				msgStr = "���� ѭ������ͬ��˳ʧ��"
				showErrorMsg (msgStr)
				MessageBox msgStr
            End if
        Else 
            no = no + 1
        End If
        globalNo = globalNo + 1
        Call Lib.���ú���.InputKeyboardCode(stockCode)
        //����ƶ�������������
        Call Lib.���ú���.MouseClick(1173,267)
		Delay 1000
		
		TracePrint "����" & stockCode & ",����:" & flowStr
		call processStockFunction(stockCode, "ENE�¹�")

        str = Plugin.File.ReadLine(stockFile)
        stockCode = Mid(str, 2)
    Wend
    Call Plugin.File.CloseFile(stockFile)
    Call Plugin.File.WriteINI("ths", "loopNum", "0", "d:\gupiao\data\Config.ini")
	TracePrint "ͬ��˳��" & filename & "���ļ�ѭ��" & flowStr & "...... ����"    
End Sub

Function processStockFunction(stockCode, flowStr)
	Select Case flowStr
		Case "ֹ������"
			//�ڶ����
			//�ɼ۵��ƽ��۵�5%,ֹ���˳�
		Case "ENE����"
			/*
			��������
			1. ���3���
			2. �������ߵ���ENE�й�
			3. 5�����߲��Ƕ�ͷ����
			*/
		Case "ENE����"
			/*
			�������� 
			1. ENE�����¹�
			2. 5���Ӷ�ͷ����
			*/
			kLineStatus = getKLineStatus()
			If InStr(kLineStatus, "���߶�ͷ����") Then 
				Call Lib.���ú���.addStockInFilename("d:\gupiao\data\���߶�ͷ����.json",stockCode)
			End If
		Case "����ENE�¹�"
			eneStatus = getENEStatus(stockCode)
			If eneStatus <> "��Ʊͣ��" Then 
				//MessageBox eneStatus
				If InStr(eneStatus, "�ɼ���ENE�¹�����") or InStr(eneStatus, "�ɼ�ѹ�ص�ENE�¹�") or InStr(eneStatus,"�ɼ۴�ENE�¹쵯��")  Then 	
        			Call Lib.���ú���.addStockInFilename("d:\gupiao\data\ENE�¹�.json",stockCode)
				End If
			End If
        Case "����"
        	cjlStatus = getCJLStatus(stockCode)
        	If InStr(cjlStatus, "5�վ�������") and InStr(cjlStatus, "63�վ�������") Then 
        		MessageBox stockCode & "����"
        		Call Lib.���ú���.addStockInFilename("d:\gupiao\data\����.json",stockCode)
        	End If
        Case "������Ʊ"
      		//MessageBox "������" & stockCode
            ret1 = Lib.����֤ȯ�Ǵ���.sellStock(stockCode)
            Delay 1000
            If ret1 > 0 Then 
              	//�ύ
               	ret = Lib.���ú���.deleteStockFromFile("d:\gupiao\data\basic\�ֲ�.txt", stockCode)
                Call Lib.�ļ�.�����ı����ݵ�ָ����("d:\gupiao\data\basic\��ѡ��.txt",stockCode,1)
            Else 
                TracePrint "����" & stockCode & " ����" & ret
            End If
        Case "�����Ʊ"
        	ret2 = Lib.����֤ȯ�Ǵ���.buyStock(stockCode)
            Delay 1000
			If ret2 > 0 Then 
				ret = Lib.���ú���.deleteStockFromFile("d:\gupiao\data\basic\��ѡ��.txt",stockCode)
				Call Lib.�ļ�.�����ı����ݵ�ָ����("d:\gupiao\data\basic\�ֲ�.txt",stockCode,1)
			Else 
				TracePrint "����" & stockCode & "����:" & ret
			End If
		Case "ͣ��"
			Call Lib.�ļ�.�����ı����ݵ�ָ����("d:\gupiao\data\basic\ͣ�ƹ�Ʊ.txt", stockCode, 1)
			����ֵ = Lib.���ú���.deleteStockFromFile("d:\gupiao\data\basic\������Ʊ.txt", stockCode)
		Case "�Ƿ���"
			
		Case "����"
			Call Lib.�ļ�.�����ı����ݵ�ָ����("d:\gupiao\data\basic\������Ʊ.txt", stockCode, 1)
			����ֵ = Lib.���ú���.deleteStockFromFile("d:\gupiao\data\basic\ͣ�ƹ�Ʊ.txt", stockCode)

	End Select
End Function




Function getKLineStatus
	tmpData = getTradeData(stockCode)
	If tmpData = "��Ʊͣ��" Then 
		getKLineStatus = "��Ʊͣ��"
		Exit Function
	End If
	pageDataStr = getKLineData()
	pageDataStr = pageDataStr & "|" & tmpData
	//MessageBox "dataStr=" & pageDataStr
	pageStr = Split(pageDataStr, "|")
	Dim pageData(50)
	If UBound(pageStr) >= 0 Then 
		For i = 0 To UBound(pageStr)
			If pageStr(i) = "-1" Then 
				pageData(i) = - 1 
				msgStr = "���ݲɼ�����" & i & "����=" & pageDataStr
				showErrorMsg(msgError)
			else
				pageData(i) = Cdbl(pageStr(i))
			End if
		Next
	End If
	ma5 = pageData(0)
	ma5Up = pageData(1)
	ma10 = pageData(2)
	ma10Up = pageData(3)
	ma20 = pageData(4)
	ma20Up = pageData(5)
	ma30 = pageData(6)
	ma30Up = pageData(7)
	ma40 = pageData(8)
	ma40Up = pageData(9)
	ma50 = pageData(10)
	ma50Up = pageData(11)	
	newPrice = pageData(12)
	openPrice = pageData(13)
	highPrice = pageData(14)
	lowerPrice = pageData(15)
	
	If ma5 >= ma10 and ma10 >= ma20 and ma20 >= ma30 and ma30 >= ma40 and ma40 >= ma50 Then 
		If ma5Up = "1" and ma10Up = "1" and ma20Up = "1" and ma30Up = "1" and ma40Up = "1" and ma50Up = "1" Then 
			getKLineStatus = "���߶�ͷ����"
			Exit Function
		End If
	End If
	getKLineStatus = ""
	End Function
	

Function getKLineData
		 //MessageBox "get KLine Data"
		ma5 = getThsDataItem(145, 69, 255, 86, "ffffff-000000")//0
		ma10 = getThsDataItem(197,69,329,86,"ffff0b-000000") //1
		ma20 = getThsDataItem(293,69,425,86,"ff80ff-000000") //2
		ma30 = getThsDataItem(368,69,500,86,"00e600-000000") //3
		ma40 = getThsDataItem(485, 70, 614, 88, "02e2f4-000000")//4
		ma50 = getThsDataItem(549,67,660,86,"ffffb9-000000")//4

		ma5Up = isUpArrow(145, 69, 255, 86)
		ma10Up = isUpArrow(197,69,329,86)
		ma20Up = isUpArrow(293,69,425,86)
		ma30Up = isUpArrow(368,69,500,86)
		ma40Up = isUpArrow(485, 70, 614, 88)
		ma50Up = isUpArrow(549,67,660,86)
		
		getKLineData = ma5 & "|" & ma5Up & "|" & ma10 & "|" & ma10Up & "|"  & ma20 & "|" & ma20Up & "|" & ma30 & "|" & ma30Up & "|" &ma40 & "|" & ma40Up
		/*
		KLineData = "{" & vbCrLf
		KLineData = KLineData & vbTab & "ma5:" & ma5 & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma5Up:" & ma5Up & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma10:" & ma10 & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma10Up:" & ma10Up & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma20:" & ma20 & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma20Up:" & ma20Up & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma30:" & ma30 & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma30Up:" & ma30Up & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma40:" & ma40 & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma40Up:" & ma40Up & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma50:" & ma50 & "," & vbCrLf
		KLineData = KLineData & vbTab & "ma50Up:" & ma50Up & vbCrLf
		KLineData = KLineData & "}"
		
		MessageBox "KLineData=" & KLineData
		*/
End Function

Function getEneStatus(stockCode)
	tmpData = getTradeData(stockCode)
	If tmpData = "��Ʊͣ��" Then 
		getEneStatus = "��Ʊͣ��"
		Exit Function
	End If
	pageDataStr = getEneData()
	pageDataStr = pageDataStr & "|" & tmpData
	//MessageBox "dataStr=" & pageDataStr
	pageStr = Split(pageDataStr, "|")
	Dim pageData(50)
	If UBound(pageStr) >= 0 Then 
		For i = 0 To UBound(pageStr)
			If pageStr(i) = "-1" Then 
				pageData(i) = - 1 
				msgStr = "���ݲɼ�����" & i & "����=" & pageDataStr
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
	
	If highPrice <= lowerEne Then 
		getEneStatus = "�ɼ���ENE�¹�����"
	ElseIf newPrice <= lowerEne Then
		getEneStatus = "�ɼ�ѹ�ص�ENE�¹�"
	ElseIf lowerPrice <= lowerEne Then
		getEneStatus = "�ɼ۴�ENE�¹쵯��"
	ElseIf highPrice <= ene Then
		getEneStatus = "�ɼ���ENE�¹���й�֮��"
	ElseIf newPrice <= ene Then
		getEneStatus = "�ɼ۱�ѹ�ص�ENE�й�"
	ElseIf lowerPrice <= ene Then
		getEneStatus = "�ɼ۴�ENE�й쵯��"
	ElseIf highPrice <= upperEne Then
		getEneStatus = "�ɼ���ENE�й���Ϲ�֮��"
	ElseIf newPrice <= upperEne Then
		getEneStatus = "�ɼ�ѹ�ص�ENE�Ϲ�"
	ElseIf lowerPrice > upperEne Then
		getEneStatus = "�ɼ���ENE�Ϲ쵯��"
	Else 
		getEneStatus="״̬δ֪�����˹�ȷ��"
	End If
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

Function getTradeData(stockCode)
	newPrice = Lib.ͬ��˳.getThsDataItem(1163,331,1249,349,"00e600-000000|ff3232-000000|ffffff-000000")	//6
	openPrice = Lib.ͬ��˳.getThsDataItem(1283,331,1368,346,"00e600-000000|ff3232-000000|ffffff-000000")	//7
	highPrice = Lib.ͬ��˳.getThsDataItem(1284,351,1369,366,"00e600-000000|ff3232-000000|ffffff-000000")	//8
	lowerPrice = Lib.ͬ��˳.getThsDataItem(1282, 368, 1367, 383, "00e600-000000|ff3232-000000|ffffff-000000")//9
	weibi = Lib.ͬ��˳.getThsDataItem(1161, 96, 1267, 113, "ff3232-000000|00e600-000000")
	
	weicha = Lib.ͬ��˳.getThsDataItem(1261,96,1367,113, "ff3232-000000|00e600-000000")
	waipan = Lib.ͬ��˳.getThsDataItem(1178,466,1246,484, "ff3232-000000|00e600-000000")
	neipan = Lib.ͬ��˳.getThsDataItem(1296,466,1364,484, "ff3232-000000|00e600-000000")
	liangbi = Lib.ͬ��˳.getThsDataItem(1298, 381, 1366, 399, "ff3232-000000|00e600-000000")
	
	curDate = Year(Date)
	curTime = Hour(Time) & Minute(Time)
	
	testPageStr = getTestPrice()//����ר��
	//MessageBox "testPageStr=" & testPageStr
	//TracePrint "testPageStr=" & testPageStr
	If InStr(testPageStr, "-1|-1|-1") = false Then 
		testStr = Split(testPageStr, "|")
		Dim testData(25)
		If UBound(testStr) >= 0 Then 
			For i = 0 To UBound(testStr)
				If testStr(i) = "-1"  Then 
					testData(i) = - 1 
					//Plugin.Msg.Tips "��ʱ���ݲɼ�����" & i& "����=" & testPageStr
				Else 
					If i = 5 or i = 6 Then 
						testData(i) = testStr(i)
					Else 
						testData(i) = Cdbl(testStr(i))
					End If
				End if
			Next
		End If
		newPrice = testData(0)
		openPrice = testData(1)
		highPrice = testData(2)
		lowerPrice = testData(3)
		cjl = testData(4)
		curDate = testData(5)
		curTime = testData(6)
	End If
	
	getTradeData =  newPrice & "|" & openPrice & "|" & highPrice & "|" & lowerPrice & "|" & weibi & "|" & weicha & "|" & waipan & "|" & neipan & "|" & liangbi & "|" & curDate & "|" & curTime
	If getTradeData = "-|-|-|-|-1|-1|-|-|0.00" Then 
		//MessageBox stockCode & "Ŀǰ����ͣ��״̬"
		Call Lib.���ú���.addStockInFilename("d:\gupiao\data\ͣ�ƹ�Ʊ.json", stockCode)
		getTradeData = "��Ʊͣ��"
	End If
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

Function getCJLStatus(stockCode)
	pageDataStr = getCJLData()
	//MessageBox "dataStr=" & pageDataStr
	pageStr = Split(pageDataStr, "|")
	Dim pageData(50)
	If UBound(pageStr) >= 0 Then 
		For i = 0 To UBound(pageStr)
			If pageStr(i) = "-1" Then 
				pageData(i) = - 1 
				msgStr = "���ݲɼ�����" & i & "����=" & pageDataStr
				showErrorMsg(msgError)
			else
				pageData(i) = Cdbl(pageStr(i))
			End if
		Next
	End If
	cjl = pageData(0)
	cjlUp = pageData(1)
	cjl5 = pageData(2)
	cjl5Up = pageData(3)
	cjl63 = pageData(4)
	cjl63Up = pageData(5)
	
	getCJLData = ""
	If cjl5Up = "1" Then 
		getCJLStatus = "5�վ�������"
	End If
	If cjl63Up = "1" Then 
		getCJLStatus = getCJLStatus & "|" & "63�վ�������"
	End If	
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
	tradeStr = getTradeData(stockCode)
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

/*
Function T0GaoPaoDiXi(stockCode)
	pageDataStr = Lib.ͬ��˳.getPageDataStr(stockCode)
	periodNo = Lib.ͬ��˳.getPeriod()
	//MessageBox "pageDataStr=" & pageDataStr
	If inStr(pageDataStr, "-1|||||-1|-1|||0.00|") Then 
		//MessageBox "ͣ�ƵĹ�Ʊ"
		Call Lib.���ú���.addStockInFilename("d:\gupiao\data\ɾ����Ʊ.json",stockCode)
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
				ElseIf periodNo = "5����" or periodNo = "30����" or periodNo = "60����" Then
					If i = 8 Then 
						pageData(i) = -1
					else
						pageData(i) = - 1 
						Plugin.Msg.Tips "���ݲɼ�����" & i & "����=" & pageDataStr
						TracePrint "���ݲɼ�����" & i & "����=" & pageDataStr
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
	macdUp = pageData(40)
	diff = pageData(41)
	diffUp = pageData(42)
	dea = pageData(43)
	deaUp = pageData(44)
	
	testPageStr = getTestPrice()//����ר��
	testStr = Split(testPageStr, "|")
	Dim testData(25)
	If UBound(testStr) >= 0 Then 
		For i = 0 To UBound(testStr)
			If testStr(i) = "-1"  Then 
				testData(i) = - 1 
				Plugin.Msg.Tips "��ʱ���ݲɼ�����" & i& "����=" & testPageStr
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
	dataStr = dataStr & ma50 & "|"
	dataStr = dataStr & ma50Up & "|"
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
		

	//MessageBox "dataStr=" & dataStr
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
						ElseIf periodNo = "5����" or periodNo = "30����" or periodNo = "60����" Then
							If i = 8 Then 
								oldPageData(i) = - 1 
							Else 
								oldPageData(i) = - 1 
								Plugin.Msg.Tips "�ϴ����ݲɼ�����" & i & "����=" &oldPageDataStr
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
	

	//�ж��Ƿ�Ϊ��һ����������	
	
	T0GaoPaoDiXi = ""
	//�ɼ�
	
	//����
	Select Case periodNo
		Case "5����"
		Case "30����"
		Case "60����"
		Case "day"
			If ma31Up = 1 Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi & "|�������������"
			Else 
				T0GaoPaoDiXi=T0GaoPaoDiXi&"|���߷���������"
			End If
		Case "week"
		Case "month"
	End Select
	//KDJ
	//BOLL
	//MACD
	
	Select Case periodNo
		Case "day"
			If ma5Up = "1" Then 
				T0GaoPaoDiXi=T0GaoPaoDiXi&"|5MA����"
			End If
			If ma31Up = "1" Then 	
				T0GaoPaoDiXi=T0GaoPaoDiXi&"|31MA����"
			End If
	End Select
	
	minPriceStr = Plugin.File.ReadINI(stockCode, "��͵�", "d:\gupiao\data\Config.ini")
	If minPriceStr <> "" Then 
		minPrice = CDbl(minPriceStr)
	Else 
		minPrice = 0
	End If
	maxPriceStr = Plugin.File.ReadINI(stockCode, "��ߵ�", "d:\gupiao\data\Config.ini")
	If maxPriceStr <> "" Then 
		maxPrice = CDbl(maxPriceStr)
	Else 
		maxPrice = 0
	End If
	
	boduanDiStr = Plugin.File.ReadINI(stockCode, "���ε�", "d:\gupiao\data\Config.ini")
	If boduanDiStr <> "" Then 
		boduanDi = CDbl(boduanDiStr)
	Else 
		boduanDi = 0
	End If
	boduanDingStr = Plugin.File.ReadINI(stockCode, "���ζ�", "d:\gupiao\data\Config.ini")
	If boduanDingStr <> "" Then 
		boduanDing = CDbl(boduanDingStr)
	Else 
		boduanDing = 0
	End If
	
	If newPrice < boduanDi Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "��ǰ��"
		Call Plugin.File.WriteINI(stockCode, "���ε�", newPrice, "d:\gupiao\data\Config.ini")		
	End If
	If newPrice < minPrice Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "���µ�"
		Call Plugin.File.WriteINI(stockCode, "��͵�", newPrice, "d:\gupiao\data\Config.ini")
	End If
	If newPrice > maxPrice Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "��ǰ��"
		Call Plugin.File.WriteINI(stockCode, "�¶ζ�", newPrice, "d:\gupiao\data\Config.ini")
	End If
	If newPrice > maxPrice Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "���¸�"
		Call Plugin.File.WriteINI(stockCode, "��ߵ�", newPrice, "d:\gupiao\data\Config.ini")
	End If
	

	If newPrice > ma31 Then 
		//�����
		If oldKvalue < oldDvalue and kValue > dValue Then 
			//KDJ���
	 		If kUp = "1" and dUp = "1" Then 
	 		 	If newPrice >= ma5 Then 
	 		 		//վ��5������
	 		 		If cjl > cjl5 Then 
						T0GaoPaoDiXi = T0GaoPaoDiXi & "|�����Ʊ"
	 		 		End If
	 		 	End If
	 		End If
		End If 
	Else 
		//������
		If oldKvalue > oldDvalue and kValue < dValue Then 
			If kUp = "0" and dUp = "0" Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi & "|������Ʊ"
	 		End If
		End If	
	End If

	//MessageBox "dataStr" & dataStr & "," & oldPageDataStr & "," & oldDataStr
	
	jxYaLiStr = Plugin.File.ReadINI(stockCode, "����ѹ��", "d:\gupiao\data\Config.ini")
	If jxYaLiStr <> "" Then 
		jxYaLi = CDbl(jxYaLiStr)
		If openPrice <= jxYaLi Then 
			If newPrice >= jxYaLi Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi & "|վ�Ͼ���ѹ��"
			ElseIf highPrice >= jxYaLi Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi & "|��������ѹ��"
			End If
			
		End If
	End If
	
	jxZhiChengStr = Plugin.File.ReadINI(stockCode, "����֧��", "d:\gupiao\data\Config.ini")
	If jxZhiChengStr <> "" Then 
		jxZhiCheng = CDbl(jxZhiCheng)
		If openPrice >= jxZhiCheng Then 
			If newPrice <= jxZhiCheng Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi & "|���ƾ���֧��"
			elseif lowerPrice <= jxZhiCheng Then 
				T0GaoPaoDiXi = T0GaoPaoDiXi& "|��������֧��"
			End If
		End If
	End If
	
	//MessageBox "���߾ݣ�" & ma5 & "," & ma11 & "," & ma31 & "," & ma63 & "," & ma250
	
	//����֧�ź�ѹ��
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
			T0GaoPaoDiXi = T0GaoPaoDiXi & "���߾���"
		End If
	End If
		
	If newPrice <= ma(0) Then 
		//��ͷ����
		priceStr = "newPrice" & "<" & maIndexStr
		jxYaLi = ma(0)
	ElseIf newPrice >= ma(4) Then 
		//��ͷ����
		priceStr = maIndexStr & "<" & newPrice
		jxZhiCheng = ma(4)
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|�������ͷ����"
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
	Call Plugin.File.WriteINI(stockCode, "����ѹ��", jxYaLi, "d:\gupiao\data\Config.ini")
	Call Plugin.File.WriteINI(stockCode, "����֧��", jxYaLi, "d:\gupiao\data\Config.ini")
	T0GaoPaoDiXi = T0GaoPaoDiXi &"|" & priceStr
	
	If ma5 >= ma11 and ma11 >= ma31 and ma31 >= ma63 and ma63 >= ma250 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|�������ͷ����"
	End If
	
	If oldNewPrice < oldMa5 and newPrice >= ma5 Then 
		T0GaoPaoDiXi = "վ��5�վ���"
	End if
	If oldNewPrice >= oldMa5 and newPrice < ma5 Then 
		If Lib.ͬ��˳.isUpArrow("MA5") Then 
			T0GaoPaoDiXi = T0GaoPaoDiXi & "|����5�վ���"
		End If
	End If
	
	If oldCciValue < 100 and cciValue >= 100 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|CCIվ��100"
	End If
	
	If oldCciValue >= 100 and cciValue < 100 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|CCI����100"
	End If
	
	cci_100Status = false
	If oldCciValue < - 100  and cciValue >= - 100  Then 
	 	T0GaoPaoDiXi = T0GaoPaoDiXi & "|CCIվ��-100"
	End If
	
	If oldCciValue >= - 100  and cciValue < - 100  Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|CCI����-100"
	End If
	
	If oldKvalue < oldDvalue and kValue > dValue Then 
	 	If kUp = "1" and dUp = "1" Then 
			T0GaoPaoDiXi = T0GaoPaoDiXi & "|KDJ���"
	 	End If
	End If
	If kValue >= 80 and dValue >= 80 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|KD���ڸ�λ"
	End If
	
	If kValue <= 20 and dValue <= 20 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|KD���ڵ�λ"
	End If
	
	//MessageBox "kValue=" & kValue & ",dValue=" & dValue & ",oldKvalue=" & oldKvalue & ",oldDvalue=" & oldDvalue
	If kValue >= oldKvalue and dValue >= oldDvalue Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|KD��������"
	End If
	If kValue <= oldKvalue and dValue <= oldDvalue Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|KD��������"
	End If
	If oldKvalue > oldDvalue and kValue < dValue Then 
		If kUp = "0" and dUp = "0" Then 
	 		T0GaoPaoDiXi = T0GaoPaoDiXi & "|KDJ����"	
		End If
	End If
	
	If newPrice <= lowerBoll and oldNewPrice >= oldLowerBoll Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|����BOLL�¹�"
	End If
	
	If oldNewPrice <= oldLowerBoll and newPrice >= lowerBoll Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|վ��BOLL�¹�"
	End If
	
	
	If newPrice >= upperBoll and oldNewPrice <= oldUpperBoll Then 
		//MessageBox "newPrice=" & newPrice & ",upperBoll=" & upperBoll & ",oldNewPrice=" & oldNewPrice & ",oldUpperBoll=" & oldUpperBoll
		T0GaoPaoDiXi = T0GaoPaoDiXi & "| վ��BOLL�Ϲ�"
	End If
		
	If oldNewPrice >= oldUpperBoll and newPrice <= upperBoll Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|����BOLL�Ϲ�"
	End If
	
	If oldNewPrice <= oldLowerBoll and newPrice >= lowerBoll Then 
			T0GaoPaoDiXi = T0GaoPaoDiXi & "|վ��BOLL�¹�"
	End If
		
	If oldNewPrice <= oldMidBoll and newPrice >= midBoll Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|վ��BOLL�й�"
	End If
	
	If oldNewPrice >= oldMidBoll and newPrice <= midBoll Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|����BOLL�й�"
		//MessageBox "����boll�й죺" & dataStr & oldDataStr
	End If
	

	If cjl5 > cjl10 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|���ܶ�ͷ����"
	End If
	If cjl > cjl5 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|����5�ճɽ���"
	End If
	
	If cjl > cjl10 Then 
		T0GaoPaoDiXi = T0GaoPaoDiXi & "|����10�ճɽ���"
	End If
	Call Plugin.File.WriteINI(stockCode, cfgDataKey , dataStr, "d:\gupiao\data\Config.ini")
	Call Plugin.File.WriteINI(stockCode, cfgStatusKey, T0GaoPaoDiXi, "d:\gupiao\data\Config.ini")
End Function
*/
Function getTestPrice()
	newPrice = Lib.ͬ��˳.getThsDataItem(24,280,84,303,"00e600-000000|ff3232-000000|ffffff-000000")
	openPrice = Lib.ͬ��˳.getThsDataItem(30,187,73,204,"00e600-000000|ff3232-000000|ffffff-000000")
	highPrice = Lib.ͬ��˳.getThsDataItem(30,220,77,235,"00e600-000000|ff3232-000000|ffffff-000000")
	lowerPrice = Lib.ͬ��˳.getThsDataItem(32,250,76,269,"00e600-000000|ff3232-000000|ffffff-000000")
	cjl = Lib.ͬ��˳.getThsDataItem(25,377,71,401, "02e2f4-000000")
	//MessageBox "cjl=" & cjl
	curDate = Lib.ͬ��˳.getThsDataItem(38,109,80,125, "c0c0c0-000000")
	//MessageBox "curDate=" & curDate
	curTimeStr = Lib.ͬ��˳.getThsDataItem(39, 118, 76, 139, "c0c0c0-000000")
	//MessageBox "curTimeStr=" & curTimeStr & ",0:" & Asc(curTimeStr)
	curTime = Mid(curTimeStr,2,4)

	//MessageBox "curTime=" & curTime
	getTestPrice = newPrice & "|" & openPrice & "|" & highPrice & "|" & lowerPrice & "|" & cjl & "|" & curDate & "|" & curTime
End Function