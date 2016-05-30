[General]
SyntaxVersion=2
MacroID=f4164dda-74e3-4ddc-bec6-432af2423c52
[Comment]

[Script]
//请在下面写上您的子程序或函数
//写完保存后，在任一命令库上点击右键并选择“刷新”即可
Sub 删除文件(filePath)
    IsFile = Plugin.File.IsFileExist(filePath)
    If IsFile = true Then 
        Call Plugin.File.DeleteFile(filePath)	
    End If	
End Sub


Function getHwnd(title)
    HwndEx = Plugin.Window.Search(title)
    //TracePrint HwndEx
    MyArray = Split(HwndEx, "|")
    If UBound(MyArray)>=0 then
        i=0
        For UBound(MyArray)
            Hwnd = Clng(MyArray(i))
            //激活显示窗口
            TracePrint title & "的指针" & Hwnd
            Call Plugin.Window.Active(Hwnd)
            Delay 500
            Call Plugin.Window.Max(Hwnd)
            Delay 500
            getHwnd = Hwnd
            Exit function
            i=i+1
        Next
    End If
    getHwnd = - 1 
End Function

Sub addStockInFilename(filename, stockCode)
	row =  Lib.公用函数.searchStockRow(filename,stockCode)
	If row > 0 Then 
		//股票已经存在，不保存直接返回
		Exit Sub
	End If
    Length = Plugin.File.GetFileLength(filename)
    adFile = Plugin.File.OpenFile(filename)
    //TracePrint Length
    If Length > 0 Then 
        Call Plugin.File.SeekFile(adFile, Length)
    End if
    Call Plugin.File.WriteLine(adFile,stockCode)
    Call Plugin.File.CloseFile(adFile)
End Sub
Function verifyInputCode(code, x1, y1, x2, y2, image)
    Call InputKeyboardCode(code)
    TracePrint "verifyInputCode=" & code
    Delay 3000
    FindPic x1, y1, x2, y2, image, 0.9, intX, intY
    //MessageBox "intx=" & intX & ",inty=" & intY
    If intX > 0 And intY > 0 Then 
        verifyInputCode = 1
    Else 
        verifyInputCode = -1
    End If
End Function
Function verifyMouse(left, up, x1, y1, x2, y2, image)
    Call MouseClick(left, up)
    Delay 1000
    FindPic x1, y1, x2, y2, image,0.9,intX,intY
    If intX > 0 And intY > 0 Then 
        verifyMouse = 1
    Else 
        verifyMouse = -1
    End If
End Function
Sub InputKeyboardCodeNoEnter(code)
    For i = 1 To Len(code)
        ch = Mid(code, i, 1)
        KeyPress ch, 1
    Next
End Sub
Sub InputKeyboardCode(code)
    For i = 1 To Len(code)
        ch = Mid(code, i, 1)
        KeyPress ch, 1
    Next
    KeyPress "enter", 1
    Delay 500
End Sub
Sub MouseClick(left, up)
    MoveTo left, up
    LeftClick 1
End Sub

Sub inputTextBox(left, up, str)
	Call MouseClick(left, up)
	Delay 500
	KeyPress "Backspace", 30
	SayString str
	Delay 500
End Sub

Function deleteStockFromFile(filename, stockCode)
	row = searchStockRow(filename, stockCode)
	If row > 0 Then 
		Call Lib.文件.删除指定行文本内容(filename, row)
		deleteStockFromFile = true
	Else 
		deleteStockFromFile = false
	End If
End Function
/*
查找StockCode在文件中的行数 
返回值： 
   -1	没有找到文件
   >0   返回指定的行数
*/
Function searchStockRow(filename, stockCode)
	fd = Plugin.File.OpenFile(filename)
	If fd = -1  Then 
		Call Plugin.File.CloseFile(fd)
		searchStockRow = - 1 
		Exit Function
	End If
	
	num = 1
	cmpStockCode = Plugin.File.ReadLine(fd)
	While cmpStockCode <> ""
		If cmpStockCode = stockCode Then 
			Call Plugin.File.CloseFile(fd)
			searchStockRow = num
			Exit Function
		End If
        str = Plugin.File.ReadLine(fd)
        cmpStockCode = Mid(str, 2)
        num = num + 1
    Wend
    searchStockRow = 0
    Call Plugin.File.CloseFile(fd)
End Function

Function trimString(str)
	//MessageBox "trimStr=" & str
	pos = InStr(str, ":")
	If pos > 1 Then 
		str = Mid(str, pos)
		//MessageBox "delete char_:" & str
	End If
	
	beginPos = 1
	For i = 1 To Len(str)
		beginPos = i
		ch = Mid(str, i, 1)
		If (ch >= "0" and ch <= "9") or ch = "+" or ch = "-" Then 
			Exit For
		End If
	Next

	endPos = Len(str)
	//MessageBox "endPos=" & endPos
	For j = endPos To 1 Step - 1 
		endPos = j
		ch = Mid(str, j, 1)
		If ch >= "0" and ch <= "9" Then 
			Exit for
		End If
		//MessageBox "ch=" & ch & ",j=" & j
	Next
	
	tmpStr = Mid(str, beginPos, endPos - beginPos + 1)
	trimString = Replace(tmpStr, "--", "-")
	
	//MessageBox "i=" & i & ",j=" & j & ",str=" & str & ",retString=" & trimString
End Function

Function sendEmail(title, msg, filename)
	返回值 = Lib.网络.发送邮件("cjing@vip.sina.com","62561741","13801164670@139.com",title,msg,filename)
	sendEmail = 返回值
End Function

Function saveScreen(filename)
	Delay 500
	Call Plugin.Pic.PrintScreen(0, 0, 1336, 768, filename)
End Function