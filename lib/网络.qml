[General]
SyntaxVersion=2
MacroID=a1c52b4e-dc44-4ded-887c-ff94144acf68
[Comment]
������ǰ�������8.0���Ƴ���ȫ�¹���
�����԰��Լ����õĺ������ӳ���д����������úܶ���ű�ȥ����
����������������ö���ű�����һ������޸�һ���͵����޸Ķദ
Ŀǰ����⹦�ܻ��ڲ��Ե��У����κν�������ڰ���������̳�������ַ��http://bbs.ajjl.cn
******ע�⣡���ǹٷ��ṩ������⣬�����޸ģ������Ժ󰴼���������ʱ���������޸ġ�******//
******          ������������⣬������������Ҽ�ѡ���½��������            ******//

[Script]
Function �����ҳԴ�ļ�(��ҳ��ַ)
    //˵����֧��Զ�̻�ȡ�ı����ݣ��磺MsgBox lib.����.�����ҳԴ�ļ�("http://www.anjian.com/test.txt")
    //���ӣ�MsgBox lib.����.�����ҳԴ�ļ�("http://www.anjian.com")
    Dim xmlHttp, xmlBody, xmlUrl
    Dim ThisCharCode ,NextCharCode ,BytesToBstr
    If InStr(��ҳ��ַ, "http://") = 0 Then 
        xmlUrl = "http://" & ��ҳ��ַ
    Else
        xmlUrl = ��ҳ��ַ
    End if
    Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
    xmlHttp.Open "Get", xmlUrl, False
    xmlHttp.Send
    xmlBody = xmlHttp.ResponseBody
    Set xmlHttp = Nothing  
    �����ҳԴ�ļ� = ""
    If Len(xmlBody) = 0 Then Exit Function
    Set ObjStream = CreateObject("Adodb.Stream")
    With ObjStream
        .Type = 1
        .Mode = 3
        .Open
        .Write xmlBody
        .Position = 0
        .Type = 2
        .Charset = "GB2312"
        BytesToBstr = .ReadText
        .Close
    End With
    Set ObjStream = Nothing    
    �����ҳԴ�ļ� = BytesToBstr
End Function
Function �������IP��ַ()   
    //���ӣ�MsgBox lib.����.�������IP��ַ()
    Dim ��ҳ����,��ʼλ��,����λ��
    ��ҳ���� = lib.����.�����ҳԴ�ļ�("http://city.ip138.com/ip2city.asp")
    ��ʼλ�� = inStr(��ҳ����,"[") + 1
    ����λ�� = inStr(��ҳ����,"]") - ��ʼλ��
    �������IP��ַ = Mid(��ҳ����,��ʼλ��,����λ��)
End Function
Function �����ʼ�(��������ʺ�, �����������, �����ʼ���ַ, �ʼ�����, �ʼ�����, �ʼ�����) 
    //���ӣ�MsgBox lib.����.�����ʼ�("ceshi0000001@163.com","ceshi000001","ceshi0000001@163.com","�ʼ�����","�ʼ�����","")
    Dim You_ID,MS_Space,Email
    '�ʺźͷ��������� 
    You_ID = Split(��������ʺ�, "@") 
    '����Ǳ���Ҫ�ģ��������Է��ĵ��£�����ͨ��΢�����ʼ� 
    MS_Space = "http://schemas.microsoft.com/cdo/configuration/" 
    Set Email = CreateObject("CDO.Message") 
    '���һ��Ҫ�ͷ����ʼ����ʺ�һ��
    Email.From = ��������ʺ� 
    'Execute "Email.to = �����ʼ���ַ"
    Email.CC = �����ʼ���ַ
    Email.Subject = �ʼ�����
    Email.Textbody = �ʼ����� 
    If �ʼ����� <> "" Then 
        Email.AddAttachment �ʼ����� 
    End If 
    With Email.Configuration.Fields 
        '���Ŷ˿� 
        .Item(MS_Space & "sendusing") = 2 
        'SMTP��������ַ 
        .Item(MS_Space & "smtpserver") = "smtp." & You_ID(1) 
        'SMTP�������˿� 
        .Item(MS_Space & "smtpserverport") = 25 
        .Item(MS_Space & "smtpauthenticate") = 1
        .Item(MS_Space & "sendusername") = You_ID(0) 
        .Item(MS_Space & "sendpassword") = �����������  
        .Update 
    End With 
    '�����ʼ� 
    Email.Send 
    '�ر���� 
    Set Email = Nothing 
    �����ʼ� = True
    '���û���κδ�����Ϣ�����ʾ���ͳɹ�,������ʧ�� 
    If Err Then 
        Err.Clear 
        �����ʼ� = False 
    End If 
End Function 
Function ��ȡ����ʱ��()
    //���ӣ�MsgBox "��ǰ��׼ʱ��Ϊ��" & lib.����.��ȡ����ʱ��()
    //�жϣ�If NowTime>CDate("2010-5-9") Then
    Dim SvrName(7),xPost,HttpAdd,NowTime,StartTime,i
    StartTime=Now 
    //SvrName(0) = "time-a.nist.gov"
    SvrName(1) = "time-a.timefreq.bldrdoc.gov"
    SvrName(2) = "time-b.timefreq.bldrdoc.gov"
    SvrName(3) = "time-c.timefreq.bldrdoc.gov"
    SvrName(4) = "utcnist.colorado.edu"
    SvrName(5) = "time.nist.gov"
    SvrName(6) = "nist1.datum.com"
    SvrName(7) = "nist1.aol-ca.truetime.com"
    Set xPost=createObject("Microsoft.XMLHTTP") 
    NowTime=""
    Do While NowTime=""
        For i=1 to 7
            NowTime=""
            HttpAdd="Http://" & SvrName(i) & ":13"
            xPost.Open "Put", HttpAdd, False
            xPost.Send
            Delay 10
            If xPost.readyState=4 Then
                NowTime=mid(xPost.responsetext, 8, 17)
                If NowTime<>"" Then
                    NowTime=CDate(NowTime) + 8 / 24
                    Exit Do
                Else
                    xPost.abort
                    NowTime=""
                End If
            End If
        Next
        If DateDiff("s", StartTime, Now)>=30 And NowTime="" Then
            Msgbox "��ȷ�����Ѿ��������˻�������", 0, "��ȡ����ʱ��"
            Exit Do 
        End If
    Loop
    xPost.abort
    Set xPost=Nothing
    ��ȡ����ʱ��=NowTime
End Function






//������һֻ��
//���ڣ�2009.12.30
//�޸ģ�2011.04.19




Function �����ҳԴ�ļ�_��ǿ��(��ҳ��ַ, ��ҳ����)     //�������.���ñ���.������˵88
    //Plugin.Sys.SetCLB(lib.����.�����ҳԴ�ļ�_��ǿ��("www.baidu.com","utf-8")) //�ٶ�
    //Plugin.Sys.SetCLB(lib.����.�����ҳԴ�ļ�_��ǿ��("www.sina.com.cn","gbk")) //����
    Dim xmlHttp, xmlUrl,ObjStream
    If InStr(��ҳ��ַ, "http://") = 0 Then 
        xmlUrl = "http://" & ��ҳ��ַ
    Else
        xmlUrl = ��ҳ��ַ
    End if
    Set xmlHttp = CreateObject("WinHttp.WinHttpRequest.5.1")   //���������,������/cookie ����˵88
    xmlHttp.Open "GET", xmlUrl, True
    xmlHttp.Send 
    If xmlhttp.waitforresponse() Then 
        Set ObjStream = CreateObject("Adodb.Stream")
        ObjStream.Type = 1
        ObjStream.Mode = 3
        ObjStream.Open
        ObjStream.Write xmlHttp.ResponseBody
        ObjStream.Position = 0
        ObjStream.Type = 2
        ObjStream.Charset = ��ҳ����
        �����ҳԴ�ļ�_��ǿ�� = ObjStream.ReadText

        Set ObjStream = Nothing
    Else 
        �����ҳԴ�ļ�_��ǿ�� = ""  //�����ȡʧ�ܷ���ֵ�� �� 
    End If

    Set xmlHttp = Nothing
End Function


Function ��ȡ����ʱ��_��ǿ��(��վ��ַ)
    //TracePrint lib.����.��ȡ����ʱ��_��ǿ��("msdn.microsoft.com") //ȥ΢��ȡʱ����
    //TracePrint lib.����.��ȡ����ʱ��_��ǿ��("www.taobao.com") //�����Ա�������ʱ��
    //TracePrint lib.����.��ȡ����ʱ��_��ǿ��("www.qq.com") //����qq��վʱ��
    Dim Http, URL,mt
    If InStr(��վ��ַ, "http://") = 0 Then 
        Url = "http://" & ��վ��ַ
    Else
        Url = ��վ��ַ
    End if
    Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Http.Open "HEAD", URL, True //head��ʽ,��Ϊ���������ص�ͷ���������ʱ��..���Բ���Ҫ��ҳ��
    Http.Send 
    If Http.waitforresponse() Then 
        mt = Http.getresponseheader("Date") //��ͷ��ȡ��ʱ��..js��ʽ��(�������ڼ���ʱ���ĸ�ʽ,��vbs��һ��)
        mt = Cdate(Mid(mt, 5, len(mt) - 8))  //��ȡjs��ʽ��ʱ���ַ���.�õ�vbs��ģʽ��ʱ��
        ��ȡ����ʱ��_��ǿ�� = DateAdd("h", 8, mt) //�й���8ʱ��,���������ձ��������׼ʱ,Ҳ����0ʱ����ʱ��.���Եü���8Сʱ
    Else 
        ��ȡ����ʱ��_��ǿ�� = "" //ʧ�ܷ���ֵ�� ��
    End If
    Set Http = Nothing
End Function



//�޸ģ�michael3636
//���ڣ�2015.4.28




