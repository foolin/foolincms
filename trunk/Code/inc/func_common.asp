<%
'��ȡRequestֵ��������SQL�����ַ�
Function Req(str)
	Dim strReq: strReq = Trim(Request(str))
	If strReq <> "" Then
		strReq = FilterStr(strReq)
	Else
		strReq = ""
	End If
	Req = strReq
End Function

' ���ݿ����
' SqlStr - SQL�ַ����� SQLType - �����ݿ������
Function DB(Byval SqlStr, Byval SQLType)
    Select Case SQLType
    Case 0
        Conn.Execute (SqlStr)
    Case 1
        Set DB = Conn.Execute(SqlStr)
    Case 2
        Set DB = Server.CreateObject("Adodb.Recordset")
        DB.Open SqlStr, Conn, 1, 1
    Case 3:
        Set DB = Server.CreateObject("Adodb.Recordset")
        DB.Open SqlStr, Conn, 1, 3
    End Select
End Function

'���
Function Echo(Byval str)
	Response.Write(str)
End Function

'ѭ�����
Function LoopEcho(Byval str, Byval iLoop)
	Dim i
	If Not IsNumeric(iLoop) Or iLoop < 1 Then iLoop = 1
	For i = 0 To iLoop
		Response.Write(str)
	Next
End Function

'�ر����ݿ�
Sub ConnClose()
	'���ݲ��ԣ��رպ�VarType(conn)=9�������رգ���ΪVarType(conn)=8
	If VarType(conn) = 8 Then Conn.close: Set Conn = Nothing
End Sub

'����SQL�Ƿ��ַ�
' FilterStr -- ��Ҫ�����ַ���
Function FilterStr(ChkStr)
	dim Str:Str=Trim(ChkStr)
	If isnull(Str) Then
	   checkStr = ""
	   Exit Function
	Else
	   Str = replace(Str,"'","")
	   Str = replace(Str,";","")
	   Str = replace(Str,"-","")
	   FilterStr = Str
	End If
End Function

' ��ֹ�ⲿ�ύ
Function CheckPost()
	Dim server_v1, server_v2
	server_v1 = cstr(request.servervariables("http_referer"))
	server_v2 = cstr(request.servervariables("server_name"))
	If Mid(server_v1, 8, len(server_v2)) <> server_v2 Then
		CheckPost = false
	Else
		CheckPost = true
	End If
End Function


'//��ȡ�����û�IP
Function GetIP()
	Dim Ip, Tmp
	Dim i, IsErr
	IsErr = False
	Ip=Request.ServerVariables("REMOTE_ADDR")
	If Len(Ip) <= 0 Then Ip = Request.ServerVariables("HTTP_X_ForWARDED_For")		
	If Len(Ip) > 15 Then 
		IsErr = True
	Else
		Tmp = Split(Ip, ".")
		If Ubound(Tmp) = 3 Then 
			For i = 0 To Ubound(Tmp)
				If Len(Tmp(i)) > 3 Then IsErr = True
			Next
		Else
			IsErr = True
		End If
	End If
	If IsErr Then 
		GetIP = "1.1.1.1"
	Else
		GetIP = Ip
	End If
End Function

'�����ַ���strWarn-������Ҫ��������֣����غ�ɫ����
Function Warn(strWarn)
	Warn = "<font color=red>" & strWarn & "</font>"
End Function

'�����ַ���strWarn-������Ҫ��������֣����غ�ɫ����
Function ErrMsg(str)
	Response.Write("<font color=red>" & str & "</font>")
	Response.End()
End Function

' IIF
Function IIF(Byval A,Byval B,Byval C)
	If A Then IIF = B Else IIF = C
End Function



'=================================================
'��  ����MsgBox
'��  �ã��ɹ���ʾ��ʾ
'��  ����Msg-�ɹ���Ϣ��Url-ת���ַ
'Author��Foolin Time:2009-3-22
'=================================================
Function MsgBox(Byval Msg,Byval Url)
	If Msg = ""  Then
		Msg = "�Բ���δ֪����"
	End If
	If UCase(Url)="BACK" Then 	'���ز�ˢ��ҳ��
		Response.write "<script type='text/javascript'>alert('"&Msg&"');history.go(-1);</script>"
	ElseIf UCase(Url)="REFRESH" Then  '����ˢ��ҳ��
		Response.write "<script type='text/javascript'>alert('"&Msg&"');this.location.href='"&request.ServerVariables("HTTP_REFERER")&"';</script>" 
	Else	'��ַ�ض���
		Response.write "<script type='text/javascript'>alert('"&Msg&"');location.href='"&Url&"';</script>"
	End If
	Response.End()
End Function

'������ת
Function JumpUrl(Byval Url)
	If UCase(Url)="BACK" Then 	'���ز�ˢ��ҳ��
		Response.write "<script type='text/javascript'>history.go(-1);</script>"
	ElseIf UCase(Url)="REFRESH" Then  '����ˢ��ҳ��
		Response.write "<script type='text/javascript'>this.location.href='"&request.ServerVariables("HTTP_REFERER")&"';</script>" 
	Else	'��ַ�ض���
		Response.write "<script type='text/javascript'>location.href='"&Url&"';</script>"
	End If
	Response.End()
End Function

'ҳ����ת����JumpUrl.asp����
Function MsgAndGo(Byval Msg,Byval Url)
	Response.Redirect "jumpurl.asp?msg=" & Msg & "&jumpurl=" & Url & "&time=3"
End Function

'����ȷ�϶Ի���Url1��Ϊ��ʱ��ת��Url2Ϊ��ʱ��ת
Function Confirm(Byval Msg,Byval Url1,Byval Url2)
	Dim strUrl1, strUrl2
	If Url1="BACK" Then 
		strUrl1="history.go(-1);"
	Else
		strUrl1="location.href='"&Url1&"';"
	End If
	If Url2="BACK" Then 
		strUrl2="history.go(-1);"
	Else
		strUrl2="location.href='"&Url2&"';"
	End If
	Response.write "<script type=""text/javascript"">If(confirm('"&Msg&"')){"&strUrl1&"}Else{"&strUrl2&"}</script>"
	Response.End()
End Function

'�����༭��
Function CreateEditor(ByVal eName, ByVal eValue)
	Dim oFCKeditor' �������
	Dim returnEditor
	Set oFCKeditor = New FCKeditor' ��ĳ�ʼ��
	oFCKeditor.config("AutoDetectLanguage") = false
	oFCKeditor.config("DefaultLanguage") = "zh-cn"
	oFCKeditor.BasePath = INSTALLDIR & "/admin/fckeditor/"' ����·�������Ǹ�·����/FCKeditor/��
	oFCKeditor.ToolbarSet = "Default"	'���幤������Ĭ��Ϊ��Default��
	oFCKeditor.Width = "100%"	'�����ȣ�Ĭ�Ͽ�ȣ�100%��
	oFCKeditor.Height = 450		'����߶ȣ�Ĭ�ϸ߶ȣ�200��
	oFCKeditor.Value = eValue	' �����ĳ�ʼֵ
	returnEditor = oFCKeditor.Create(eName)
	Set oFCKeditor = Nothing
	CreateEditor = returnEditor
End Function


'ʱ���ʽ���
Function Fdate( Dat, n)
	Fdate = FormatDateTime( Dat , n )	
End Function


'��ȡ�ַ���
Function CutStr(str, length)
	Dim temp,intLen
	intLen = Cint(length)
	If Len(str) > intLen Then
		temp = Left(str, intLen) & "..."
	Else
		temp = str
	End If
	CutStr = temp
End Function


Function GetUrl() 
	Dim strHostName,strScriptName,strSubUrl,strRequestItem 
	strHostName = CStr(Request.ServerVariables("LOCAL_ADDR"))
	strScriptName = CStr(Request.ServerVariables("SCRIPT_NAME"))
	strSubUrl = ""
	If Request.QueryString<>"" Then
	   strScriptName=strScriptName&"?"
	   For Each strRequestItem In Request.QueryString
		If InStr(strScriptName,strRequestItem)=0 Then
		 If strSubUrl="" Then
		  strSubUrl=strSubUrl&strRequestItem&"="&Server.URLEncode(Request.QueryString(""&strRequestItem&""))
		 Else
		  strSubUrl=strSubUrl&"&"&strRequestItem&"="&Server.URLEncode(Request.QueryString(""&strRequestItem&""))
		 End If
		End If
	   Next
	End If
	GetUrl="http://"&strHostName&strScriptName&strSubUrl
End Function


'������ת�������֣��ϴ��ļ��õ�
Function DateToNum()
	DateToNum = Replace(Replace(Replace(Now(),"-","")," ",""),":","")
End Function

'==================================================================
'����ҳ����ת��Cpage
'��ʼ��ҳ���õ�
'==================================================================
Function CPage(page)
	if Len(page) = 0 or not isnumeric(page) or instr(page,",") > 0 then page = 1 else page = Int(page)
	if page < 1 then page = 1
	CPage = page
End Function

'ͳ��ASP����ʱ��
Function RunTime()
	Dim EndTime
	EndTime = Timer()	'StartTime��const.asp������
	RunTime = FormatNumber((EndTime - StartTime) * 1000, 3)
End Function


'��ʽ��ʱ�䣬ֻ����ʱ���ʽ���ֶ���Ч���� $yyyy-mm-dd hh:nn:ss��yy��ʾ��λ��ݣ�yyyy��ʾ��λ��ݣ�mm dd hh nn ss ���Զ�λ��ʾ��
'timeVal - ʱ�䣬 timeFormat - ��ʽ���ĸ�ʽ
Function FormatTime(timeVal, timeFormat)
	Dim tempVal
	If IsDate(timeVal) Then
		tempVal = timeVal: tempVal = LCase(timeFormat): tempVal = Replace(tempVal, "weeka", "WEEKA"): tempVal = Replace(tempVal, "montha", "MONTHA"): tempVal = Replace(tempVal, "week", "WEEK"): tempVal = Replace(tempVal, "month", "MONTH")
		If InStr(tempVal, "WEEKA") Then tempVal = Replace(tempVal, "WEEKA", Lang_Week_Abbr(Weekday(timeVal)))
		If InStr(tempVal, "WEEK") Then tempVal = Replace(tempVal, "WEEK", Lang_Week(Weekday(timeVal)))
		If InStr(tempVal, "MONTHA") Then tempVal = Replace(tempVal, "MONTHA", Lang_Month_Abbr(Month(timeVal)))
		If InStr(tempVal, "MONTH") Then tempVal = Replace(tempVal, "MONTH", Lang_Month(Month(timeVal)))
		If InStr(tempVal, "yyyy") > 0 Then tempVal = Replace(tempVal, "yyyy", Year(timeVal))
		If InStr(tempVal, "yy") > 0 Then tempVal = Replace(tempVal, "yy", Right(Year(timeVal), 2))
		If InStr(tempVal, "mm") > 0 Then tempVal = Replace(tempVal, "mm", Right("0" & Month(timeVal), 2))
		If InStr(tempVal, "m") > 0 Then tempVal = Replace(tempVal, "m", Month(timeVal))
		If InStr(tempVal, "dd") > 0 Then tempVal = Replace(tempVal, "dd", Right("0" & Day(timeVal), 2))
		If InStr(tempVal, "d") > 0 Then tempVal = Replace(tempVal, "d", Day(timeVal))
		If InStr(tempVal, "hh") > 0 Then tempVal = Replace(tempVal, "hh", Right("0" & Hour(timeVal), 2))
		If InStr(tempVal, "h") > 0 Then tempVal = Replace(tempVal, "h", Hour(timeVal))
		If InStr(tempVal, "nn") > 0 Then tempVal = Replace(tempVal, "nn", Right("0" & Minute(timeVal), 2))
		If InStr(tempVal, "n") > 0 Then tempVal = Replace(tempVal, "n", Minute(timeVal))
		If InStr(tempVal, "ss") > 0 Then tempVal = Replace(tempVal, "ss", Right("0" & Second(timeVal), 2))
		If InStr(tempVal, "s") > 0 Then tempVal = Replace(tempVal, "s", Second(timeVal))
	Else
		tempVal = timeVal
	End If
	FormatTime  =  tempVal
End Function

'ȥ��HTML����
Function ClearHtml(strHtml)
	Dim objRegExp, strOutput
	Set objRegExp = New Regexp ' ����������ʽ
	
	objRegExp.IgnoreCase = True ' �����Ƿ����ִ�Сд
	objRegExp.Global = True '��ƥ�������ַ�������ֻ�ǵ�һ��
	objRegExp.Pattern = "<[^>]*>" ' ����ģʽ�����е���������ʽ�������ҳ�html��ǩ
	
	strOutput = objRegExp.Replace(strHtml, "") '��html��ǩȥ��
	Set objRegExp = Nothing
	ClearHtml = Trim(strOutput)
End Function


'---------------------------------------------------
'	������	IndexPath
'	���ܣ�	��ҳ���ӵ���
'	������	��
'---------------------------------------------------
Function IndexPath()
	IndexPath = ToLink("��ҳ", "index.asp")
End Function

'---------------------------------------------------
'	������	ColPath
'	���ܣ�	��Ŀ���ӵ���
'	������	id - ��ĿID, srcType - 1 ͼƬ�� 0 - ����
'---------------------------------------------------
Function ColPath(ByVal id, ByVal colType)
	ColPath = ToPath(IndexPath, GetColLink(id, colType))
End Function

'---------------------------------------------------
'	������	ArtPath
'	���ܣ�	��Ŀ���ӵ���
'	������	id - ����ID
'---------------------------------------------------
Function ArtPath(ByVal id)
	Dim Rs, strSql, strPath
	If Len(id) = 0 Or Not IsNumeric(id) Then PicPath = "": Exit Function
	strSql = "SELECT ColID FROM Article WHERE ID = " & id
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, GetColLink(Rs("ColID"), 0))
		strPath = ToPath(strPath, "�������")
	Else
		strPath = "�������"
	End If
	Rs.Close: Set Rs = Nothing
	ArtPath = strPath
End Function

'---------------------------------------------------
'	������	PicPath
'	���ܣ�	��Ŀ���ӵ���
'	������	id - ͼƬid
'---------------------------------------------------
Function PicPath(ByVal id)
	Dim Rs, strSql, strPath
	If Len(id) = 0 Or Not IsNumeric(id) Then PicPath = "": Exit Function
	strSql = "SELECT ColID FROM Picture WHERE ID = " & id
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, GetColLink(Rs("ColID"), 1))
		strPath = ToPath(strPath, "ͼƬ���")
	Else
		strPath = "ͼƬ���"
	End If
	Rs.Close: Set Rs = Nothing
	PicPath = strPath
End Function

'---------------------------------------------------
'	������	DiyPagePath
'	���ܣ�	��Ŀ���ӵ���
'	������	id - ҳ��id
'---------------------------------------------------
Function DiyPagePath(ByVal id)
	Dim Rs, strSql, strPath
	If Len(id) = 0 Or Not IsNumeric(id) Then DiyPagePath = "": Exit Function
	strSql = "SELECT Title FROM DiyPage WHERE ID = " & id
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, Rs("Title"))
	Else
		strPath = "�Զ���ҳ��"
	End If
	Rs.Close: Set Rs = Nothing
	DiyPagePath = strPath
End Function

'---------------------------------------------------
'	������	ToPath
'	���ܣ�	���������������Ӵ�������
'	������	strLinkA - ��һ�����ӣ� strLinkB - �ڶ�������
'---------------------------------------------------
Function ToPath(ByVal strLinkA, ByVal strLinkB)
	ToPath = strLinkA & SitePathSplit & strLinkB
End Function

'---------------------------------------------------
'	������	ToLink
'	���ܣ�	������ת��������
'	������	strName - ���ƣ� strUrl - ����
'---------------------------------------------------
Function ToLink(ByVal strName, ByVal strUrl)
	ToLink = "<a href=""" & strUrl & """>" & strName & "</a>"
End Function
%>
