<%
'获取Request值，并过滤SQL过敏字符
Function Req(str)
	Dim strReq: strReq = Trim(Request(str))
	If strReq <> "" Then
		strReq = FilterStr(strReq)
	Else
		strReq = ""
	End If
	Req = strReq
End Function

' 数据库操作
' SqlStr - SQL字符串， SQLType - 打开数据库的类型
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

'输出
Function Echo(Byval str)
	Response.Write(str)
End Function

'循环输出
Function LoopEcho(Byval str, Byval iLoop)
	Dim i
	If Not IsNumeric(iLoop) Or iLoop < 1 Then iLoop = 1
	For i = 0 To iLoop
		Response.Write(str)
	Next
End Function

'关闭数据库
Sub ConnClose()
	'根据测试：关闭后VarType(conn)=9，而不关闭，则为VarType(conn)=8
	If VarType(conn) = 8 Then Conn.close: Set Conn = Nothing
End Sub

'过滤SQL非法字符
' FilterStr -- 需要检测的字符串
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

' 禁止外部提交
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


'//获取来访用户IP
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

'警告字符：strWarn-输入需要警告的文字，返回红色字体
Function Warn(strWarn)
	Warn = "<font color=red>" & strWarn & "</font>"
End Function

'警告字符：strWarn-输入需要警告的文字，返回红色字体
Function ErrMsg(str)
	Response.Write("<font color=red>" & str & "</font>")
	Response.End()
End Function

' IIF
Function IIF(Byval A,Byval B,Byval C)
	If A Then IIF = B Else IIF = C
End Function



'=================================================
'函  数：MsgBox
'作  用：成功显示显示
'参  数：Msg-成功信息，Url-转向地址
'Author：Foolin Time:2009-3-22
'=================================================
Function MsgBox(Byval Msg,Byval Url)
	If Msg = ""  Then
		Msg = "对不起，未知错误！"
	End If
	If UCase(Url)="BACK" Then 	'返回不刷新页面
		Response.write "<script type='text/javascript'>alert('"&Msg&"');history.go(-1);</script>"
	ElseIf UCase(Url)="REFRESH" Then  '返回刷新页面
		Response.write "<script type='text/javascript'>alert('"&Msg&"');this.location.href='"&request.ServerVariables("HTTP_REFERER")&"';</script>" 
	Else	'地址重定向
		Response.write "<script type='text/javascript'>alert('"&Msg&"');location.href='"&Url&"';</script>"
	End If
	Response.End()
End Function

'连接跳转
Function JumpUrl(Byval Url)
	If UCase(Url)="BACK" Then 	'返回不刷新页面
		Response.write "<script type='text/javascript'>history.go(-1);</script>"
	ElseIf UCase(Url)="REFRESH" Then  '返回刷新页面
		Response.write "<script type='text/javascript'>this.location.href='"&request.ServerVariables("HTTP_REFERER")&"';</script>" 
	Else	'地址重定向
		Response.write "<script type='text/javascript'>location.href='"&Url&"';</script>"
	End If
	Response.End()
End Function

'页面跳转，与JumpUrl.asp连用
Function MsgAndGo(Byval Msg,Byval Url)
	Response.Redirect "jumpurl.asp?msg=" & Msg & "&jumpurl=" & Url & "&time=3"
End Function

'弹出确认对话框，Url1则为真时跳转，Url2为假时跳转
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

'创建编辑器
Function CreateEditor(ByVal eName, ByVal eValue)
	Dim oFCKeditor' 定义变量
	Dim returnEditor
	Set oFCKeditor = New FCKeditor' 类的初始化
	oFCKeditor.config("AutoDetectLanguage") = false
	oFCKeditor.config("DefaultLanguage") = "zh-cn"
	oFCKeditor.BasePath = INSTALLDIR & "/admin/fckeditor/"' 定义路径（这是根路径：/FCKeditor/）
	oFCKeditor.ToolbarSet = "Default"	'定义工具条（默认为：Default）
	oFCKeditor.Width = "100%"	'定义宽度（默认宽度：100%）
	oFCKeditor.Height = 450		'定义高度（默认高度：200）
	oFCKeditor.Value = eValue	' 输入框的初始值
	returnEditor = oFCKeditor.Create(eName)
	Set oFCKeditor = Nothing
	CreateEditor = returnEditor
End Function


'时间格式输出
Function Fdate( Dat, n)
	Fdate = FormatDateTime( Dat , n )	
End Function


'截取字符串
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


'将日期转换成数字，上传文件用到
Function DateToNum()
	DateToNum = Replace(Replace(Replace(Now(),"-","")," ",""),":","")
End Function

'==================================================================
'设置页数，转换Cpage
'初始化页数用到
'==================================================================
Function CPage(page)
	if Len(page) = 0 or not isnumeric(page) or instr(page,",") > 0 then page = 1 else page = Int(page)
	if page < 1 then page = 1
	CPage = page
End Function

'统计ASP运行时间
Function RunTime()
	Dim EndTime
	EndTime = Timer()	'StartTime在const.asp定义了
	RunTime = FormatNumber((EndTime - StartTime) * 1000, 3)
End Function


'格式化时间，只对于时间格式的字段有效，如 $yyyy-mm-dd hh:nn:ss，yy表示二位年份，yyyy表示四位年份，mm dd hh nn ss 都以二位表示。
'timeVal - 时间， timeFormat - 格式化的格式
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

'去除HTML代码
Function ClearHtml(strHtml)
	Dim objRegExp, strOutput
	Set objRegExp = New Regexp ' 建立正则表达式
	
	objRegExp.IgnoreCase = True ' 设置是否区分大小写
	objRegExp.Global = True '是匹配所有字符串还是只是第一个
	objRegExp.Pattern = "<[^>]*>" ' 设置模式引号中的是正则表达式，用来找出html标签
	
	strOutput = objRegExp.Replace(strHtml, "") '将html标签去掉
	Set objRegExp = Nothing
	ClearHtml = Trim(strOutput)
End Function


'---------------------------------------------------
'	函数：	IndexPath
'	功能：	首页链接导航
'	参数：	无
'---------------------------------------------------
Function IndexPath()
	IndexPath = ToLink("首页", "index.asp")
End Function

'---------------------------------------------------
'	函数：	ColPath
'	功能：	栏目链接导航
'	参数：	id - 栏目ID, srcType - 1 图片， 0 - 文章
'---------------------------------------------------
Function ColPath(ByVal id, ByVal colType)
	ColPath = ToPath(IndexPath, GetColLink(id, colType))
End Function

'---------------------------------------------------
'	函数：	ArtPath
'	功能：	栏目链接导航
'	参数：	id - 文章ID
'---------------------------------------------------
Function ArtPath(ByVal id)
	Dim Rs, strSql, strPath
	If Len(id) = 0 Or Not IsNumeric(id) Then PicPath = "": Exit Function
	strSql = "SELECT ColID FROM Article WHERE ID = " & id
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, GetColLink(Rs("ColID"), 0))
		strPath = ToPath(strPath, "文章浏览")
	Else
		strPath = "文章浏览"
	End If
	Rs.Close: Set Rs = Nothing
	ArtPath = strPath
End Function

'---------------------------------------------------
'	函数：	PicPath
'	功能：	栏目链接导航
'	参数：	id - 图片id
'---------------------------------------------------
Function PicPath(ByVal id)
	Dim Rs, strSql, strPath
	If Len(id) = 0 Or Not IsNumeric(id) Then PicPath = "": Exit Function
	strSql = "SELECT ColID FROM Picture WHERE ID = " & id
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, GetColLink(Rs("ColID"), 1))
		strPath = ToPath(strPath, "图片浏览")
	Else
		strPath = "图片浏览"
	End If
	Rs.Close: Set Rs = Nothing
	PicPath = strPath
End Function

'---------------------------------------------------
'	函数：	DiyPagePath
'	功能：	栏目链接导航
'	参数：	id - 页面id
'---------------------------------------------------
Function DiyPagePath(ByVal id)
	Dim Rs, strSql, strPath
	If Len(id) = 0 Or Not IsNumeric(id) Then DiyPagePath = "": Exit Function
	strSql = "SELECT Title FROM DiyPage WHERE ID = " & id
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, Rs("Title"))
	Else
		strPath = "自定义页面"
	End If
	Rs.Close: Set Rs = Nothing
	DiyPagePath = strPath
End Function

'---------------------------------------------------
'	函数：	ToPath
'	功能：	辅助，将两个连接串接起来
'	参数：	strLinkA - 第一个连接， strLinkB - 第二个链接
'---------------------------------------------------
Function ToPath(ByVal strLinkA, ByVal strLinkB)
	ToPath = strLinkA & SitePathSplit & strLinkB
End Function

'---------------------------------------------------
'	函数：	ToLink
'	功能：	辅助，转换成连接
'	参数：	strName - 名称， strUrl - 连接
'---------------------------------------------------
Function ToLink(ByVal strName, ByVal strUrl)
	ToLink = "<a href=""" & strUrl & """>" & strName & "</a>"
End Function
%>
