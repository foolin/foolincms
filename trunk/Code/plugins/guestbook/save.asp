<!--#include file="../../inc/config.asp"-->
<!--#include file="../../inc/func_common.asp"-->
<!--#include file="../inc/plugin.conn.asp"-->
<%	'判断是否开放留言
	If ISOPENGBOOK = 0 Then
		Call MsgBox("对不起，本站留言功能暂不开放！","BACK")
	End If
	'判断数据来源是否非法
	Dim serverUrl1, serverUrl2
	serverUrl1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
	serverUrl2 = Cstr(Request.ServerVariables("SERVER_NAME"))
	If Mid(serverUrl1, 8, Len(serverUrl2)) <>  serverUrl2 Then
		Call MsgBox("对不起，本站禁止外部提交数据！","REFRESH")
	End If
	'检测是否频繁留言
	If DateDIff("s", CDate(Session("PostTime")), Now()) < GBOOKTIME Then
		Call MsgBox("禁止频繁提交留言！请"& GBOOKTIME &"秒之后再留言...","BACK")
	End If
	'检查验证码
	If Len(Request("fChkCode"))=0 Then Call MsgBox("验证码不能为空！", "BACK")
	If Request("fChkCode") <> Session("ChkCode") Then
		Call MsgBox("验证码不正确！", "BACK")
	End If
	'处理留言
	Dim strTitle, strContent, strUser, strEmail, strHomepage, strIP, strSql
		strTitle = Req("fTitle")
		strContent = Req("fContent")
		strUser = Left(Req("fUser"),20)
		strEmail = Left(Req("fEmail"),20)
		strHomePage = Left(Req("fHomePage"),50)
		strIP = GetIP()
		If Len(strTitle)<1 Or Len(strContent)>50 Then Call MsgBox("标题的长度请控制在 1 至 50 位","BACK")
		If Len(strContent)<1 Or Len(strContent)>250 Then Call MsgBox("内容的长度请控制在 1 至 250 位","BACK")
		If Len(strUser) = 0 Then strUser = "匿名"
		If Len(strEmail) = 0 Then strEmail = "@"
		If Len(strHomePage) = 0 Then strHomePage = "#"
		'过滤敏感字符
		strTitle = FilterDirtyStr(strTitle)
		strContent = FilterDirtyStr(strContent)
		'判断是否需要审核
		If ISAUDITGBOOK = 0 Then
			strSql = "INSERT INTO GuestBook([Title],[Content],[User],[Email],[HomePage],[IP],[CreateTime], [State])"
			strSql = strSql & " VALUES('"& strTitle &"','"& strContent &"','"& strUser &"','"& strEmail &"','"& strHomePage &"','"& strIP &"','"& Now() &"', 1)"
		Else
			strSql = "INSERT INTO GuestBook([Title],[Content],[User],[Email],[HomePage],[IP],[CreateTime], [State])"
			strSql = strSql & " VALUES('"& strTitle &"','"& strContent &"','"& strUser &"','"& strEmail &"','"& strHomePage &"','"& strIP &"','"& Now() &"', 0)"
		End If
	Call DB(strSql,0)
	Session("PostTime") = Now()	'记录提交时间
	Call MsgBox("你的留言已经提交，感谢您的留言！", "REFRESH")
%>
