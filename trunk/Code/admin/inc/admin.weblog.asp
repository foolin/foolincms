<%
'添加日志
Function WebLog(ByVal strAction, ByVal strUser)
	If ISWEBLOG = 1 Or ISWEBLOG = True Then
		If Len(Trim(strUser)) = 0 Then
			strUser = "游客"
		End If
		Select Case UCase(strUser)
			Case "SESSION"
				strUser = Session("AdminName")
			Case "COOKIES"
				strUser = GetCookies("AdminName")
			Case "SELF"
				strUser = GetCookies("AdminName")
		End Select
		'如果用户名为空，则为自己
		If Len(strUser) = 0 Then
			strUser = GetCookies("AdminName")
		End If
		DB "INSERT INTO WebLog(Username, UserAction, UserIP, ActionUrl, CreateTime) VALUES('" & strUser & "', '" & strAction & "', '" & GetIP() & "', '" & Request.ServerVariables("HTTP_REFERER") & "', '" & Now() & "')", 0
	End If
End Function

'删除日志
Function DelWebLog(ByVal ids)
	If Not IsNumeric(ids) Then Call MsgBox("参数错误","BACK"): DelWebLog = False
	'DB "DELETE WebLog WHERE ID IN (" & ids & ")", 0
	DB "Delete From [WebLog] Where [ID] IN ("& ids &")" ,0
	DelWebLog = True
End Function

'清空整个日志
Function ClearWebLog()
	DB "DELETE FROM [WebLog]", 0
	ClearWebLog = True
End Function
%>