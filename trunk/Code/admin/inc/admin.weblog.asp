<%
'�����־
Function WebLog(ByVal strAction, ByVal strUser)
	If ISWEBLOG = 1 Or ISWEBLOG = True Then
		If Len(Trim(strUser)) = 0 Then
			strUser = "�ο�"
		End If
		Select Case UCase(strUser)
			Case "SESSION"
				strUser = Session("AdminName")
			Case "COOKIES"
				strUser = GetCookies("AdminName")
			Case "SELF"
				strUser = GetCookies("AdminName")
		End Select
		'����û���Ϊ�գ���Ϊ�Լ�
		If Len(strUser) = 0 Then
			strUser = GetCookies("AdminName")
		End If
		DB "INSERT INTO WebLog(Username, UserAction, UserIP, ActionUrl, CreateTime) VALUES('" & strUser & "', '" & strAction & "', '" & GetIP() & "', '" & Request.ServerVariables("HTTP_REFERER") & "', '" & Now() & "')", 0
	End If
End Function

'ɾ����־
Function DelWebLog(ByVal ids)
	If Not IsNumeric(ids) Then Call MsgBox("��������","BACK"): DelWebLog = False
	'DB "DELETE WebLog WHERE ID IN (" & ids & ")", 0
	DB "Delete From [WebLog] Where [ID] IN ("& ids &")" ,0
	DelWebLog = True
End Function

'���������־
Function ClearWebLog()
	DB "DELETE FROM [WebLog]", 0
	ClearWebLog = True
End Function
%>