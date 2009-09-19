<%
Function WebLog(ByVal strAction, ByVal strUser)
	If ISWEBLOG = 1 Or ISWEBLOG = True Then
		If Len(Trim(strUser)) = 0 Then
			strUser = "сн©м"
		End If
		If UCase(strUser) = "SESSION" Then
			strUser = Session("AdminName")
		End If
		DB "INSERT INTO WebLog(Username, UserAction, UserIP, ActionUrl, CreateTime) VALUES('" & strUser & "', '" & strAction & "', '" & GetIP() & "', '" & GetUrl() & "', '" & Now() & "')", 0
	End If
End Function
%>