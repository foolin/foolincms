<%
Function ChkLogin()
	If Session("AdminName")="" or Session("AdminLevel")="" Then
		Response.write "<script type='text/javascript'>alert('����δ��¼');this.top.location.href='login.asp';</script>"
	End If
End Function
%>
