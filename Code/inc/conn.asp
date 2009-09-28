<%
'判断是否限制IP用户
If ChkLimitIp = True Then
	Call MsgBox("您的IP已经被锁定，请联系管理员", "CLOSE")
End If

'打开数据库连接
Dim ConnStr, Conn
ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBPath)
Set   Conn=Server.CreateObject("ADODB.Connection")  
Conn.Open ConnStr
If Err Then
	Err.Clear
	Set Conn = Nothing
	Response.Write "数据库连接出错，请检查数据库连接文件中的数据库参数设置。"
	Response.End
End If
%>
