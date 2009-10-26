<%
'判断是否限制IP用户
If ChkLimitIp = True Then
	Call MsgBox("您的IP已经被锁定，请联系管理员", "CLOSE")
End If

'打开数据库连接
Dim ConnStr, Conn
ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(Replace(InstallDir & "/" & DBPath,"//","/"))
Set   Conn=Server.CreateObject("ADODB.Connection")  
Conn.Open ConnStr
If Err Then
	Err.Clear
	Set Conn = Nothing
	Response.Write("数据库连接出错，请检查您网站配置参数是否正确。<br /><br />")
	Response.Write("提示：登录后台管理 → 系统配置 → 自动配置 → 完成。<br /><br />")
	Response.Write("如果有任何疑问，请到官方进行反馈<a href='http://www.eekku.com'>http://www.eekku.com</a>。<br />")
	Response.End
End If
%>
