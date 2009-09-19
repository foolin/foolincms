<% 	
Dim ConnStr, Conn

'--------打开数据库-------------
ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../" & DBPath)
Set   Conn=Server.CreateObject("ADODB.Connection")  
Conn.Open ConnStr
If Err Then
	Err.Clear
	Set Conn = Nothing
	Response.Write "数据库连接出错，请检查数据库连接文件中的数据库参数设置。"
	Response.End
End If
%>
