<% 	
Dim ConnStr, Conn

'--------�����ݿ�-------------
ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../" & DBPath)
Set   Conn=Server.CreateObject("ADODB.Connection")  
Conn.Open ConnStr
If Err Then
	Err.Clear
	Set Conn = Nothing
	Response.Write "���ݿ����ӳ����������ݿ������ļ��е����ݿ�������á�"
	Response.End
End If
%>
