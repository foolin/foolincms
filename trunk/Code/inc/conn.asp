<%
'�ж��Ƿ�����IP�û�
If ChkLimitIp = True Then
	Call MsgBox("����IP�Ѿ�������������ϵ����Ա", "CLOSE")
End If

'�����ݿ�����
Dim ConnStr, Conn
ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBPath)
Set   Conn=Server.CreateObject("ADODB.Connection")  
Conn.Open ConnStr
If Err Then
	Err.Clear
	Set Conn = Nothing
	Response.Write "���ݿ����ӳ����������ݿ������ļ��е����ݿ�������á�"
	Response.End
End If
%>
