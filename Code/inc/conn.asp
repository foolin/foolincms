<%
'�ж��Ƿ�����IP�û�
If ChkLimitIp = True Then
	Call MsgBox("����IP�Ѿ�������������ϵ����Ա", "CLOSE")
End If

'�����ݿ�����
Dim ConnStr, Conn
ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(Replace(InstallDir & "/" & DBPath,"//","/"))
Set   Conn=Server.CreateObject("ADODB.Connection")  
Conn.Open ConnStr
If Err Then
	Err.Clear
	Set Conn = Nothing
	Response.Write("���ݿ����ӳ�����������վ���ò����Ƿ���ȷ��<br /><br />")
	Response.Write("��ʾ����¼��̨���� �� ϵͳ���� �� �Զ����� �� ��ɡ�<br /><br />")
	Response.Write("������κ����ʣ��뵽�ٷ����з���<a href='http://www.eekku.com'>http://www.eekku.com</a>��<br />")
	Response.End
End If
%>
