<!--#include file="include/include.asp"-->
<%
Dim id: id = Req("id")

Response.Write id & "<br>"

'Response.Write(Replace("hehe//helo/world", "//", "/")): Response.End()

Dim tpl	'ģ����ʵ��
Set tpl = New TemplateClass
	'Call tpl.Parser_Run()			'���б�ǩ����
	Call tpl.Parser_DiyPage(id)	'���б�ǩfield����
	Response.Write(tpl.Content)		'�������
Set tpl = Nothing


Response.Write( "<br>�����ٶ�" & RunTime() & "����")

%>
