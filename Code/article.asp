<!--#include file="include/include.asp"-->
<%
Dim page: page = CPage(request("page"))
Dim id: id = CInt(Request("id"))

Dim tpl	'ģ����ʵ��
Set tpl = New TemplateClass
	tpl.Page = page					'���õ�ǰҳ
	Call tpl.Load("article.html")		'����ģ��
	'Call tpl.Parser_Run()			'���б�ǩ����
	Call tpl.Parser_Field(id, False)	'���б�ǩfield����
	Response.Write(tpl.Content)		'�������
Set tpl = Nothing

Response.Write GetPreLink(10, 0, 2) & "<br>"

Response.Write GetNextLink(10, 0, 2) & "<br>"

Response.Write( "�����ٶ�" & RunTime() & "����")

%>