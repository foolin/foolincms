<!--#include file="../plugin_inc.asp"-->
<%
Dim page: page = CPage(Req("page"))	'��ǰҳ��
Dim id: id = Req("id")

If Len(id) = 0 Or Not IsNumeric(id) Then id = 0
If InStr(id, ",") > 0 Then id = Mid(id, 1, InStr(id, ",") - 1)
Dim SitePath: SitePath = ColPath(id, 0)	'��ǰ·��

Dim tpl	'ģ����ʵ��
Set tpl = New ClassTemplate
	tpl.Page = page						'���õ�ǰҳ
	Call tpl.LoadTpl("blog.html")		'����ģ��
	Call tpl.Compile_List(id)				'���б�ǩ����
	Response.Write(tpl.Content)			'�������
Set tpl = Nothing

%>