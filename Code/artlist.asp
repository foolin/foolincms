<!--#include file="inc/include.asp"-->
<%
Dim page: page = CPage(Req("page"))	'��ǰҳ��
Dim id: id = Req("id")
If Len(id) = 0 Or Not IsNumeric(id) Then id = 0
If InStr(id, ",") > 0 Then id = Mid(id, 1, InStr(id, ",") - 1)
'��ǰҳ����
Dim Title: Title = "�����б�"	
If id > 0 Then Title = GetNameOfColumn(id, "ARTICLE") & " - �����б�"
'��ǰ·��
Dim SitePath: SitePath = ColPath(id, 0)	

Dim tpl	'ģ����ʵ��
Set tpl = New ClassTemplate
	tpl.Page = page						'���õ�ǰҳ
	Call tpl.LoadTpl("artlist.html")		'����ģ��
	Call tpl.Compile_List(id)				'���б�ǩ����
	Response.Write(tpl.Content)			'�������
Set tpl = Nothing

Call ConnClose()	'�ر�����
%>