<!--#include file="include/include.asp"-->
<%
Dim page: page = CPage(Req("page"))	'��ǰҳ��
Dim SitePath: SitePath = IndexPath()	'��ǰ·��

Dim tpl	'ģ����ʵ��
Set tpl = New ClassTemplate
	tpl.Page = page					'���õ�ǰҳ
	Call tpl.LoadTpl("index.html")		'����ģ��
	Call tpl.Compile_Index()		'���б�ǩ����
	Response.Write(tpl.Content)		'�������
Set tpl = Nothing

Call ConnClose()	'�ر�����
%>
