<!--#include file="inc/include.asp"-->
<%
Dim page: page = CPage(Req("page"))	'��ǰҳ��
Dim id: id = CInt(Req("id"))
Dim SitePath: SitePath = DiyPagePath(id)	'��ǰ·��


Dim tpl	'ģ����ʵ��
Set tpl = New ClassTemplate
	tpl.Page = page						'���õ�ǰҳ
	'Call tpl.Load("diypage.html")		'����ģ��
	Call tpl.Compile_DiyPage(id)			'���б�ǩ����
	Response.Write(tpl.Content)			'�������
Set tpl = Nothing

Call ConnClose()	'�ر�����
%>
