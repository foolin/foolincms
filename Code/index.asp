<!--#include file="include/include.asp"-->
<%
Dim page: page = CPage(request("page"))	'��ǰҳ��
Dim SitePath: SitePath = IndexPath()	'��ǰ·��

Dim tpl	'ģ����ʵ��
Set tpl = New TemplateClass
	tpl.Page = page					'���õ�ǰҳ
	Call tpl.Load("index.html")		'����ģ��
	Call tpl.Parser_Run()			'���б�ǩ����
	Response.Write(tpl.Content)		'�������
Set tpl = Nothing

Response.Write( FormatTime(Now(), "yy-mm-dd-ss"))

Response.Write( ArtPath(10))

Response.Write( PicPath(52))

Response.Write( "�����ٶ�" & RunTime() & "����")

%>
