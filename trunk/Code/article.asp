<!--#include file="inc/include.asp"-->
<%
Dim page: page = CPage(Req("page"))	'��ǰҳ��
Dim id: id = Req("id")
If Len(id) = 0 Or Not IsNumeric(id) Then ErrMsg("id��������")
If InStr(id, ",") > 0 Then id = Mid(id, 1, InStr(id, ",") - 1)
'��ǰҳ����
Dim Title: Title = "����"	
If id > 0 Then Title = GetTitleOfArtOrPic(id, "ARTICLE")
'��ǰ·��
Dim SitePath: SitePath = ArtPath(id)	


Dim tpl	'ģ����ʵ��
Set tpl = New ClassTemplate
	tpl.Page = page						'���õ�ǰҳ
	Call tpl.LoadTpl("article.html")		'����ģ��
	Call tpl.Compile_Field(id, False)	'���б�ǩ����
	Response.Write(tpl.Content)			'�������
Set tpl = Nothing

Call ConnClose()	'�ر�����
%>