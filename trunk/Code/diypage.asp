<!--#include file="inc/include.asp"-->
<%
Dim page: page = CPage(Req("page"))	'��ǰҳ��
Dim param
If Len(Req("id")) > 0 Then
	param = Cint(Req("id"))
ElseIf Len(Req("url")) > 0 Then
	param = Req("url")
Else
	Response.Write(Warn("�����������飡")): Response.End()
End If
Dim SitePath: SitePath = DiyPagePath(param)	'��ǰ·��


Dim tpl	'ģ����ʵ��
Set tpl = New ClassTemplate
	tpl.Page = page						'���õ�ǰҳ
	Call tpl.Compile_DiyPage(param)			'���б�ǩ����
	Response.Write(tpl.Content)			'�������
Set tpl = Nothing

Call ConnClose()	'�ر�����
%>
