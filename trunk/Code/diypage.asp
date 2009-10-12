<!--#include file="inc/include.asp"-->
<%
Dim page: page = CPage(Req("page"))	'当前页数
Dim param
If Len(Req("id")) > 0 Then
	param = Cint(Req("id"))
ElseIf Len(Req("url")) > 0 Then
	param = Req("url")
Else
	Response.Write(Warn("参数错误，请检查！")): Response.End()
End If
'当前页标题
Dim Title: Title = "DIY页面": Title = GetTitleOfDiypage(param)
'当前路径
Dim SitePath: SitePath = DiyPagePath(param)	

Dim tpl	'模板类实例
Set tpl = New ClassTemplate
	tpl.Page = page						'设置当前页
	Call tpl.Compile_DiyPage(param)			'运行标签分析
	Response.Write(tpl.Content)			'输出内容
Set tpl = Nothing

Call ConnClose()	'关闭连接
%>
