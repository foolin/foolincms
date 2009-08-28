<!--#include file="../plugin_inc.asp"-->
<%
Dim page: page = CPage(Req("page"))	'当前页数
Dim id: id = Req("id")

If Len(id) = 0 Or Not IsNumeric(id) Then id = 0
If InStr(id, ",") > 0 Then id = Mid(id, 1, InStr(id, ",") - 1)
Dim SitePath: SitePath = ColPath(id, 0)	'当前路径

Dim tpl	'模板类实例
Set tpl = New ClassTemplate
	tpl.Page = page						'设置当前页
	Call tpl.LoadTpl("blog.html")		'载入模板
	Call tpl.Compile_List(id)				'运行标签分析
	Response.Write(tpl.Content)			'输出内容
Set tpl = Nothing

%>