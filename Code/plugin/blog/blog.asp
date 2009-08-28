<!--#include file="../plugin_inc.asp"-->
<%
Dim page: page = CPage(Req("page"))	'当前页数
Dim id: id = Req("id")

If Len(id) = 0 Or Not IsNumeric(id) Then id = 0
If InStr(id, ",") > 0 Then id = Mid(id, 1, InStr(id, ",") - 1)
Dim SitePath: SitePath = ColPath(id, 0)	'当前路径

Dim plugin	'插件类实例
Set plugin = New ClassPlugin
	plugin.ChkState
	plugin.NewTpl("blog.html")
	Response.Write plugin.GetTpl()
Set plugin = Nothing

%>