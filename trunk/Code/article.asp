<!--#include file="include/include.asp"-->
<%
Dim page: page = CPage(request("page"))
Dim id: id = CInt(Request("id"))

Dim tpl	'模板类实例
Set tpl = New TemplateClass
	tpl.Page = page					'设置当前页
	Call tpl.Load("article.html")		'载入模板
	'Call tpl.Parser_Run()			'运行标签分析
	Call tpl.Parser_Field(id, False)	'运行标签field分析
	Response.Write(tpl.Content)		'输出内容
Set tpl = Nothing

Response.Write GetPreLink(10, 0, 2) & "<br>"

Response.Write GetNextLink(10, 0, 2) & "<br>"

Response.Write( "运行速度" & RunTime() & "毫秒")

%>