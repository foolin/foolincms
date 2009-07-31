<!--#include file="include/include.asp"-->
<%
Dim id: id = Req("id")

Response.Write id & "<br>"

'Response.Write(Replace("hehe//helo/world", "//", "/")): Response.End()

Dim tpl	'模板类实例
Set tpl = New TemplateClass
	'Call tpl.Parser_Run()			'运行标签分析
	Call tpl.Parser_DiyPage(id)	'运行标签field分析
	Response.Write(tpl.Content)		'输出内容
Set tpl = Nothing


Response.Write( "<br>运行速度" & RunTime() & "毫秒")

%>
