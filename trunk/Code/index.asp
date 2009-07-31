<!--#include file="include/include.asp"-->
<%
Dim page: page = CPage(request("page"))	'当前页数
Dim SitePath: SitePath = IndexPath()	'当前路径

Dim tpl	'模板类实例
Set tpl = New TemplateClass
	tpl.Page = page					'设置当前页
	Call tpl.Load("index.html")		'载入模板
	Call tpl.Parser_Run()			'运行标签分析
	Response.Write(tpl.Content)		'输出内容
Set tpl = Nothing

Response.Write( FormatTime(Now(), "yy-mm-dd-ss"))

Response.Write( ArtPath(10))

Response.Write( PicPath(52))

Response.Write( "运行速度" & RunTime() & "毫秒")

%>
