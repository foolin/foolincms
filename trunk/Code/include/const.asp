<%
TemplatePath = InstallDir & TemplatePath

Dim StartTime
	StartTime = Timer()

Dim SiteUrl		'网站路径
	SiteUrl = Replace(HttpUrl & "/" & InstallDir & "/", "///", "/")
	SiteUrl = Replace(SiteUrl, "//", "/")
	SiteUrl = Replace(SiteUrl, "http:/", "http://")
	
Dim SkinUrl		'皮肤目录Url，方便调用images/css
	SkinUrl	= Replace(SiteUrl & "/" & TemplatePath & "/", "///", "/")
	SkinUrl	= Replace(SkinUrl, "//", "/")
	SkinUrl = Replace(SkinUrl, "http:/", "http://")
	
Dim PluginUrl		'插件目录Url
	PluginUrl	= SiteUrl & "plugin/"
	
Dim SitePathSplit	'路径分隔符
	SitePathSplit = " → "

	
Dim StudioName	'官方
	StudioName = "E酷工作室"	
	
Dim StudioUrl	'官方Url
	StudioUrl = "http://www.eekku.com"
	
Dim SysName	'系统名称
	SysName = "Eekku CMS"

Dim SysVersion	'系统版本
	SysVersion = "v0.01.0718"
%>