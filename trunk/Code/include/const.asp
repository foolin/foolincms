<%
Dim StartTime
	StartTime = Timer()

Dim SiteUrl		'网站路径
	SiteUrl = Replace(HttpUrl & "/" & InstallDir, "///", "/")
	SiteUrl = Replace(SiteUrl, "//", "/")
	SiteUrl = Replace(SiteUrl, "http:/", "http://")
	
Dim SkinUrl		'皮肤目录Url，方便调用images/css
	SkinUrl	= Replace(HttpUrl & "/" & TemplatePath & "/", "///", "/")
	SkinUrl	= Replace(SkinUrl, "//", "/")
	SkinUrl = Replace(SkinUrl, "http:/", "http://")
	
Dim SitePathSplit	'路径分隔符
	SitePathSplit = " → "
	
Dim OfficialUrl	'官方Url
	OfficialUrl = "http://www.eekku.com"
	
	
Dim SysName	'系统名称
	SysName = "Eekku CMS"

Dim SysVersion	'系统版本
	SysVersion = "v0.01.0718"
%>