<%
Dim TEMPLATEPATH	'模板路径
	TemplatePath = InstallDir & "/template/" & TemplateDir

Dim STARTTIME	'记录开始执行时间
	STARTTIME = Timer()

Dim SITEURL		'网站路径
	SITEURL = Replace(HTTPURL & "/" & INSTALLDIR & "/", "///", "/")
	SITEURL = Replace(SITEURL, "//", "/")
	SITEURL = Replace(SITEURL, "http:/", "http://")
	
Dim SKINURL		'皮肤目录Url，方便调用images/css
	SKINURL	= Replace(SITEURL & "/" & TEMPLATEPATH & "/", "///", "/")
	SKINURL	= Replace(SKINURL, "//", "/")
	SKINURL = Replace(SKINURL, "http:/", "http://")
	
Dim PLUGINURL		'插件目录Url
	PLUGINURL	= SITEURL & "plugin/"
	
Dim SITEPATHSPLIT	'路径分隔符
	SITEPATHSPLIT = " → "

	
Dim STUDIONAME	'官方
	STUDIONAME = "E酷工作室"	
	
Dim STUDIOURL	'官方Url
	STUDIOURL = "http://www.eekku.com"
	
Dim SYSNAME	'系统名称
	SYSNAME = "EekkuCMS"

Dim SYSVERSION	'系统版本
	SYSVERSION = "v0.01.0718"
	
Dim SYS	'系统
	SYS = SYSNAME & "  " & SYSVERSION
	
Dim SYSLINK	'系统连接
	SYSLINK = "<a href=" & STUDIOURL & " target=""_blank"">" & SYS & "</a>"
%>