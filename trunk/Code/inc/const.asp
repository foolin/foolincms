<%
Dim SITEKEYWORDS	'��վ�ؼ��ʣ�����V1.0.3��ǰ�汾�����ԣ�
	SITEKEYWORDS = KEYWORDS

Dim SITEDESC		'��վ����������V1.0.3��ǰ�汾�����ԣ�
	SITEDESC = DESCRIPTION

Dim SITE
	SITE = "<a href=""" & SITEURL & """ target=""_blank"">" & SITENAME & "</a>"
Dim TEMPLATEPATH	'ģ��·��
	TemplatePath = "template/" & TemplateDir
	'TemplatePath = InstallDir & "/template/" & TemplateDir

Dim STARTTIME	'��¼��ʼִ��ʱ��
	STARTTIME = Timer()

Dim SITEURL		'��վ·��
	SITEURL = Replace(HTTPURL & "/" & INSTALLDIR & "/", "///", "/")
	SITEURL = Replace(SITEURL, "//", "/")
	SITEURL = Replace(SITEURL, "http:/", "http://")
	
Dim SKINURL		'Ƥ��Ŀ¼Url���������images/css
	SKINURL	= Replace(SITEURL & "/" & TEMPLATEPATH & "/", "///", "/")
	SKINURL	= Replace(SKINURL, "//", "/")
	SKINURL = Replace(SKINURL, "http:/", "http://")
	
Dim PLUGINURL		'���Ŀ¼Url
	PLUGINURL	= SITEURL & "plugins/"
	
Dim SITEPATHSPLIT	'·���ָ���
	SITEPATHSPLIT = " �� "

	
Dim STUDIONAME	'�ٷ�
	STUDIONAME = "Eekku Studio"	
	
Dim STUDIOURL	'�ٷ�Url
	STUDIOURL = "http://www.eekku.com"
	
Dim STUDIO
	STUDIO = "<a href=""" & STUDIOURL & """ target=""_blank"">" & STUDIONAME & "</a>"
	
Dim SYSNAME	'ϵͳ����
	SYSNAME = "EekkuCMS"

Dim SYSVERSION	'ϵͳ�汾
	SYSVERSION = " V1.0.4beta build091015"
	
Dim SYS	'ϵͳ
	SYS = SYSNAME & "  " & SYSVERSION
	
Dim SYSLINK	'ϵͳ����
	SYSLINK = "<a href=""" & STUDIOURL & """ target=""_blank"">" & SYS & "</a>"
	
%>