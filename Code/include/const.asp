<%
TemplatePath = InstallDir & TemplatePath

Dim StartTime
	StartTime = Timer()

Dim SiteUrl		'��վ·��
	SiteUrl = Replace(HttpUrl & "/" & InstallDir & "/", "///", "/")
	SiteUrl = Replace(SiteUrl, "//", "/")
	SiteUrl = Replace(SiteUrl, "http:/", "http://")
	
Dim SkinUrl		'Ƥ��Ŀ¼Url���������images/css
	SkinUrl	= Replace(SiteUrl & "/" & TemplatePath & "/", "///", "/")
	SkinUrl	= Replace(SkinUrl, "//", "/")
	SkinUrl = Replace(SkinUrl, "http:/", "http://")
	
Dim PluginUrl		'���Ŀ¼Url
	PluginUrl	= SiteUrl & "plugin/"
	
Dim SitePathSplit	'·���ָ���
	SitePathSplit = " �� "

	
Dim StudioName	'�ٷ�
	StudioName = "E�Ṥ����"	
	
Dim StudioUrl	'�ٷ�Url
	StudioUrl = "http://www.eekku.com"
	
Dim SysName	'ϵͳ����
	SysName = "Eekku CMS"

Dim SysVersion	'ϵͳ�汾
	SysVersion = "v0.01.0718"
%>