<%
Dim StartTime
	StartTime = Timer()

Dim SiteUrl		'��վ·��
	SiteUrl = Replace(HttpUrl & "/" & InstallDir, "///", "/")
	SiteUrl = Replace(SiteUrl, "//", "/")
	SiteUrl = Replace(SiteUrl, "http:/", "http://")
	
Dim SkinUrl		'Ƥ��Ŀ¼Url���������images/css
	SkinUrl	= Replace(HttpUrl & "/" & TemplatePath & "/", "///", "/")
	SkinUrl	= Replace(SkinUrl, "//", "/")
	SkinUrl = Replace(SkinUrl, "http:/", "http://")
	
Dim SitePathSplit	'·���ָ���
	SitePathSplit = " �� "
	
Dim OfficialUrl	'�ٷ�Url
	OfficialUrl = "http://www.eekku.com"
	
	
Dim SysName	'ϵͳ����
	SysName = "Eekku CMS"

Dim SysVersion	'ϵͳ�汾
	SysVersion = "v0.01.0718"
%>