<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit		'强制声明
'On Error Resume Next	'容错处理

Dim DBPath			'Access数据库路径
	DBpath = "database/data.mdb"
	
Dim SiteName		'网站名称
	SiteName = "E酷网"
	
	
Dim SiteKeywords	'网站关键词，多用逗号分隔。
	SiteKeywords = "E酷网,E酷工作室,CMS,eekku.com"
	
Dim HttpUrl			'网站网址前缀，前面要加http
	HttpUrl = "http://localhost/"
	
Dim InstallDir		'安装目录，后面不用加/
	InstallDir = "/eekku"
		
Dim TemplatePath	'模板路径，后面不用加/
	TemplatePath = "/template/default"

	
Dim IsHideTempPath	'是否隐藏模板路径，若隐藏路径，则相对页面载入速度会慢一些
	IsHideTempPath = 0

Dim IsCache	'是否缓存模板, 1表示缓存，0表示不缓存
	IsCache = 0

Dim CacheFlag		'缓存标志
	CacheFlag = "EEKKU"

Dim CacheTime		'缓存时间,单位为分
	CacheTime = 0
	
Dim LimitIP			'限制访问IP，多IP用|分隔
	LimitIP = "172.168.168.20|"

Dim DirtyWords		'非法（脏话）过滤，多词用|分隔
	DirtyWords = "江泽民|胡锦涛|温家宝|他妈的|操你妈|草你妈|妈逼|法轮功|李洪志|我操|我草"
%>