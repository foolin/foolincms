<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit		'强制声明
'On Error Resume Next	'容错处理

Dim DBPATH			'Access数据库路径
	DBPATH = "database/data.mdb"
	
Dim SITENAME		'网站名称
	SITENAME = "E酷网"
	
	
Dim SITEKEYWORDS	'网站关键词，多用逗号分隔。
	SITEKEYWORDS = "E酷网,E酷工作室,CMS,eekku.com"
	
Dim HTTPURL			'网站网址前缀，前面要加http
	HTTPURL	 = "http://localhost/"
	
Dim INSTALLDIR		'安装目录，后面不用加/
	INSTALLDIR = "/eekku"
		
Dim TEMPLATEPATH	'模板路径，后面不用加/
	TEMPLATEPATH = "/template/default"

	
Dim ISHIDETEMPPATH	'是否隐藏模板路径，若隐藏路径，则相对页面载入速度会慢一些
	ISHIDETEMPPATH = 0

Dim ISCACHE	'是否缓存模板, 1表示缓存，0表示不缓存
	ISCACHE = 0

Dim CACHEFLAG		'缓存标志
	CACHEFLAG = "EEKKU"

Dim CACHETIME		'缓存时间,单位为分
	CACHETIME = 0
	
Dim ISWEBLOG		'是否记录后台操作
	ISWEBLOG = 0
	
Dim LIMITIP			'限制访问IP，多IP用|分隔
	LIMITIP = "172.168.168.20|"

Dim DIRTYWORDS		'非法（脏话）过滤，多词用|分隔
	DIRTYWORDS = "江泽民|胡锦涛|温家宝|他妈的|操你妈|草你妈|妈逼|法轮功|李洪志|我操|我草"
%>