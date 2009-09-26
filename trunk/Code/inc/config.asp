<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit		'强制声明
'On Error Resume Next		'容错处理
'=========================================================
' File Name：	config.asp
' Purpose：		系统配置文件
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Created on: 	2009-9-9 10:27:17
' Update on: 	2009-9-26 10:03:37
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================

Dim DBPATH		'Access数据库路径
	DBPATH = "database/data.mdb"

Dim SITENAME		'网站名称
	SITENAME = "E酷工作室"

Dim HTTPURL		'网站网址前缀
	HTTPURL = "http://localhost"

Dim INSTALLDIR		'网站安装目录，根目录则为：/
	INSTALLDIR = "/eekku"

Dim SITEKEYWORDS		'网站关键词
	SITEKEYWORDS = "E酷网，E酷Cms，E酷工作室,www.eekku.com，零星碎事，ling.liufu.org"

Dim TEMPLATEDIR		'网站模板路径，例如：default表示template/default/
	TEMPLATEDIR = "blog"

Dim ISHIDETEMPPATH		'是否隐藏模板路径，隐藏则会影响载入速度
	ISHIDETEMPPATH = 0

Dim ISCACHE		'是否缓存，建议是，减轻服务器负载量
	ISCACHE = 0

Dim CACHEFLAG		'缓存标志，可以任意英文字母
	CACHEFLAG = "EekkuCms_"

Dim CACHETIME		'缓存时间，默认是60分
	CACHETIME = 60

Dim ISWEBLOG		'是否记录后台管理操作记录
	ISWEBLOG = 0

Dim LIMITIP		'限制IP，多用逗号进行分割
	LIMITIP = ""

Dim DIRTYWORDS		'脏话过滤,多用逗号进行分割
	DIRTYWORDS = ""

%>

