<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'Option Explicit		'强制声明
On Error Resume Next		'容错处理
'=========================================================
' File Name：	config.asp
' Purpose：		系统配置文件
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Created on: 	2009-9-9 10:27:17
' Update on: 	2009-10-13 0:00:17
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================

Dim DBPATH		'Access数据库路径
	DBPATH = "database/Fl28#Ek_7348D432AF.mdb"

Dim SITENAME		'网站名称
	SITENAME = "E酷科技网"

Dim HTTPURL		'网站网址前缀
	HTTPURL = "http://localhost"

Dim INSTALLDIR		'网站安装目录，根目录则为：/
	INSTALLDIR = "/eekku"

Dim SITEKEYWORDS		'网站关键词
	SITEKEYWORDS = "E酷网，E酷Cms，E酷工作室,www.eekku.com，零星碎事，ling.liufu.org"

Dim SITEDESC		'网站描述
	SITEDESC = "E酷Cms(EekkuCMS)是一种小型站点内容管理系统，内含文章、图片、留言等基本功能，简单且方便使用！"

Dim TEMPLATEDIR		'网站模板路径，例如：default表示template/default/
	TEMPLATEDIR = "blog"

Dim ISHIDETEMPPATH		'是否隐藏模板路径，隐藏则会影响载入速度
	ISHIDETEMPPATH = 0

Dim ISOPENGBOOK		'是否开放留言，默认开放
	ISOPENGBOOK = 1

Dim ISAUDITGBOOK		'是否需要审核留言，是-1，否-0
	ISAUDITGBOOK = 0

Dim GBOOKTIME		'允许留言最短时间间隔，单位秒，默认60秒
	GBOOKTIME = 60

Dim ISCACHE		'是否缓存，建议是，减轻服务器负载量
	ISCACHE = 0

Dim CACHEFLAG		'缓存标志，可以任意英文字母
	CACHEFLAG = "Eekku_"

Dim CACHETIME		'缓存时间，默认是60分
	CACHETIME = 60

Dim ISWEBLOG		'是否记录后台管理操作记录
	ISWEBLOG = 1

Dim LIMITIP		'限制IP，多用|进行分割
	LIMITIP = "127.0.0.2|127.0.0.3"

Dim DIRTYWORDS		'脏话过滤,多用|进行分割
	DIRTYWORDS = "fuck|sex"

%>

