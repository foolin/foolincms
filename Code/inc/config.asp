<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit		'ǿ������
'On Error Resume Next		'�ݴ���
'=========================================================
' File Name��	config.asp
' Purpose��		ϵͳ�����ļ�
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Created on: 	2009-9-9 10:27:17
' Update on: 	2009-9-27 17:51:20
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================

Dim DBPATH		'Access���ݿ�·��
	DBPATH = "database/data.mdb"

Dim SITENAME		'��վ����
	SITENAME = "E�Ṥ����"

Dim HTTPURL		'��վ��ַǰ׺
	HTTPURL = "http://localhost"

Dim INSTALLDIR		'��վ��װĿ¼����Ŀ¼��Ϊ��/
	INSTALLDIR = "/eekku"

Dim SITEKEYWORDS		'��վ�ؼ���
	SITEKEYWORDS = "E������E��Cms��E�Ṥ����,www.eekku.com���������£�ling.liufu.org"

Dim TEMPLATEDIR		'��վģ��·�������磺default��ʾtemplate/default/
	TEMPLATEDIR = "default"

Dim ISHIDETEMPPATH		'�Ƿ�����ģ��·�����������Ӱ�������ٶ�
	ISHIDETEMPPATH = 0

Dim ISOPENGBOOK		'�Ƿ񿪷����ԣ�Ĭ�Ͽ���
	ISOPENGBOOK = 1

Dim ISAUDITGBOOK		'�Ƿ���Ҫ������ԣ���-1����-0
	ISAUDITGBOOK = 1

Dim ISCACHE		'�Ƿ񻺴棬�����ǣ����������������
	ISCACHE = 0

Dim CACHEFLAG		'�����־����������Ӣ����ĸ
	CACHEFLAG = "EekkuCms_"

Dim CACHETIME		'����ʱ�䣬Ĭ����60��
	CACHETIME = 60

Dim ISWEBLOG		'�Ƿ��¼��̨���������¼
	ISWEBLOG = 0

Dim LIMITIP		'����IP�����ö��Ž��зָ�
	LIMITIP = ""

Dim DIRTYWORDS		'�໰����,���ö��Ž��зָ�
	DIRTYWORDS = ""

%>

