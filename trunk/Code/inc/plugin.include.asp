<!--#include file="config.asp"-->
<!--#include file="const.asp"-->
<!--#include file="func_common.asp"-->
<!--#include file="func_cache.asp"-->
<!--#include file="func_file.asp"-->
<!--#include file="func_db.asp"-->
<!--#include file="func_sitepath.asp"-->
<!--#include file="class_pagelist.asp"-->
<!--#include file="class_template.asp"-->

<% 	
Dim ConnStr, Conn

'--------�����ݿ�-------------
ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(InstallDir & "/" & DBPath)
Set   Conn=Server.CreateObject("ADODB.Connection")  
Conn.Open ConnStr
If Err Then
	Err.Clear
	Set Conn = Nothing
	Response.Write "����վ���ò���ȷ������ȷ���ð�װĿ¼�����ݿ�·����"
	Response.End
End If
%>
