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

'--------打开数据库-------------
ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(InstallDir & "/" & DBPath)
Set   Conn=Server.CreateObject("ADODB.Connection")  
Conn.Open ConnStr
If Err Then
	Err.Clear
	Set Conn = Nothing
	Response.Write "你网站配置不正确，请正确配置安装目录和数据库路径。"
	Response.End
End If
%>
