<!--#include file="../../inc/config.asp"-->
<!--#include file="../../inc/const.asp"-->
<!--#include file="../../inc/func_file.asp"-->
<!--#include file="../inc/admin.chklogin.asp"-->
<%
'=========================================================
' Purpose：		模板管理
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-3 16:40:50
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Call ChkLogin()
Dim extCode: extCode = Request("ext")
If LCase(extCode) = "js" Then extCode = "javascript"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<title>管理</title>
<script type="text/javascript" src="../inc/base.js"></script>
<script src="codepress/codepress.js" type="text/javascript"></script> 
<script type="text/javascript">
//li滑动效果 : 在<ul>中加入class='li-slide'即可。
function liEffect(){
	var nUl = document.getElementsByTagName("ul");
	for(var i = 0; i < nUl.length; i++){
			if(nUl[i].className != 'list') 	continue;
			var nLi = nUl[i].childNodes;
			for(var j = 0; j < nLi.length; j++){
					nLi[j].onmouseover = function(){ this.style.background = '#b5e28d'}    
					nLi[j].onmouseout = function(){ this.style.background = '#FFF';}
			}
	}
}
window.onload = liEffect; //网页载入运行

function saveCode(form){
	var input1 = document.createElement('input');
	input1.type = 'hidden';
	input1.name = 'content';
	input1.value = content1.getCode();
	form.insertBefore(input1);
	//alert(input1.value);
	form.submit(); 
}
</script>
<link type="text/css" href="codepress/languages/<%=extCode%>.css" rel="stylesheet" id="cp-lang-style" />
<style type="text/css">
<!--
body{
	font-size:14px;
}
.css_table{ background:#F4FAFF; border:#68B4FF 1px solid;}
a{text-decoration:none;color:#06F;}
a:hover{ text-decoration:none; color:#F00;}
.css_list{ background:#FFF;}
.list { margin: 0px; padding: 0px;}
.list li { list-style-image: none; list-style-type: none; float: left; text-align: center; padding:15px 15px 10px 15px; }
.list a{color:#000;}
.img { border: 1px solid #ECF7D9; padding: 20px; }
.img img { border: 0px; height:60px; width:60px; }
.img a{ }
.img a:hover{ background:#C00;}
.txt { width: 100px; white-space:nowrap; text-overflow:ellipsis; overflow: hidden; padding-top: 5px; }
-->
</style>
</head>
<body>
<%
Dim Url,Urli,Urlr
Url = "_WebFtp.asp"
Urli = Request("urli")
Urlr = Request("urlr")
if len(Urli) = 0 then Urli =  "/" & INSTALLDIR  & "/template/" &  Urli 
if len(Urlr) = 0 then Urlr = "/" & INSTALLDIR  & "/template/" 
Urli = replace(Urli,"//","/")
Urlr = replace(Urlr,"//","/")

Dim ArrowEXT,Href
ArrowEXT = "htm/html/txt/js/css"


If request("act") = "save" Then
	Dim SaveFile,SaveContent
	SaveFile = Request("file")
	SaveContent = Request("content")
	Call CreateFile(SaveContent,SaveFile)
End If 

Dim Fso: Set Fso = CreateObject("Scripting.FileSystemObject")

Dim Fileurl
Fileurl =  replace(request("file"),"//","/")
%>
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="css_table">
	<tr class="css_menu">
		<td><table width=100% border=0 cellpadding=4 cellspacing=0 class=css_main_table>
				<form name="frmSearch" method="post" action="<%=Url%>">
					<tr>
						<td class="css_main">
                        	<a href="../index.asp">后台管理首页</a> |
                        	<a href="<%=Url%>">WebFtp首页</a> |
							<a href="<%=Url%>?urli=<%=Urlr%>"><=返回上一级</a></td>
						<td class="css_search">&nbsp;</td>
					</tr>
				</form>
			</table></td>
	</tr>
	<tr>
		<td class='css_top'>URL: <%If Len(Fileurl) > 0 Then Response.write Fileurl Else Response.Write Urli%></td>
	</tr>
	<%
	Dim Root,F,Ext,Extarr,Extimg
	Set Root = Fso.GetFolder(Server.Mappath(Urli))
	%>
	<%If Len(Fileurl) > 0 Then%>
	<form id="form1" name="form1" method="post" action="<%=URL%>?act=save&file=<%=Server.HTMLEncode(Fileurl)%>">
	<tr>
		<td class='css_list'>
			<textarea name="content1" id="content1" class="codepress <%=extCode%> linenumbers-on" style="width:98.9%;height:440px;"><%=getfile(Server.HTMLEncode(Fileurl))%></textarea>
			<input type="button" name="save" onclick="saveCode(this.form)" value="保存该文件" /> </td>
	</tr>
	</form>
	 <%End If%>
	<tr>
		<td class='css_list'><ul class='list'>
			<%
			For Each F In Root.SubFolders
				Response.write "<li><a href=?urli=" & server.urlencode(urli) & "/" & server.urlencode(F.name) & "&urlr=" & server.urlencode(Urli) & "><div class='img'><img src=images/folder.gif></div><div class='txt'>" & F.Name & "</div></a></li>"
			Next
			For Each F In Root.files
				Extarr = Split(F.Name,".")
				Ext = LCase(Extarr(Ubound(Extarr)))
				Href = "#"
				If Instr(LCase(ArrowEXT),LCase(Ext)) > 0 Then Href="?urli=" & server.urlencode(urli) & "&urlr="&server.urlencode(urli) & "&file=" & server.urlencode(urli) & "/" & server.urlencode(F.name)
				if Instr("/png/jsp/asa/bat/rm/mp3/pdf/wma/rmvb/asp/html/htm/shtm/shtml/php/css/js/txt/gif/jpeg/jpg/bmp/swf/mdb/doc/xls/rar/zip/exe/xml/xsl/vbs/","/" & Ext & "/") > 0 Then Extimg = Ext Else Extimg = "file"
				Response.write "<li><a href=""" & href &"&ext="& Ext&""" title=""NAME: " & F.Name & "; TYPE: " & F.Type & ";""><div class='img'><img src=images/" & Extimg & ".gif></div><div class='txt'>" & F.Name & "</div></a></li>"
			Next
			%></ul>
		</td>
	</tr>
   
	<%
	set Root = nothing
	%>
</table>
</body>
</html>