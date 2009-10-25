<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/const.asp"-->
<%
Dim act : act = LCase(Request("action"))
Dim SUCCESS,FAIL
If act = "update" Then
	
	Call CreateConfig() '创建配置环境
	If Err Then FAIL = FAIL & "错误：" & Err.Description & "(" & Now() & ")<br />": Err.Clear
	
	If FAIL = "" Then
		SUCCESS = "恭喜，升级成功！请务必立刻把本升级文件(install/update.asp)删除！(" & Now() & ")"
	End If
	
End If

Function CreateConfig()
 	Dim strTemp, keyTab, keyEnter
	keyTab = Chr(9) & Chr(9)
	keyEnter = vbcrlf & vbcrlf
	Dim strSiteName: strSiteName = Replace(Trim(Request("Sitename")), """", "")
	
	'--------------增加的变量-----------------------
	Dim CODEPAGE: CODEPAGE = "936"		'页面编码65001|936
	Dim CHARSET: CHARSET = "GB2312"		'编码名称utf-8|gb2312
	Dim KEYWORDS: KEYWORDS = SiteKeywords	'网站关键词
	Dim DESCRIPTION: DESCRIPTION = SiteDesc	'网站描述
	'-------------增加的变量------------------------
	
	'系统信息
	strTemp =  Chr(60) & "%@LANGUAGE=""VBSCRIPT"" CODEPAGE="""& CODEPAGE &"""%" & Chr(62) & Chr(10)
	strTemp = strTemp & Chr(60) & "%" & Chr(10)
	strTemp = strTemp & "'Option Explicit" & keyTab & "'强制声明" & Chr(10)
	strTemp = strTemp & "On Error Resume Next" & keyTab & "'容错处理" & Chr(10)
	strTemp= strTemp & "Dim CODEPAGE: CODEPAGE = " & Chr(34) & CODEPAGE & Chr(34) & keyTab & "'页面编码65001|936" & Chr(10)
	strTemp= strTemp & "Dim CHARSET: CHARSET = " & Chr(34) & CHARSET & Chr(34)& keyTab & "'编码名称utf-8|gb2312" & Chr(10)
	strTemp = strTemp & "'=========================================================" & Chr(10)
	strTemp = strTemp & "' File Name：	config.asp" & Chr(10)
	strTemp = strTemp & "' Purpose：		系统配置文件" & Chr(10)
	strTemp = strTemp & "' Auhtor: 		Foolin" & Chr(10)
	strTemp = strTemp & "' E-mail: 		Foolin@126.com" & Chr(10)
	strTemp = strTemp & "' Created on: 	2009-9-9 10:27:17" & Chr(10)
	strTemp = strTemp & "' Update on: 	" & Now() & Chr(10)
	strTemp = strTemp & "' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved" & Chr(10)
	strTemp = strTemp & "'=========================================================" & keyEnter
	' DBPATH变量
	strTemp= strTemp & "Dim DBPATH" & keyTab & "'Access数据库路径" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DBPATH = " & Chr(34) & DBPATH & Chr(34) & keyEnter
	' SITENAME变量
	strTemp= strTemp & "Dim SITENAME" & keyTab & "'网站名称" & Chr(10) & Chr(9) 
	strTemp= strTemp & "SITENAME = " & Chr(34) & SITENAME & Chr(34) & keyEnter
	' HTTPURL变量
	strTemp= strTemp & "Dim HTTPURL" & keyTab & "'网站网址前缀" & Chr(10) & Chr(9) 
	strTemp= strTemp & "HTTPURL = " & Chr(34) & HTTPURL & Chr(34) & keyEnter
	' INSTALLDIR变量
	strTemp= strTemp & "Dim INSTALLDIR" & keyTab & "'网站安装目录，根目录则为：/" & Chr(10) & Chr(9) 
	strTemp= strTemp & "INSTALLDIR = " & Chr(34) & INSTALLDIR & Chr(34) & keyEnter
	' SITEKEYWORDS变量
	strTemp= strTemp & "Dim KEYWORDS" & keyTab & "'网站关键词" & Chr(10) & Chr(9) 
	strTemp= strTemp & "KEYWORDS = " & Chr(34) & KEYWORDS & Chr(34) & keyEnter
	' SITEDESC变量
	strTemp= strTemp & "Dim DESCRIPTION" & keyTab & "'网站描述" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DESCRIPTION = " & Chr(34) & DESCRIPTION & Chr(34) & keyEnter
	' TEMPLATEDIR变量
	strTemp= strTemp & "Dim TEMPLATEDIR" & keyTab & "'网站模板路径，例如：default表示template/default/" & Chr(10) & Chr(9) 
	strTemp= strTemp & "TEMPLATEDIR = " & Chr(34) & TEMPLATEDIR & Chr(34) & keyEnter
	' ISHIDETEMPPATH变量
	strTemp= strTemp & "Dim ISHIDETEMPPATH" & keyTab & "'是否隐藏模板路径，隐藏则会影响载入速度" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISHIDETEMPPATH = " & ISHIDETEMPPATH & keyEnter
	' ISOPENGBOOK变量
	strTemp= strTemp & "Dim ISOPENGBOOK" & keyTab & "'是否开放留言，默认开放" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISOPENGBOOK = " & ISOPENGBOOK & keyEnter
	' ISAUDITGBOOK变量
	strTemp= strTemp & "Dim ISAUDITGBOOK" & keyTab & "'是否需要审核留言，是-1，否-0" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISAUDITGBOOK = " & ISAUDITGBOOK & keyEnter
	' GBOOKTIME变量
	strTemp= strTemp & "Dim GBOOKTIME" & keyTab & "'允许留言最短时间间隔，单位秒，默认60秒" & Chr(10) & Chr(9) 
	strTemp= strTemp & "GBOOKTIME = " & GBOOKTIME & keyEnter
	' ISCACHE变量
	strTemp= strTemp & "Dim ISCACHE" & keyTab & "'是否缓存，建议是，减轻服务器负载量" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISCACHE = " & ISCACHE & keyEnter
	' CACHEFLAG变量
	strTemp= strTemp & "Dim CACHEFLAG" & keyTab & "'缓存标志，可以任意英文字母" & Chr(10) & Chr(9) 
	strTemp= strTemp & "CACHEFLAG = " & Chr(34) & CACHEFLAG & Chr(34) & keyEnter
	' CACHETIME变量
	strTemp= strTemp & "Dim CACHETIME" & keyTab & "'缓存时间，默认是60分" & Chr(10) & Chr(9) 
	strTemp= strTemp & "CACHETIME = " & CACHETIME & keyEnter
	' ISWEBLOG变量
	strTemp= strTemp & "Dim ISWEBLOG" & keyTab & "'是否记录后台管理操作记录" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISWEBLOG = " & ISWEBLOG & keyEnter
	' LIMITIP变量
	strTemp= strTemp & "Dim LIMITIP" & keyTab & "'限制IP，多用|进行分割" & Chr(10) & Chr(9) 
	strTemp= strTemp & "LIMITIP = " & Chr(34) & LIMITIP & Chr(34) & keyEnter
	' DIRTYWORDS变量
	strTemp= strTemp & "Dim DIRTYWORDS" & keyTab & "'脏话过滤,多用|进行分割" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DIRTYWORDS = " & Chr(34) & DIRTYWORDS & Chr(34) & keyEnter
	'标记结束
	strTemp = strTemp & "%" & Chr(62) & Chr(10)
	
	If CreateFile(strTemp, "../inc/config.asp") = True Then
		CreateConfig = True
	Else
		CreateConfig = False
	End If
	
End Function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>系统升级</title>
<style type="text/css">
<!--
body{
	font-family:Georgia, "Times New Roman", Times, serif;
	font-size:13px;
}
p{margin:5px;}

.wrapper{
	text-align:center;
}

.footer{
	line-height:22px;
	text-align:center;
	padding-top:30px;
}
.footer p{margin:5px;}

.title{
	font-size:24px;
	font-weight:bold;
	text-align:center;
	padding-top:20px;
	padding-bottom:20px;
}
.btn{
	text-align:center;
	padding:10px;
}
.btn input{
	padding:6px;
	font-size:14px;
}
.content {
	margin:0px auto;
	line-height:22px;
	height:400px;
	width:600px;
	padding:10px 20px;
	border:#EBEBEB 8px solid;
	overflow:auto;
	scrollbar-face-color:#EEE ;
	scrollbar-shadow-color: #ffffff; 
	scrollbar-highlight-color:#ffffff; 
	scrollbar-3dlight-color: #ffffff;  
	scrollbar-darkshadow-color: #ffffff; 
	scrollbar-track-color:#ffffff; 
	scrollbar-arrow-color: ffffff;
	background:#F9F9F9;
}
.red{ color:red;}
.green{ color:green;}
.blue{ color:blue;}
.gray{ color:gray;}
.result{
	font-size:16px;
	font-weight:bold;
}
-->
</style>
<script type="text/javascript">
function update(form){
	if(!confirm('请先备份好您网站的全部数据，然后再升级。\n\n我已经备份好所有数据了，现在进行升级?')){
		return;
	}
	form.submit();
}
</script>
</head>

<body>
<div class="wapper">

    	<div class="title">V1.0.3升级到EekkuCMS V1.0.5</div>
        <div class="content">
        	<b>注意事项</b>：<br />
            <ol>
        		<li>本次升级系统适合<span class="blue">EekkuCMS V1.0.3</span>升级到 <span class="blue">EekkuCMS V1.0.5</span>，请检查您的系统是否合适。</li>
                <li>本文件升级只是从新配置一下inc/config.asp文件，增加pagecode,charset,keywords,description变量，系统升级对模板不影响。</li>
                <li><span class="red">请先备份您网站的所有数据。</span></li>
                <li>系统检测您的系统版本为：<span class="blue"><%=Sys%></span></li>
                <li>升级完成之后，请立刻把<span class="blue">本升级文件（update.asp）</span>删除！</li>
                <li>如果有任何升级不成功或者升级出错，请到官方：http://www.eekku.com论坛进行反馈。</li>
           	</ol>
            <div class="result">
            	<div class="green"><%=success%></div>
                <div class="red"><%=fail%></div>
            </div>
        </div>
        <div class="btn">
        	<form action="update.asp" method="post">
            	<input type="hidden" name="action" value="update" />
                <input type="button" value="升级"  onclick="update(this.form);"/>
            </form>
        </div>
        
        <div class="footer">
                <p>版权所有 (c)2009-2010，E酷工作室 (www.eekku.com) 保留所有权利。 </p>
                <p>本系统由Foolin(负零)独立开发。Author: Foolin &nbsp;&nbsp; QQ: 970026999 E-mail: Foolin@126.com </p>
        </div>

</div>
</body>
</html>
