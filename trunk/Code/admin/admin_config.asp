<!--#include file="inc/admin.include.asp"-->
<%
 ChkLogin()
 If LCase(Request("action")) = "update" Then
 	Dim strTemp, keyTab, keyEnter
	
	keyTab = Chr(9) & Chr(9)
	keyEnter = vbcrlf & vbcrlf
	If Len(Req("SiteName"))  = 0 Then Call MsgBox("网站名称不能为空","BACK")
	SITENAME = Replace(Req("SiteName"), chr(34), "'")
	If Len(Req("HttpUrl")) <> 0 Then
		If LCase(Left(Req("HttpUrl"),7)) <> "http://" Then
			Call MsgBox("网站网址不合法，必须为：http://开头","BACK")
		End If
		HTTPURL = Replace(Req("HttpUrl"), chr(34), "'")
	End If
	If Len(Req("InstallDir")) <> 0 Then
		INSTALLDIR = Replace(Req("InstallDir"), chr(34), "'")
	End If
	SITEKEYWORDS = Replace(Req("SiteKeywords"), chr(34), "'")
	If  Len(Req("IsHideTempPath")) <> 0 And Cint(Req("IsHideTempPath")) = 1 Then
		ISHIDETEMPPATH = 1
	Else
		ISHIDETEMPPATH = 0
	End If
	If Len(Req("TemplateDir")) = 0 Then
		Call MsgBox("模板目录不能为空","BACK")
	End If
	TEMPLATEDIR = Req("TemplateDir")
	If  Len(Req("IsCache")) <> 0 And Cint(Req("IsCache")) = 1 Then
		ISCACHE = 1
	Else
		ISCACHE = 0
	End If
	If Len(Req("CacheFlag")) = 0 Then
		CACHEFLAG = "EEKKU_COM"
	Else
		CACHEFLAG = Req("CacheFlag")
	End If
	If Len(Req("CacheTime")) = 0 Or Not IsNumeric(Req("CacheTime")) Then
		CACHETIME = 60
	Else
		CACHETIME = Req("CacheTime")
	End If
	LIMITIP =  Replace(Req("LimitIp"), chr(34), "'")
	DIRTYWORDS = Replace(Req("DirtyWords"), chr(34), "'")
	If  Len(Req("IsWebLog")) <> 0 And Cint(Req("IsWebLog")) = 1 Then
		ISCACHE = 1
	Else
		ISCACHE = 0
	End If
	
	'系统信息
	strTemp =  Chr(60) & "%@LANGUAGE=""VBSCRIPT"" CODEPAGE=""936""%" & Chr(62) & Chr(10)
	strTemp = strTemp & Chr(60) & "%" & Chr(10)
	strTemp = strTemp & "Option Explicit" & keyTab & "'强制声明" & Chr(10)
	strTemp = strTemp & "'On Error Resume Next" & keyTab & "'容错处理" & Chr(10)
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
	strTemp= strTemp & "Dim SITEKEYWORDS" & keyTab & "'网站关键词" & Chr(10) & Chr(9) 
	strTemp= strTemp & "SITEKEYWORDS = " & Chr(34) & SITEKEYWORDS & Chr(34) & keyEnter
	' TEMPLATEDIR变量
	strTemp= strTemp & "Dim TEMPLATEDIR" & keyTab & "'网站模板路径，例如：default表示template/default/" & Chr(10) & Chr(9) 
	strTemp= strTemp & "TEMPLATEDIR = " & Chr(34) & TEMPLATEDIR & Chr(34) & keyEnter
	' ISHIDETEMPPATH变量
	strTemp= strTemp & "Dim ISHIDETEMPPATH" & keyTab & "'是否隐藏模板路径，隐藏则会影响载入速度" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISHIDETEMPPATH = " & ISHIDETEMPPATH & keyEnter
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
	strTemp= strTemp & "Dim LIMITIP" & keyTab & "'限制IP" & Chr(10) & Chr(9) 
	strTemp= strTemp & "LIMITIP = " & Chr(34) & LIMITIP & Chr(34) & keyEnter
	' DIRTYWORDS变量
	strTemp= strTemp & "Dim DIRTYWORDS" & keyTab & "'脏话过滤" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DIRTYWORDS = " & Chr(34) & DIRTYWORDS & Chr(34) & keyEnter
	'标记结束
	strTemp = strTemp & "%" & Chr(62) & Chr(10)
	
	If CreateFile(strTemp, "../inc/config.asp") = True Then
		Call WebLog("配置网站成功!", "SESSION")
		Call MsgAndGo("配置网站成功!", "REFRESH")
	Else
		Call MsgBox("对不起，配置系统失败！\n\n请按照说明自行修改inc/config.asp配置文件！","BACK")
	End If
 End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>后台管理 - 配置管理 - <%=SYS%></title>
<link href="css/common.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.green {color:green;}
.red{ color:#F00;}
.blue{ color:blue;}
.gray{ color:gray;}
form{ margin:0px;}
table.form{
	width:100%;
	border:1px #88C4FF solid;
	background:#F0F8FF;
	border-collapse:collapse;
	line-height:30px;
}
table.form th{
	background:#6FB7FF;
	color:#FFF;
	border:#F0F8FF 1px solid;
	padding:4px;
	text-align:center;
	font-size:14px;
	line-height:20px;
}
table.form td{
	border:#ACD8FF 1px solid;
	padding:2px 5px;
	line-height:20px;
}
input{ background:#FFFFFF; padding:3px; border:#C4E1FF 1px solid;}
-->
</style>
</head>

<body>
<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td colspan="2" id="topNav">
			<%Call TopNav("config")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">管理首页</a></li>
                </ul>
                 <%Call MyInfo()%>
                <ul class="menu">
                 <li class="mTitle">--== 系统管理 ==--</li>
                 <li class="on"><a href="admin_config.asp">系统配置</a></li>
                 <li><a href="admin_user.asp">团队管理</a></li>
                 <li><a href="admin_mytag.asp">标签管理</a></li>
                 <li><a href="admin_diypage.asp">DIY页面管理</a></li>
                 <li><a href="admin_weblog.asp">操作记录管理</a></li>
                </ul>
                
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top" height="100%">
            	<div class="content">
                	<div class="status"> 您的位置：<a href="index.asp">管理首页</a> → 系统配置</div>
                    <form action="?action=update" id="form1" name="form1" method="post">
                        <table class="form" style="border:1px #88C4FF solid;">
                            <tr><th colspan="2">
                                系统配置
                            </th></tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">网站名称：</td>
                                <td><input type="text" name="SiteName" value="<%=SITENAME%>" style="width:250px;"/> <span class="gray">例如:E酷网</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">网站网址：</td>
                                <td><input type="text" name="HttpUrl" value="<%=HTTPURL%>" style="width:250px;"/> <span class="gray">例如：http://www.eekku.com （不能加目录）</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">安装目录：</td>
                                <td><input type="text" name="InstallDir" value="<%=INSTALLDIR%>" style="width:250px;"/> <span class="gray">安装目录 （前面加/，后面不用加/，不要加类似Http;//这样的绝对路径）</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">网站关键词：</td>
                                <td><textarea name="SiteKeywords" cols="50" rows="5"><%=SITEKEYWORDS%></textarea>
                                <span class="gray">网站关键词，多用逗号分隔。</span>
                                </td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">隐藏模板路径：</td>
                              <td>
                              		是<input type="radio" name="IsHideTempPath" value="1" <%If ISHIDETEMPPATH=1 THEN Echo("checked=""checked""")%> />
                                	否<input type="radio" name="IsHideTempPath" value="0" <%If ISHIDETEMPPATH=0 THEN Echo("checked=""checked""")%> />
                                 &nbsp;&nbsp;<span class="gray">隐藏路径可以保证模板安全，但会影响网页载入速度。</span>
                              </td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">模板目录：</td>
                                <td><input type="text" name="TemplateDir" value="<%=TEMPLATEDIR%>" style="width:250px;"/> <span class="gray">模板目录（例如:default表示目录template/default/）</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">是否缓存：</td>
                              <td>
                              		是<input type="radio" name="IsCache" value="1" <%If ISCACHE=1 THEN Echo("checked=""checked""")%> />
                                	否<input type="radio" name="IsCache" value="0" <%If ISCACHE=0 THEN Echo("checked=""checked""")%> />
                                 &nbsp;&nbsp;<span class="gray">隐藏路径可以保证模板安全，但会影响网页载入速度。</span>
                              </td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">缓存标志：</td>
                                <td><input type="text" name="CacheFlag" value="<%=CACHEFLAG%>" style="width:250px;"/> <span class="gray">缓存标志，如果同一台服务器安装两个CMS，则必须不同。</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">缓存时间：</td>
                                <td><input type="text" name="CacheTime" value="<%=CACHETIME%>" style="width:250px;"/> <span class="gray">缓存时间，默认0。</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">限制访问者IP：</td>
                                <td><textarea name="LimitIp" cols="50" rows="5"><%=LIMITIP%></textarea>
                                <span class="gray">限制访问者IP，多用逗号分隔。</span>
                                </td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">脏话过滤：</td>
                                <td><textarea name="DirtyWords" cols="50" rows="5"><%=DIRTYWORDS%></textarea>
                                <span class="gray">多用逗号分隔。</span>
                                </td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">记录管理操作：</td>
                                <td>                              		
                              		是<input type="radio" name="IsWebLog" value="1" <%If ISWebLog=1 THEN Echo("checked=""checked""")%> />
                                	否<input type="radio" name="IsWebLog" value="0" <%If ISWebLog=0 THEN Echo("checked=""checked""")%> />
                                 &nbsp;&nbsp;<span class="gray">选择是否记录后台管理操作记录。</span>
                                 </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="center">
                                    <input type="submit" class="btn" value="提交" />
                                    <input type="reset" class="btn" value="重置" />
                                </td>
                            </tr>
                        </table>
                    </form>
                    </div>
					<script type="text/javascript">
                    <!--
                    var oInputs = document.getElementsByTagName("input");
                    for(var i = 0; i < oInputs.length; i++){
                     if(oInputs.item(i).name != "")
                        oInputs.item(i).onmouseover = function(){
                            this.style.background='#FF0';
                            //this.style.borderColor = '#09F';
                            this.style.border = '#09F 2px solid';
                        };  
                        oInputs.item(i).onmouseout = function(){
                            this.style.background='#FFF';
                            //this.style.borderColor = '#C4E1FF';
                            this.style.border = '#C4E1FF 1px solid';
                        };
                    }
                    //-->
                    </script>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
