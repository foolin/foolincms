<!--#include file="inc/admin.include.asp"-->
<%
'===========================================
'File Name：	Index.asp
'Purpose：		后台管理首页
'Auhtor: 		Foolin
'Create on:		2009-8-31 19:11:54
'Copyright:		E酷工作室(www.eekku.com)
'===========================================
Call ChkLogin()
Dim Act: Act = Request("action")
If LCase(Act) = "clearcache" Then
	Call ClearCache()
	Call MsgAndGo("更新缓存成功!", "index.asp")
End If 
%>
<!DOCTYPE html PUBLIC "=//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1=transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http=equiv="Content=Type" content="text/html; charset=gb2312" />
<title>网站管理-首页-<%=SYS%></title>
<link href="images/common.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.indexNav {
	border:solid 1px #BCE1FC;
	background:#FFF;
	font-size:14px;
	line-height:22px;
	background:#F0F8FF;
	padding:5px;
}
.indexNav table{
	border:solid 1px #AAD5FF;
	border-collapse:collapse;
	font-size:12px;
}
.indexNav td{
	border:solid 1px #AAD5FF;
	padding:3px 5px;
}
.indexNav ul{ margin:5px; clear:both; padding:0px; list-style:none;}
.indexNav li{
	margin:2px 5px;
	float:left;
}
.indexNav a{ color:#069; text-decoration:none;}
.indexNav a:hover{ color:#F00;}
.sysinfo{
	padding:5px;
	border-top:solid 1px #BCE1FC;
}
.sysinfo p{ margin:5px; padding:5px;}
-->
</style>
</head>

<body>
<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td colspan="2" id="topNav">
			<%Call TopNav("index")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li class="on"><a href="index.asp">管理首页</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">==== 相关操作 ====</li>
                 <li><a href="?action=clearcache">更新缓存</a></li>
                 <li><a href="../index.asp" target="_blank">前台首页</a></li>
                </ul>
                <%Call MyInfo()%>
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top" height="100%">
              <div class="content">
                  <div class="indexNav">
                  	<ul>
                    	<li><a href="modify_password.asp">修改密码</a> | </li>
                    	<li><a href="?action=clearcache">更新缓存</a> | </li>
                        <li><a href="admin_article.asp">管理文章</a> | </li>
                        <li><a href="admin_picture.asp">管理图片</a> | </li>
                        <li><a href="admin_guestbook.asp">管理留言 | </a></li>
                        <li><a href="../index.asp">前台首页</a></li>
                    </ul>
                 	<div class="clear"></div>
                    <div class="sysinfo">
                          &nbsp; <b style="color:#093;"><%=Session("AdminName")%>，欢迎您进入后台管理。</b><br /> 
                		<p>
                        
                	 &nbsp;网站名称:&nbsp;&nbsp;&nbsp; <%=SITE%><br />
                     &nbsp;网站网址:&nbsp;&nbsp;&nbsp; <a href="<%=SITEURL%>"><%=SITEURL%></a><br />
       
                            </p>
                           

                  </div>
                  
 					<table  class="list">
                      <tr height="18">
                        <th colspan="2" align="center">相关参数</th>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;服务器名：</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;服务器IP：</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;服务器端口：</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;服务器时间：</td>
                        <td class="td">&nbsp;<%=now%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;IIS版本：</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;脚本超时时间：</td>
                        <td class="td">&nbsp;<%=Server.ScriptTimeout%> 秒</td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;本文件路径：</td>
                        <td class="td">&nbsp;<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;服务器CPU数量：</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;服务器解译引擎：</td>
                        <td class="td">&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;服务器操作系统：</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("OS")%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;程序内核</td>
                        <td class="td">&nbsp;<%=syslink%></td>
                      </tr>
                     <tr height="18">
                        <td align="left" class="td">&nbsp;<%=SysName%>官方：</td>
                        <td class="td">&nbsp;<%=studio%></td>
                      </tr>
                     <tr height="18">
                        <td align="left" class="td">&nbsp;作者电子邮件：</td>
                        <td class="td">&nbsp; Foolin@126.com</td>
                      </tr>
                     <tr height="18">
                        <td align="left" class="td">&nbsp;作者主页：</td>
                        <td class="td">&nbsp;<a href="http://www.liufu.org/ling" target="_blank">http://www.liufu.org/ling</a></td>
                      </tr>
                     <tr height="18">
                        <td align="left" class="td">&nbsp;程序最新版本：</td>
                        <td class="td">&nbsp;<a href="http://code.google.com/p/foolincms/" target="_blank">http://code.google.com/p/foolincms/</a></td>
                      </tr>
                  </table>
                  
                   <p><b>版权声明：</b>
                            <br />
                            您现在使用的系统内核是 <%=syslink%>，本网站及其内容版权归[<%=Site%>]所有，程序版权归[<%=studio%>]所有。 </p>
                  
 				</div>
         </div>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
