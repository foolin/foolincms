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
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网站管理-首页-<%=SYS%></title>
<link href="css/common.css" rel="stylesheet" type="text/css" />
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
                 <%Call MyInfo()%>
                <ul class="menu">
                 <li class="mTitle">--== 文章管理 ==--</li>
                 <li><a href="admin_article.asp">文章管理</a></li>
                 <li><a href="admin_artcolumn.asp">栏目管理</a></li>
                 
                 <li class="mTitle">--== 图片管理 ==--</li>
                 <li><a href="admin_picture.asp">图片管理</a></li>
                 <li><a href="admin_piccolumn.asp">栏目管理</a></li>
                 
                 <li class="mTitle">--== 事务管理 ==--</li>
                 <li><a href="admin_guestbook.asp">留言管理</a></li>
                 <li><a href="admin_comment.asp">评论管理</a></li>
                 <li><a href="admin_uploadfile.asp">上传文件管理</a></li>
                 
                 <li class="mTitle">--== 系统管理 ==--</li>
                 <li><a href="admin_config.asp">系统配置</a></li>
                 <li><a href="admin_user.asp">团队管理</a></li>
                 <li><a href="admin_mytag.asp">标签管理</a></li>
                 <li><a href="admin_diypage.asp">DIY页面管理</a></li>
                 <li><a href="admin_weblog.asp">操作记录管理</a></li>
                </ul>
                
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top" height="100%">
                    	 Foolin，欢迎你进入后台管理，你的等级是普通管理员。
                  <table style="width:85%; border:1px solid #93C9FF; margin:5px;">
                      <tr bgcolor="#E8F1FF" height=18>
                        <td colspan="2" align=center class="td2">服务器相关参数</td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;服务器名</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;服务器IP</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;服务器端口</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;服务器时间</td>
                        <td class="td">&nbsp;<%=now%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;IIS版本</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;脚本超时时间</td>
                        <td class="td">&nbsp;<%=Server.ScriptTimeout%> 秒</td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;本文件路径</td>
                        <td class="td">&nbsp;<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;服务器CPU数量</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;服务器解译引擎</td>
                        <td class="td">&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;服务器操作系统</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("OS")%></td>
                      </tr>
                  </table>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
