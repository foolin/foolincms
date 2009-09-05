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
                
                 <%Call MyInfo()%>
                 
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
            <td id="content" valign="top">
                Foolin，欢迎你进入后台管理，你的等级是普通管理员。
                </div>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
