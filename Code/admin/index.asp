<!--#include file="inc/admin.include.asp"-->
<%
'===========================================
'File Name��	Index.asp
'Purpose��		��̨������ҳ
'Auhtor: 		Foolin
'Create on:		2009-8-31 19:11:54
'Copyright:		E�Ṥ����(www.eekku.com)
'===========================================
Call ChkLogin()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��վ����-��ҳ-<%=SYS%></title>
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
                 
                 <li class="mTitle">--== ���¹��� ==--</li>
                 <li><a href="admin_article.asp">���¹���</a></li>
                 <li><a href="admin_artcolumn.asp">��Ŀ����</a></li>
                 
                 <li class="mTitle">--== ͼƬ���� ==--</li>
                 <li><a href="admin_picture.asp">ͼƬ����</a></li>
                 <li><a href="admin_piccolumn.asp">��Ŀ����</a></li>
                 
                 <li class="mTitle">--== ������� ==--</li>
                 <li><a href="admin_guestbook.asp">���Թ���</a></li>
                 <li><a href="admin_comment.asp">���۹���</a></li>
                 <li><a href="admin_uploadfile.asp">�ϴ��ļ�����</a></li>
                 
                 <li class="mTitle">--== ϵͳ���� ==--</li>
                 <li><a href="admin_config.asp">ϵͳ����</a></li>
                 <li><a href="admin_user.asp">�Ŷӹ���</a></li>
                 <li><a href="admin_mytag.asp">��ǩ����</a></li>
                 <li><a href="admin_diypage.asp">DIYҳ�����</a></li>
                 <li><a href="admin_weblog.asp">������¼����</a></li>
                </ul>
                
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top">
                Foolin����ӭ������̨������ĵȼ�����ͨ����Ա��
                </div>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
