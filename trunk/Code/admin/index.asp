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
Dim Act: Act = Request("action")
If LCase(Act) = "clearcache" Then
	Call ClearCache()
	Call MsgAndGo("���»���ɹ�!", "index.asp")
End If 
%>
<!DOCTYPE html PUBLIC "=//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1=transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http=equiv="Content=Type" content="text/html; charset=gb2312" />
<title>��վ����-��ҳ-<%=SYS%></title>
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
                	<li class="on"><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">==== ��ز��� ====</li>
                 <li><a href="?action=clearcache">���»���</a></li>
                 <li><a href="../index.asp" target="_blank">ǰ̨��ҳ</a></li>
                </ul>
                <%Call MyInfo()%>
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top" height="100%">
              <div class="content">
                  <div class="indexNav">
                  	<ul>
                    	<li><a href="modify_password.asp">�޸�����</a> | </li>
                    	<li><a href="?action=clearcache">���»���</a> | </li>
                        <li><a href="admin_article.asp">��������</a> | </li>
                        <li><a href="admin_picture.asp">����ͼƬ</a> | </li>
                        <li><a href="admin_guestbook.asp">�������� | </a></li>
                        <li><a href="../index.asp">ǰ̨��ҳ</a></li>
                    </ul>
                 	<div class="clear"></div>
                    <div class="sysinfo">
                          &nbsp; <b style="color:#093;"><%=Session("AdminName")%>����ӭ�������̨����</b><br /> 
                		<p>
                        
                	 &nbsp;��վ����:&nbsp;&nbsp;&nbsp; <%=SITE%><br />
                     &nbsp;��վ��ַ:&nbsp;&nbsp;&nbsp; <a href="<%=SITEURL%>"><%=SITEURL%></a><br />
       
                            </p>
                           

                  </div>
                  
 					<table  class="list">
                      <tr height="18">
                        <th colspan="2" align="center">��ز���</th>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;����������</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;������IP��</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;�������˿ڣ�</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;������ʱ�䣺</td>
                        <td class="td">&nbsp;<%=now%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;IIS�汾��</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;�ű���ʱʱ�䣺</td>
                        <td class="td">&nbsp;<%=Server.ScriptTimeout%> ��</td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;���ļ�·����</td>
                        <td class="td">&nbsp;<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;������CPU������</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;�������������棺</td>
                        <td class="td">&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;����������ϵͳ��</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("OS")%></td>
                      </tr>
                      <tr height="18">
                        <td align="left" class="td">&nbsp;�����ں�</td>
                        <td class="td">&nbsp;<%=syslink%></td>
                      </tr>
                     <tr height="18">
                        <td align="left" class="td">&nbsp;<%=SysName%>�ٷ���</td>
                        <td class="td">&nbsp;<%=studio%></td>
                      </tr>
                     <tr height="18">
                        <td align="left" class="td">&nbsp;���ߵ����ʼ���</td>
                        <td class="td">&nbsp; Foolin@126.com</td>
                      </tr>
                     <tr height="18">
                        <td align="left" class="td">&nbsp;������ҳ��</td>
                        <td class="td">&nbsp;<a href="http://www.liufu.org/ling" target="_blank">http://www.liufu.org/ling</a></td>
                      </tr>
                     <tr height="18">
                        <td align="left" class="td">&nbsp;�������°汾��</td>
                        <td class="td">&nbsp;<a href="http://code.google.com/p/foolincms/" target="_blank">http://code.google.com/p/foolincms/</a></td>
                      </tr>
                  </table>
                  
                   <p><b>��Ȩ������</b>
                            <br />
                            ������ʹ�õ�ϵͳ�ں��� <%=syslink%>������վ�������ݰ�Ȩ��[<%=Site%>]���У������Ȩ��[<%=studio%>]���С� </p>
                  
 				</div>
         </div>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
