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
	Call MsgAndGo("���»���ɹ�!", "REFRESH")
End If 
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
                	<li class="on"><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== ��ز��� ==--</li>
                 <li><a href="?action=clearcache">���»���</a></li>
                </ul>
                <%Call MyInfo()%>
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top" height="100%">
                    	 Foolin����ӭ������̨������ĵȼ�����ͨ����Ա��
                  <table style="width:85%; border:1px solid #93C9FF; margin:5px;">
                      <tr bgcolor="#E8F1FF" height=18>
                        <td colspan="2" align=center class="td2">��������ز���</td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;��������</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;������IP</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("LOCAL_ADDR")%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;�������˿�</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;������ʱ��</td>
                        <td class="td">&nbsp;<%=now%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;IIS�汾</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;�ű���ʱʱ��</td>
                        <td class="td">&nbsp;<%=Server.ScriptTimeout%> ��</td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;���ļ�·��</td>
                        <td class="td">&nbsp;<%=server.mappath(Request.ServerVariables("SCRIPT_NAME"))%></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;������CPU����</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;��������������</td>
                        <td class="td">&nbsp;<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
                      </tr>
                      <tr bgcolor="#E8F1FF" height=18>
                        <td align=center class="td">&nbsp;����������ϵͳ</td>
                        <td class="td">&nbsp;<%=Request.ServerVariables("OS")%></td>
                      </tr>
                  </table>
                  <div >
                  	<ul>
                    	<li><a href="?action=clearcache">���»���</a></li>
                    </ul>
                  </div>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
