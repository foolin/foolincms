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
	Call MsgAndGo("更新缓存成功!", "REFRESH")
End If 
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
                <ul class="menu">
                 <li class="mTitle">--== 相关操作 ==--</li>
                 <li><a href="?action=clearcache">更新缓存</a></li>
                </ul>
                <%Call MyInfo()%>
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
                  <div >
                  	<ul>
                    	<li><a href="?action=clearcache">更新缓存</a></li>
                    </ul>
                  </div>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
