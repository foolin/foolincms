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

Dim Msg, Url, WaitTime
Msg = Request("msg")
Url = Request("jumpurl")
WaitTime = Request("time")

If Len(Msg) = 0 Then Msg = "未知信息"
If Len(Url) = 0 Then Url = Request.ServerVariables("HTTP_REFERER")
If UCase(Url)="BACK" Or UCase(Url)="REFRESH" Then Url = Request.ServerVariables("HTTP_REFERER")
If Len(WaitTime) = 0 Or Not IsNumeric(WaitTime) Or WaitTime < 0 Then WaitTime  = 3
If WaitTime = 0 Then Response.Redirect(Url)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网站管理-首页-<%=SYS%></title>
<link href="css/common.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.tipsBox {
	margin:5px;
	border:#A4D1FF 5px solid;
	padding:5px;
	text-align:center;
	line-height:22px;
	font-size:14px;
	background:#E8F3FF;
}
.tipsBox a{ text-decoration:underline; color:green;}
.tipsBox a:hover{ text-decoration:underline; color:red;}
.tips{ color:#F00; padding:5px; background:#FFF;}
.msg {font-weight:bold; padding:5px; font-size:16px;}
.timeTips{ color:#000; padding:5px; color:#333;}
#waitTime{ font-weight:bold; color:#F00;}
-->
</style>
<script type="text/javascript" src="inc/base.js"></script>
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

                 <%Call MyInfo()%>
                
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top">
            	<div class="tipsBox">
                	<div class="tips">
                        <div class="msg">
                            <%=Request("msg")%>
                        </div>
                        <div class="timeTips">
                        	<span id="waitTime">3</span>秒钟后即将跳转，请稍后...如果无法跳转，请<a href="<%=Url%>">点击这里</a>。
                            <script type="text/javascript">
							 <!--
							  var time = <%=WaitTime%>;
							  function waitTime() {
								   $("waitTime").innerHTML = time;
								   time = time - 1;
								   if(time == -1) this.top.location.href = '<%=Url%>';
								   window.setTimeout('waitTime()', 1000);
							  }
							  waitTime();
							  //-->
							</script>
                        </div>
                    </div>
                    
                </div>
                
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>