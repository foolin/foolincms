<!--#include file="inc/admin.include.asp"-->
<%
If Request("action") = "login" Then
	Dim strUsername, strPassword
	strUsername = Req("Username")
	strPassword = MD5(Req("Password"))
	If Len(strUsername) < 3 Then
		Call MsgBox("�û����ĳ��Ȳ�������3���ַ���", "REFRESH")
	End If
	If Request("ChkCode") <> Session("ChkCode") Then
		Call MsgBox("��֤�벻��ȷ��", "REFRESH")
	End If
	Dim Rs, strSql, rsUsername, rsPassword
		strSql = "Select * From [Admin] Where [Username]='"&strUserName&"' and [Password]='"&strPassword&"'"
	Set Rs = DB(strSql,3)
		If Rs.Eof Then
			Call WebLog("�û���["& strUsername &"]�������벻��ȷ��", strUsername)
			Call MsgBox("�û����������벻��ȷ","REFRESH")
		Else
			rsUsername = Rs("Username")
			rsPassword = Rs("Password")
		End If
		If strUsername <> rsUsername Then
			Call WebLog("�û���["& strUsername &"]����ȷ��", strUsername)
			Call MsgBox("�û�������ȷ","REFRESH")
		End If
		If strPassword <> rsPassword Then
			Call WebLog("[User:"& strUsername &"]���벻��ȷ��", strUsername)
			Call MsgBox("���벻��ȷ","REFRESH")
		End If
		If Rs("Level") < 0 Then
			Call WebLog("�û�[User:"& strUsername &"]�Ƕ����û�����¼ʧ�ܣ�", strUsername)	'���Ӽ�¼
			Call MsgBox("�����˺��Ѿ������ᣡ����ϵ����Ա��", "BACK")
		End If
		'��������¼��Ϣ
		Rs("LoginTime") = Now()
		Rs("LoginCount") = Rs("LoginCount") + 1
		Rs("LoginIP") = GetIP()
		Rs.Update
		Call WebLog("�û�[User:"& strUsername &"]��¼�ɹ���", strUsername)	'���Ӽ�¼
		Session("AdminName") = Rs("Username")		'����Session����
		Session("AdminLevel") = Rs("Level")
		Session.Timeout = 120
	Rs.Close
	Set Rs = Nothing
	Call ConnClose()
	Response.Redirect("index.asp")
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SiteName%> - Powered by www.eekku.com</title>
<style type="text/css">
<!--
.loginForm {
	margin:20px auto;
	width:300px;
	line-height:25px;
	font-size: 14px;
	border:#E3E3E3 5px solid;
	padding:5px;
	background:#F7F7F7;
	background:#F3F3F3;
	text-align:center;
}
.loginForm table{
	border-collapse:collapse;
}
.loginForm table td{
	border:#FFF 2px solid;
	padding:5px;
}

.input{
	border:#aaa dashed 1px;
	font-family:Tahoma, Geneva, sans-serif;
	font-size:15px;
	font-weight:bold;
	height:25px;
	line-height:25px;
	color:#090;
	padding:2px 5px;
}
.btn{
	line-height:22px;
	padding:3px 10px;
}

.txtL{ text-align:left;}
.txtR{ text-align:right;}
.title{ font-size:16px; font-weight:bold; color:#666;}
.footer{
	margin:5px auto;
	font-size:12px;
	text-align:center;
	line-height:22px;
	color:#666;
}
.footer a{color:#666;}
a {color:#000;text-decoration:none;}
a:hover{ color:#F00; text-decoration:underline;}
-->
</style>
</head>

<body>
	<div class="loginForm">
	<form action="" method="post">
        <table width="100%">
          <tr>
            <td colspan="2" class="title"><%=SiteName%>�����¼</td>
            <input type="hidden" name="action" value="login" />
          </tr>
          <tr>
            <td class="txtR">�û�����</td>
            <td class="txtL"><input name="Username" class="input" style="width:150px;" type="text" /></td>
          </tr>
          <tr>
            <td class="txtR">��&nbsp;&nbsp;�룺</td>
            <td class="txtL"><input name="Password" class="input" style="width:150px;"  type="password" /></td>
          </tr>
          <tr>
            <td class="txtR">��֤�룺</td>
            <td class="txtL"><input name="ChkCode" class="input"  style="width:100px;"  type="text" /> <img src="../inc/chkcode.asp" alt="��֤��,�������?����ˢ����֤��" style="cursor:pointer;" onclick="this.src='../inc/chkcode.asp?t='+Math.random()"/></td>
          </tr>
          <tr>
            <td colspan="2"><input type="submit" class="btn" value="��¼" />
            <input type="reset" class="btn" value="����" /></td>
          </tr>
        </table>
	</form>
    </div>
    <div class="footer">
    	 <br />
    	 &copy; 2009 <%=Studio%> All rights reserved. Powered by <%=SysLink%> <br />
    </div>
</body>
</html>
