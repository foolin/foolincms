<!--#include file="inc/admin.include.asp"-->
<%
 ChkLogin()
 If LCase(Request("action")) = "update" Then
	Dim rRs
	Dim rUsername, rNickname, rOldPassword, rPassword
	rUsername = GetCookies("AdminName")
	rNickname = Req("fNickname")
	rOldPassword = Req("fOldPassword")
	If Req("fPassword") <> "" Then
		rPassword = MD5(Req("fPassword"))
	Else
		rPassword = ""
	End If
	If rOldPassword = "" Then Call MsgBox("����д������", "BACK")
	If rNickname = "" And rPassword = "" Then Call MsgBox("û�κ�������Ҫ�޸ģ�", "BACK")
	Set rRs = DB("SELECT Nickname, Password FROM Admin WHERE Username='" & rUsername & "' AND Password='" & MD5(rOldPassword) & "'", 3)
	If rRs.Eof Then rRs.Close: Set rRs = Nothing: Call MsgBox("�����벻��ȷ��" ,"BACK")
		rRs("Nickname") = rNickname
		If rPassword <> "" Then rRs("Password") = rPassword
	rRs.Update
	rRs.Close: Set rRs = Nothing
	Call WebLog("����Ա["& rUsername &"]�޸����ϳɹ���", "SESSION")
	Call MsgAndGo("�޸����ϳɹ���", "BACK")
 End If
 
 Dim objRs
 Dim strUsername, strNickname
 Set objRs = DB("SELECT Username,Nickname FROM Admin WHERE Username = '" & GetLogin("AdminName") & "'", 1)
  If objRs.Eof Then objRs.Close: Set objRs = Nothing: Call MsgBox("�ʺŷǷ��������µ�¼��" ,"logout.asp")
  strUsername = objRs("Username")
  strNickname = objRs("Nickname")
 objRs.Close: Set objRs = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>���� - �޸����� - Powered by eekku.com</title>
<link href="images/common.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="inc/base.js"></script>
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
			<%Call TopNav("password")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                 <%Call MyInfo()%> 
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top" height="100%">
            	<div class="content">
                	<div class="status"> ����λ�ã�<a href="index.asp">������ҳ</a> �� �޸�����</div>
                    <form action="?action=update" id="form1" name="form1" method="post" onsubmit="return chkForm();">
                        <table class="form" style="border:1px #88C4FF solid;">
                            <tr><th colspan="2">
                                �޸�����
                            </th></tr>
                            <tr>
                                <td align="right" width="15%">�û�����</td>
                                <td><input type="text" name="fUsername" value="<%=strUsername%>"  class="gray" readonly="readonly" style="width:250px;"/> <span class="blue">�����޸�</span></td>
                            </tr>
                            <tr>
                                <td align="right" width="15%">�ǳƣ�</td>
                                <td><input type="text" name="fNickname" value="<%=strNickname%>" style="width:250px;"/> <span class="red">�������� ��*����Ϊ�գ�</span></td>
                            </tr>
                            <tr>
                                <td align="right" width="15%">�����룺</td>
                                <td><input type="password" name="fOldPassword" id="fOldPassword" value="" style="width:250px;"/> <span class="red">��������ǰ�����루*���</span></td>
                            </tr>
                            <tr>
                                <td align="right" width="15%">�����룺</td>
                                <td><input type="password" name="fPassword" id="fPassword" value="" style="width:250px;"/> <span class="gray">������޸�����</span></td>
                            </tr>
                            <tr>
                                <td align="right" width="15%">�ظ������룺</td>
                                <td><input type="password" name="fRePassword" id="fRePassword" value="" style="width:250px;"/> </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="center">
                                    <input type="submit" class="btn" value="�ύ" />
                                    <input type="reset" class="btn" value="����" />
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
						
						//����
						function chkForm(){
							if( $("fNickname").value == ""){
								alert("�������ǳƣ�");
								$("fNickname").focus();
								return false;
							}
							if( $("fOldPassword").value == ""){
								alert("����������룡");
								$("fOldPassword").focus();
								return false;
							}
							if( $("fRePassword").value != $("fPassword").value){
								alert("���������������벻һ�£�");
								$("fPassword").focus();
								return false;
							}
							return true;
						}
                    //-->
                    </script>
					<script type="text/javascript" src="inc/slide-effect.js"></script>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
