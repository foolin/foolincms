<!--#include file="inc/admin.include.asp"-->
<%
 ChkLogin()
 If LCase(Request("action")) = "update" Then
 	Dim strTemp, keyTab, keyEnter
	
	keyTab = Chr(9) & Chr(9)
	keyEnter = vbcrlf & vbcrlf
	If Len(Req("SiteName"))  = 0 Then Call MsgBox("��վ���Ʋ���Ϊ��","BACK")
	SITENAME = Replace(Req("SiteName"), chr(34), "'")
	If Len(Req("HttpUrl")) <> 0 Then
		If LCase(Left(Req("HttpUrl"),7)) <> "http://" Then
			Call MsgBox("��վ��ַ���Ϸ�������Ϊ��http://��ͷ","BACK")
		End If
		HTTPURL = Replace(Req("HttpUrl"), chr(34), "'")
	End If
	If Len(Req("InstallDir")) <> 0 Then
		INSTALLDIR = Replace(Req("InstallDir"), chr(34), "'")
	End If
	SITEKEYWORDS = Replace(Req("SiteKeywords"), chr(34), "'")
	If  Len(Req("IsHideTempPath")) <> 0 And Cint(Req("IsHideTempPath")) = 1 Then
		ISHIDETEMPPATH = 1
	Else
		ISHIDETEMPPATH = 0
	End If
	If Len(Req("TemplateDir")) = 0 Then
		Call MsgBox("ģ��Ŀ¼����Ϊ��","BACK")
	End If
	TEMPLATEDIR = Req("TemplateDir")
	If  Len(Req("IsCache")) <> 0 And Cint(Req("IsCache")) = 1 Then
		ISCACHE = 1
	Else
		ISCACHE = 0
	End If
	If Len(Req("CacheFlag")) = 0 Then
		CACHEFLAG = "EEKKU_COM"
	Else
		CACHEFLAG = Req("CacheFlag")
	End If
	If Len(Req("CacheTime")) = 0 Or Not IsNumeric(Req("CacheTime")) Then
		CACHETIME = 60
	Else
		CACHETIME = Req("CacheTime")
	End If
	LIMITIP =  Replace(Req("LimitIp"), chr(34), "'")
	DIRTYWORDS = Replace(Req("DirtyWords"), chr(34), "'")
	If  Len(Req("IsWebLog")) <> 0 And Cint(Req("IsWebLog")) = 1 Then
		ISCACHE = 1
	Else
		ISCACHE = 0
	End If
	
	'ϵͳ��Ϣ
	strTemp =  Chr(60) & "%@LANGUAGE=""VBSCRIPT"" CODEPAGE=""936""%" & Chr(62) & Chr(10)
	strTemp = strTemp & Chr(60) & "%" & Chr(10)
	strTemp = strTemp & "Option Explicit" & keyTab & "'ǿ������" & Chr(10)
	strTemp = strTemp & "'On Error Resume Next" & keyTab & "'�ݴ���" & Chr(10)
	strTemp = strTemp & "'=========================================================" & Chr(10)
	strTemp = strTemp & "' File Name��	config.asp" & Chr(10)
	strTemp = strTemp & "' Purpose��		ϵͳ�����ļ�" & Chr(10)
	strTemp = strTemp & "' Auhtor: 		Foolin" & Chr(10)
	strTemp = strTemp & "' E-mail: 		Foolin@126.com" & Chr(10)
	strTemp = strTemp & "' Created on: 	2009-9-9 10:27:17" & Chr(10)
	strTemp = strTemp & "' Update on: 	" & Now() & Chr(10)
	strTemp = strTemp & "' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved" & Chr(10)
	strTemp = strTemp & "'=========================================================" & keyEnter
	' DBPATH����
	strTemp= strTemp & "Dim DBPATH" & keyTab & "'Access���ݿ�·��" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DBPATH = " & Chr(34) & DBPATH & Chr(34) & keyEnter
	' SITENAME����
	strTemp= strTemp & "Dim SITENAME" & keyTab & "'��վ����" & Chr(10) & Chr(9) 
	strTemp= strTemp & "SITENAME = " & Chr(34) & SITENAME & Chr(34) & keyEnter
	' HTTPURL����
	strTemp= strTemp & "Dim HTTPURL" & keyTab & "'��վ��ַǰ׺" & Chr(10) & Chr(9) 
	strTemp= strTemp & "HTTPURL = " & Chr(34) & HTTPURL & Chr(34) & keyEnter
	' INSTALLDIR����
	strTemp= strTemp & "Dim INSTALLDIR" & keyTab & "'��վ��װĿ¼����Ŀ¼��Ϊ��/" & Chr(10) & Chr(9) 
	strTemp= strTemp & "INSTALLDIR = " & Chr(34) & INSTALLDIR & Chr(34) & keyEnter
	' SITEKEYWORDS����
	strTemp= strTemp & "Dim SITEKEYWORDS" & keyTab & "'��վ�ؼ���" & Chr(10) & Chr(9) 
	strTemp= strTemp & "SITEKEYWORDS = " & Chr(34) & SITEKEYWORDS & Chr(34) & keyEnter
	' TEMPLATEDIR����
	strTemp= strTemp & "Dim TEMPLATEDIR" & keyTab & "'��վģ��·�������磺default��ʾtemplate/default/" & Chr(10) & Chr(9) 
	strTemp= strTemp & "TEMPLATEDIR = " & Chr(34) & TEMPLATEDIR & Chr(34) & keyEnter
	' ISHIDETEMPPATH����
	strTemp= strTemp & "Dim ISHIDETEMPPATH" & keyTab & "'�Ƿ�����ģ��·�����������Ӱ�������ٶ�" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISHIDETEMPPATH = " & ISHIDETEMPPATH & keyEnter
	' ISCACHE����
	strTemp= strTemp & "Dim ISCACHE" & keyTab & "'�Ƿ񻺴棬�����ǣ����������������" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISCACHE = " & ISCACHE & keyEnter
	' CACHEFLAG����
	strTemp= strTemp & "Dim CACHEFLAG" & keyTab & "'�����־����������Ӣ����ĸ" & Chr(10) & Chr(9) 
	strTemp= strTemp & "CACHEFLAG = " & Chr(34) & CACHEFLAG & Chr(34) & keyEnter
	' CACHETIME����
	strTemp= strTemp & "Dim CACHETIME" & keyTab & "'����ʱ�䣬Ĭ����60��" & Chr(10) & Chr(9) 
	strTemp= strTemp & "CACHETIME = " & CACHETIME & keyEnter
	' ISWEBLOG����
	strTemp= strTemp & "Dim ISWEBLOG" & keyTab & "'�Ƿ��¼��̨���������¼" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISWEBLOG = " & ISWEBLOG & keyEnter
	' LIMITIP����
	strTemp= strTemp & "Dim LIMITIP" & keyTab & "'����IP" & Chr(10) & Chr(9) 
	strTemp= strTemp & "LIMITIP = " & Chr(34) & LIMITIP & Chr(34) & keyEnter
	' DIRTYWORDS����
	strTemp= strTemp & "Dim DIRTYWORDS" & keyTab & "'�໰����" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DIRTYWORDS = " & Chr(34) & DIRTYWORDS & Chr(34) & keyEnter
	'��ǽ���
	strTemp = strTemp & "%" & Chr(62) & Chr(10)
	
	If CreateFile(strTemp, "../inc/config.asp") = True Then
		Call WebLog("������վ�ɹ�!", "SESSION")
		Call MsgAndGo("������վ�ɹ�!", "REFRESH")
	Else
		Call MsgBox("�Բ�������ϵͳʧ�ܣ�\n\n�밴��˵�������޸�inc/config.asp�����ļ���","BACK")
	End If
 End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>��̨���� - ���ù��� - <%=SYS%></title>
<link href="css/common.css" rel="stylesheet" type="text/css" />
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
			<%Call TopNav("config")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                 <%Call MyInfo()%>
                <ul class="menu">
                 <li class="mTitle">--== ϵͳ���� ==--</li>
                 <li class="on"><a href="admin_config.asp">ϵͳ����</a></li>
                 <li><a href="admin_user.asp">�Ŷӹ���</a></li>
                 <li><a href="admin_mytag.asp">��ǩ����</a></li>
                 <li><a href="admin_diypage.asp">DIYҳ�����</a></li>
                 <li><a href="admin_weblog.asp">������¼����</a></li>
                </ul>
                
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top" height="100%">
            	<div class="content">
                	<div class="status"> ����λ�ã�<a href="index.asp">������ҳ</a> �� ϵͳ����</div>
                    <form action="?action=update" id="form1" name="form1" method="post">
                        <table class="form" style="border:1px #88C4FF solid;">
                            <tr><th colspan="2">
                                ϵͳ����
                            </th></tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">��վ���ƣ�</td>
                                <td><input type="text" name="SiteName" value="<%=SITENAME%>" style="width:250px;"/> <span class="gray">����:E����</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">��վ��ַ��</td>
                                <td><input type="text" name="HttpUrl" value="<%=HTTPURL%>" style="width:250px;"/> <span class="gray">���磺http://www.eekku.com �����ܼ�Ŀ¼��</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">��װĿ¼��</td>
                                <td><input type="text" name="InstallDir" value="<%=INSTALLDIR%>" style="width:250px;"/> <span class="gray">��װĿ¼ ��ǰ���/�����治�ü�/����Ҫ������Http;//�����ľ���·����</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">��վ�ؼ��ʣ�</td>
                                <td><textarea name="SiteKeywords" cols="50" rows="5"><%=SITEKEYWORDS%></textarea>
                                <span class="gray">��վ�ؼ��ʣ����ö��ŷָ���</span>
                                </td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">����ģ��·����</td>
                              <td>
                              		��<input type="radio" name="IsHideTempPath" value="1" <%If ISHIDETEMPPATH=1 THEN Echo("checked=""checked""")%> />
                                	��<input type="radio" name="IsHideTempPath" value="0" <%If ISHIDETEMPPATH=0 THEN Echo("checked=""checked""")%> />
                                 &nbsp;&nbsp;<span class="gray">����·�����Ա�֤ģ�尲ȫ������Ӱ����ҳ�����ٶȡ�</span>
                              </td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">ģ��Ŀ¼��</td>
                                <td><input type="text" name="TemplateDir" value="<%=TEMPLATEDIR%>" style="width:250px;"/> <span class="gray">ģ��Ŀ¼������:default��ʾĿ¼template/default/��</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">�Ƿ񻺴棺</td>
                              <td>
                              		��<input type="radio" name="IsCache" value="1" <%If ISCACHE=1 THEN Echo("checked=""checked""")%> />
                                	��<input type="radio" name="IsCache" value="0" <%If ISCACHE=0 THEN Echo("checked=""checked""")%> />
                                 &nbsp;&nbsp;<span class="gray">����·�����Ա�֤ģ�尲ȫ������Ӱ����ҳ�����ٶȡ�</span>
                              </td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">�����־��</td>
                                <td><input type="text" name="CacheFlag" value="<%=CACHEFLAG%>" style="width:250px;"/> <span class="gray">�����־�����ͬһ̨��������װ����CMS������벻ͬ��</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">����ʱ�䣺</td>
                                <td><input type="text" name="CacheTime" value="<%=CACHETIME%>" style="width:250px;"/> <span class="gray">����ʱ�䣬Ĭ��0��</span></td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">���Ʒ�����IP��</td>
                                <td><textarea name="LimitIp" cols="50" rows="5"><%=LIMITIP%></textarea>
                                <span class="gray">���Ʒ�����IP�����ö��ŷָ���</span>
                                </td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">�໰���ˣ�</td>
                                <td><textarea name="DirtyWords" cols="50" rows="5"><%=DIRTYWORDS%></textarea>
                                <span class="gray">���ö��ŷָ���</span>
                                </td>
                            </tr>
                            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                                <td align="right" width="15%">��¼���������</td>
                                <td>                              		
                              		��<input type="radio" name="IsWebLog" value="1" <%If ISWebLog=1 THEN Echo("checked=""checked""")%> />
                                	��<input type="radio" name="IsWebLog" value="0" <%If ISWebLog=0 THEN Echo("checked=""checked""")%> />
                                 &nbsp;&nbsp;<span class="gray">ѡ���Ƿ��¼��̨���������¼��</span>
                                 </td>
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
                    //-->
                    </script>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
