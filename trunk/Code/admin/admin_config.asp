<!--#include file="inc/admin.include.asp"-->
<%
 Call ChkLogin()
 Call ChkPower("config","all") '���Ȩ��
 If LCase(Request("action")) = "update" Then
 	Dim strTemp, keyTab, keyEnter
	
	keyTab = Chr(9) & Chr(9)
	keyEnter = vbcrlf & vbcrlf
	If Len(Request("SiteName"))  = 0 Then Call MsgBox("��վ���Ʋ���Ϊ��","BACK")
	SITENAME = Replace(Request("SiteName"), chr(34), "'")
	If LCase(Request("mode")) = "auto" Then
		HTTPURL = "http://" & Request.ServerVariables("Http_Host")
		INSTALLDIR = GetInstallDir()
	Else
		If Len(Request("HttpUrl")) <> 0 Then
			If LCase(Left(Request("HttpUrl"),7)) <> "http://" Then
				Call MsgBox("��վ��ַ���Ϸ�������Ϊ��http://��ͷ","BACK")
			End If
			HTTPURL = Replace(Request("HttpUrl"), chr(34), "'")
		End If
		If Len(Request("InstallDir")) <> 0 Then
			INSTALLDIR = Replace(Request("InstallDir"), chr(34), "'")
		End If
	End If
	SITEKEYWORDS = Replace(Request("SiteKeywords"), chr(34), "'")
	SITEDESC= Replace(Request("SiteDesc"), chr(34), "'")
	If  Len(Request("IsHideTempPath")) <> 0 And Cint(Request("IsHideTempPath")) = 1 Then
		ISHIDETEMPPATH = 1
	Else
		ISHIDETEMPPATH = 0
	End If
	If Len(Request("TemplateDir")) = 0 Then
		Call MsgBox("ģ��Ŀ¼����Ϊ��","BACK")
	End If
	TEMPLATEDIR = Request("TemplateDir")
	If  Len(Request("IsOpenGbook")) <> 0 And Cint(Request("IsOpenGbook")) = 1 Then
		ISOPENGBOOK = 1
	Else
		ISOPENGBOOK = 0
	End If
	If  Len(Request("IsAuditGbook")) <> 0 And Cint(Request("IsAuditGbook")) = 1 Then
		ISAUDITGBOOK = 1
	Else
		ISAUDITGBOOK = 0
	End If
	If Len(Request("GbookTime")) = 0 Or Not IsNumeric(Request("GbookTime")) Then
		GBOOKTIME = 60
	Else
		GBOOKTIME = Request("GbookTime")
	End If
	If  Len(Request("IsCache")) <> 0 And Cint(Request("IsCache")) = 1 Then
		ISCACHE = 1
	Else
		ISCACHE = 0
	End If
	CACHEFLAG = Replace(Request("CacheFlag"), chr(34), "'")
	If Len(Request("CacheFlag")) = 0 Then
		CACHEFLAG = "EEKKU_"
	End If
	If Len(Request("CacheTime")) = 0 Or Not IsNumeric(Request("CacheTime")) Then
		CACHETIME = 60
	Else
		CACHETIME = Request("CacheTime")
	End If
	LIMITIP =  Replace(Request("LimitIp"), chr(34), "'")
	DIRTYWORDS = Replace(Request("DirtyWords"), chr(34), "'")
	If  Len(Request("IsWebLog")) <> 0 And Cint(Request("IsWebLog")) = 1 Then
		ISWEBLOG = 1
	Else
		ISWEBLOG = 0
	End If
	
	'ϵͳ��Ϣ
	strTemp =  Chr(60) & "%@LANGUAGE=""VBSCRIPT"" CODEPAGE=""936""%" & Chr(62) & Chr(10)
	strTemp = strTemp & Chr(60) & "%" & Chr(10)
	strTemp = strTemp & "'Option Explicit" & keyTab & "'ǿ������" & Chr(10)
	strTemp = strTemp & "On Error Resume Next" & keyTab & "'�ݴ���" & Chr(10)
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
	' SITEDESC����
	strTemp= strTemp & "Dim SITEDESC" & keyTab & "'��վ����" & Chr(10) & Chr(9) 
	strTemp= strTemp & "SITEDESC = " & Chr(34) & SITEDESC & Chr(34) & keyEnter
	' TEMPLATEDIR����
	strTemp= strTemp & "Dim TEMPLATEDIR" & keyTab & "'��վģ��·�������磺default��ʾtemplate/default/" & Chr(10) & Chr(9) 
	strTemp= strTemp & "TEMPLATEDIR = " & Chr(34) & TEMPLATEDIR & Chr(34) & keyEnter
	' ISHIDETEMPPATH����
	strTemp= strTemp & "Dim ISHIDETEMPPATH" & keyTab & "'�Ƿ�����ģ��·�����������Ӱ�������ٶ�" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISHIDETEMPPATH = " & ISHIDETEMPPATH & keyEnter
	' ISOPENGBOOK����
	strTemp= strTemp & "Dim ISOPENGBOOK" & keyTab & "'�Ƿ񿪷����ԣ�Ĭ�Ͽ���" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISOPENGBOOK = " & ISOPENGBOOK & keyEnter
	' ISAUDITGBOOK����
	strTemp= strTemp & "Dim ISAUDITGBOOK" & keyTab & "'�Ƿ���Ҫ������ԣ���-1����-0" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISAUDITGBOOK = " & ISAUDITGBOOK & keyEnter
	' GBOOKTIME����
	strTemp= strTemp & "Dim GBOOKTIME" & keyTab & "'�����������ʱ��������λ�룬Ĭ��60��" & Chr(10) & Chr(9) 
	strTemp= strTemp & "GBOOKTIME = " & GBOOKTIME & keyEnter
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
	strTemp= strTemp & "Dim LIMITIP" & keyTab & "'����IP������|���зָ�" & Chr(10) & Chr(9) 
	strTemp= strTemp & "LIMITIP = " & Chr(34) & LIMITIP & Chr(34) & keyEnter
	' DIRTYWORDS����
	strTemp= strTemp & "Dim DIRTYWORDS" & keyTab & "'�໰����,����|���зָ�" & Chr(10) & Chr(9) 
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
 
 
Function GetInstallDir()
	Dim strDir: strDir = Request.ServerVariables("Path_Info")
	strDir = Left(strDir,InStrRev(strDir,"/")-1)	'���ء�/��װĿ¼/admin��
	strDir = Left(strDir,InStrRev(strDir,"/")-1)	'���ء�/��װĿ¼��
	If Trim(strDir) = "" Then strDir = "/"
	GetInstallDir = strDir
End Function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>��̨���� - ���ù��� - Powered by eekku.com</title>
<link href="images/common.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.green {color:green;}
.red{ color:#F00;}
.blue{ color:blue;}
.gray{ color:gray; line-height:20px;}
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
	border:solid 1px #CCC;
	padding:5px 10px;
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
                            <%If INSTALLDIR <> GetInstallDir Then%>
                             <tr>
                                <td colspan="2">
                                	<span class="red" style="font-size:12px;">ע��:����վ���ð�װĿ¼Ϊ��<span class="blue"><%=INSTALLDIR%></span>��ϵͳ��⵽����ǰ��װĿ¼Ϊ<span class="blue"><%=GetInstallDir%></span>��</span>
                                 </td>
                            </tr>
                            <%End If%>
                            <tr>
                                <td align="right" width="100">��վ���ƣ�</td>
                                <td>
                                	<input type="text" name="SiteName" value="<%=SITENAME%>" style="width:250px;"/> <br /> <span class="gray">����:E����</span>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">��վ������</td>
                                <td>
                                    <input type="text" name="HttpUrl" value="<%=HTTPURL%>" style="width:250px;"/> <br /> <span class="gray">���磺http://www.eekku.com�����ܼ�Ŀ¼����</span>
                                    <%If HTTPURL <> ("http://"&Request.ServerVariables("Http_Host")) Then%>
                                        <span class="blue">��⵽����Ϊ��http://<%=Request.ServerVariables("Http_Host")%></span>
                                    <%End If%>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">��װĿ¼��</td>
                                <td>
                                	<input type="text" name="InstallDir" value="<%=INSTALLDIR%>" style="width:250px;"/> <br /> <span class="gray">ǰ��ӡ�/�������治�üӡ�/������Ŀ¼ֱ���á�/����</span>
                                    <%If INSTALLDIR <> GetInstallDir Then%>
                                    	<span class="blue">��⵽��װĿ¼Ϊ��<%=GetInstallDir%></span>
                                    <%End If%>
                                 </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">��վ�ؼ��ʣ�</td>
                                <td>
                                	<textarea name="SiteKeywords" cols="60" rows="3"><%=SITEKEYWORDS%></textarea>									
                               		<span class="gray">���ö��ŷָ���</span>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">��վ������</td>
                                <td><textarea name="SiteDesc" cols="60" rows="3"><%=SITEDESC%></textarea>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">ģ�壺</td>
                                <td>
                                <select name="TemplateDir">
                                	<option value="<%=TEMPLATEDIR%>"> => <%=TEMPLATEDIR%> <= </option>
                                <%
                                    Dim Fso: Set Fso = CreateObject("Scripting.FileSystemObject")
                                    Dim Root: Set Root = Fso.GetFolder(Server.Mappath("../template"))
                                    Dim F
                                    For Each F In Root.SubFolders
                                        Response.write "<option value=""" & F.Name & """>" & F.Name & "</option>"& chr(10) & chr(10) & chr(9)
                                    Next
                                %>
                                </select>
                                 <span class="gray">ģ��Ŀ¼������:default��ʾĿ¼template/default/��</span></td>
                            </tr>
                            <tr>
                                <td align="right" width="100">����ģ��·����</td>
                              <td>
                              		��<input type="radio" name="IsHideTempPath" value="1" <%If ISHIDETEMPPATH=1 THEN Echo("checked=""checked""")%> />
                                	��<input type="radio" name="IsHideTempPath" value="0" <%If ISHIDETEMPPATH=0 THEN Echo("checked=""checked""")%> />
                                  &nbsp;&nbsp;<span class="gray">����·�����Է�ֹ��������ģ�壬����Ӱ����ҳ�����ٶȡ�</span>
                              </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">�������ԣ�</td>
                              <td>
                              		��<input type="radio" name="IsOpenGbook" value="1" <%If ISOPENGBOOK=1 THEN Echo("checked=""checked""")%> />
                                	��<input type="radio" name="IsOpenGbook" value="0" <%If ISOPENGBOOK=0 THEN Echo("checked=""checked""")%> />
                                  &nbsp;&nbsp;<span class="gray">������ţ����οͿ������ԡ�</span>
                              </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">������ԣ�</td>
                              <td>
                              		��<input type="radio" name="IsAuditGbook" value="1" <%If ISAUDITGBOOK=1 THEN Echo("checked=""checked""")%> />
                                	��<input type="radio" name="IsAuditGbook" value="0" <%If ISAUDITGBOOK=0 THEN Echo("checked=""checked""")%> />
                                  &nbsp;&nbsp;<span class="gray">��:��ʾ��Ҫ������Բ���ʾ��</span>
                              </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">����ʱ������</td>
                                <td><input type="text" name="GbookTime" value="<%=GBOOKTIME%>" style="width:250px;"/> <span class="gray">�����������ʱ��������λ�룬Ĭ��60�롣</span></td>
                            </tr>
                            <tr>
                                <td align="right" width="100">�Ƿ񻺴棺</td>
                              <td>
                              		��<input type="radio" name="IsCache" value="1" <%If ISCACHE=1 THEN Echo("checked=""checked""")%> />
                                	��<input type="radio" name="IsCache" value="0" <%If ISCACHE=0 THEN Echo("checked=""checked""")%> />
                                  &nbsp;&nbsp;<span class="gray">�������������ҳ����ٶȡ�</span>
                              </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">�����־��</td>
                                <td><input type="text" name="CacheFlag" value="<%=CACHEFLAG%>" style="width:250px;"/>  <span class="gray">�����־�����ͬһ̨��������װ����CMS������벻ͬ��</span></td>
                            </tr>
                            <tr>
                                <td align="right" width="100">����ʱ�䣺</td>
                                <td><input type="text" name="CacheTime" value="<%=CACHETIME%>" style="width:250px;"/>  <span class="gray">����ʱ�䣬Ĭ��0��</span></td>
                            </tr>
                            <tr>
                                <td align="right" width="100">���Ʒ�����IP��</td>
                                <td><textarea name="LimitIp" cols="50" rows="5"><%=LIMITIP%></textarea>
                                <span class="gray">���Ʒ�����IP������|�ָ���</span>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">�໰���ˣ�</td>
                                <td><textarea name="DirtyWords" cols="50" rows="5"><%=DIRTYWORDS%></textarea>
                               <span class="gray">����|�ָ���</span>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" width="100">��¼���������</td>
                                <td>                              		
                              		��<input type="radio" name="IsWebLog" value="1" <%If ISWebLog=1 THEN Echo("checked=""checked""")%> />
                                	��<input type="radio" name="IsWebLog" value="0" <%If ISWebLog=0 THEN Echo("checked=""checked""")%> />
                                  &nbsp;&nbsp;<span class="gray">�ǣ���ʾ��¼������־��</span>
                                 </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="center" >
                                    <input type="submit" class="btn" value="����" />
                                    <input type="reset" class="btn" value="����" />
                                     <input type="button" class="btn" value="�Զ�����" onclick="onAutoConfig(this.form);" />
                                    
                                </td>
                            </tr>
                        </table>
                    </form>
                      <div class="blue" style="line-height:32px; padding:5px;">���������վ֮����ִ�������������inc/config.asp�ļ���</div>
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
                    var oTextAreas = document.getElementsByTagName("textarea");
                    for(var i = 0; i < oTextAreas.length; i++){
                     if(oTextAreas.item(i).name != "")
                        oTextAreas.item(i).onmouseover = function(){
                            this.style.background='#FF0';
                            //this.style.borderColor = '#09F';
                            this.style.border = '#09F 2px solid';
                        };  
                        oTextAreas.item(i).onmouseout = function(){
                            this.style.background='#FFF';
                            //this.style.borderColor = '#C4E1FF';
                            this.style.border = '#C4E1FF 1px solid';
                        };
                    }
					function onAutoConfig(form){
						if (confirm('ϵͳ���Զ�����[��վ����]��[��װĿ¼]������ѡ�����ѡ��䡣\n\nȷ���Զ����ã�')){	
							form.action  = '?action=update&mode=auto';
							form.submit();  
						}
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
