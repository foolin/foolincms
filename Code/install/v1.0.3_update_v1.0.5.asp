<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/const.asp"-->
<%
Dim act : act = LCase(Request("action"))
Dim SUCCESS,FAIL
If act = "update" Then
	
	Call CreateConfig() '�������û���
	If Err Then FAIL = FAIL & "����" & Err.Description & "(" & Now() & ")<br />": Err.Clear
	
	If FAIL = "" Then
		SUCCESS = "��ϲ�������ɹ�����������̰ѱ������ļ�(install/update.asp)ɾ����(" & Now() & ")"
	End If
	
End If

Function CreateConfig()
 	Dim strTemp, keyTab, keyEnter
	keyTab = Chr(9) & Chr(9)
	keyEnter = vbcrlf & vbcrlf
	Dim strSiteName: strSiteName = Replace(Trim(Request("Sitename")), """", "")
	
	'--------------���ӵı���-----------------------
	Dim CODEPAGE: CODEPAGE = "936"		'ҳ�����65001|936
	Dim CHARSET: CHARSET = "GB2312"		'��������utf-8|gb2312
	Dim KEYWORDS: KEYWORDS = SiteKeywords	'��վ�ؼ���
	Dim DESCRIPTION: DESCRIPTION = SiteDesc	'��վ����
	'-------------���ӵı���------------------------
	
	'ϵͳ��Ϣ
	strTemp =  Chr(60) & "%@LANGUAGE=""VBSCRIPT"" CODEPAGE="""& CODEPAGE &"""%" & Chr(62) & Chr(10)
	strTemp = strTemp & Chr(60) & "%" & Chr(10)
	strTemp = strTemp & "'Option Explicit" & keyTab & "'ǿ������" & Chr(10)
	strTemp = strTemp & "On Error Resume Next" & keyTab & "'�ݴ���" & Chr(10)
	strTemp= strTemp & "Dim CODEPAGE: CODEPAGE = " & Chr(34) & CODEPAGE & Chr(34) & keyTab & "'ҳ�����65001|936" & Chr(10)
	strTemp= strTemp & "Dim CHARSET: CHARSET = " & Chr(34) & CHARSET & Chr(34)& keyTab & "'��������utf-8|gb2312" & Chr(10)
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
	strTemp= strTemp & "Dim KEYWORDS" & keyTab & "'��վ�ؼ���" & Chr(10) & Chr(9) 
	strTemp= strTemp & "KEYWORDS = " & Chr(34) & KEYWORDS & Chr(34) & keyEnter
	' SITEDESC����
	strTemp= strTemp & "Dim DESCRIPTION" & keyTab & "'��վ����" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DESCRIPTION = " & Chr(34) & DESCRIPTION & Chr(34) & keyEnter
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
		CreateConfig = True
	Else
		CreateConfig = False
	End If
	
End Function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>ϵͳ����</title>
<style type="text/css">
<!--
body{
	font-family:Georgia, "Times New Roman", Times, serif;
	font-size:13px;
}
p{margin:5px;}

.wrapper{
	text-align:center;
}

.footer{
	line-height:22px;
	text-align:center;
	padding-top:30px;
}
.footer p{margin:5px;}

.title{
	font-size:24px;
	font-weight:bold;
	text-align:center;
	padding-top:20px;
	padding-bottom:20px;
}
.btn{
	text-align:center;
	padding:10px;
}
.btn input{
	padding:6px;
	font-size:14px;
}
.content {
	margin:0px auto;
	line-height:22px;
	height:400px;
	width:600px;
	padding:10px 20px;
	border:#EBEBEB 8px solid;
	overflow:auto;
	scrollbar-face-color:#EEE ;
	scrollbar-shadow-color: #ffffff; 
	scrollbar-highlight-color:#ffffff; 
	scrollbar-3dlight-color: #ffffff;  
	scrollbar-darkshadow-color: #ffffff; 
	scrollbar-track-color:#ffffff; 
	scrollbar-arrow-color: ffffff;
	background:#F9F9F9;
}
.red{ color:red;}
.green{ color:green;}
.blue{ color:blue;}
.gray{ color:gray;}
.result{
	font-size:16px;
	font-weight:bold;
}
-->
</style>
<script type="text/javascript">
function update(form){
	if(!confirm('���ȱ��ݺ�����վ��ȫ�����ݣ�Ȼ����������\n\n���Ѿ����ݺ����������ˣ����ڽ�������?')){
		return;
	}
	form.submit();
}
</script>
</head>

<body>
<div class="wapper">

    	<div class="title">V1.0.3������EekkuCMS V1.0.5</div>
        <div class="content">
        	<b>ע������</b>��<br />
            <ol>
        		<li>��������ϵͳ�ʺ�<span class="blue">EekkuCMS V1.0.3</span>������ <span class="blue">EekkuCMS V1.0.5</span>����������ϵͳ�Ƿ���ʡ�</li>
                <li>���ļ�����ֻ�Ǵ�������һ��inc/config.asp�ļ�������pagecode,charset,keywords,description������ϵͳ������ģ�岻Ӱ�졣</li>
                <li><span class="red">���ȱ�������վ���������ݡ�</span></li>
                <li>ϵͳ�������ϵͳ�汾Ϊ��<span class="blue"><%=Sys%></span></li>
                <li>�������֮�������̰�<span class="blue">�������ļ���update.asp��</span>ɾ����</li>
                <li>������κ��������ɹ��������������뵽�ٷ���http://www.eekku.com��̳���з�����</li>
           	</ol>
            <div class="result">
            	<div class="green"><%=success%></div>
                <div class="red"><%=fail%></div>
            </div>
        </div>
        <div class="btn">
        	<form action="update.asp" method="post">
            	<input type="hidden" name="action" value="update" />
                <input type="button" value="����"  onclick="update(this.form);"/>
            </form>
        </div>
        
        <div class="footer">
                <p>��Ȩ���� (c)2009-2010��E�Ṥ���� (www.eekku.com) ��������Ȩ���� </p>
                <p>��ϵͳ��Foolin(����)����������Author: Foolin &nbsp;&nbsp; QQ: 970026999 E-mail: Foolin@126.com </p>
        </div>

</div>
</body>
</html>
