<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/const.asp"-->
<%
Dim act : act = LCase(Request("action"))
Dim SUCCESS,FAIL
If act = "update" Then
	'�����ݿ�����
	Dim ConnStr, Conn
	ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../" & DBPath)
	Set   Conn=Server.CreateObject("ADODB.Connection")  
	Conn.Open ConnStr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "���ݿ����ӳ����������ݿ������ļ��е����ݿ�������á�"
		Response.End
	End If
	Conn.execute("ALTER TABLE [ArtColumn] ADD [Sort] integer Default 0")
	If Err Then FAIL = FAIL & "����" & Err.Description & "(" & Now() & ")<br />": Err.Clear
	
	Conn.execute("UPDATE [ArtColumn] SET Sort = 0")
	If Err Then FAIL = FAIL & "����" & Err.Description & "(" & Now() & ")<br />": Err.Clear
	
	Conn.execute("ALTER TABLE [PicColumn] ADD [Sort] integer Default 0")
	If Err Then FAIL = FAIL & "����" & Err.Description & "(" & Now() & ")<br />": Err.Clear
	
	Conn.execute("UPDATE [PicColumn] SET Sort = 0")
	If Err Then FAIL = FAIL & "����" & Err.Description & "(" & Now() & ")<br />": Err.Clear
	
	If FAIL = "" Then
		SUCCESS = "��ϲ�������ɹ�����������̰ѱ������ļ�(install/update.asp)ɾ����(" & Now() & ")"
	End If
End If
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

    	<div class="title">E��CMS������EekkuCMS V1.0.3</div>
        <div class="content">
        	<b>ע������</b>��<br />
            <ol>
            	<li>�������վ������С��������ֱ�Ӱ�װ��Ȼ��ֱ��ʹ�þ�ģ�弴�ɡ�</li>
        		<li>��������ϵͳ�ʺ�<span class="blue">EekkuCMS V1.0.0</span>������ <span class="blue">EekkuCMS V1.0.3</span>����������ϵͳ�Ƿ���ʡ�</li>
                <li>���ļ�����ֻ�Ƕ����ݿ������ֶΣ����������뿴����˵����</li>
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
        	<form action="" method="post">
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
