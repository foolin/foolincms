<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/const.asp"-->
<%
Dim act : act = LCase(Request("action"))
Dim SUCCESS,FAIL
If act = "update" Then
	'打开数据库连接
	Dim ConnStr, Conn
	ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../" & DBPath)
	Set   Conn=Server.CreateObject("ADODB.Connection")  
	Conn.Open ConnStr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "数据库连接出错，请检查数据库连接文件中的数据库参数设置。"
		Response.End
	End If
	Conn.execute("ALTER TABLE [ArtColumn] ADD [Sort] integer Default 0")
	If Err Then FAIL = FAIL & "错误：" & Err.Description & "(" & Now() & ")<br />": Err.Clear
	
	Conn.execute("UPDATE [ArtColumn] SET Sort = 0")
	If Err Then FAIL = FAIL & "错误：" & Err.Description & "(" & Now() & ")<br />": Err.Clear
	
	Conn.execute("ALTER TABLE [PicColumn] ADD [Sort] integer Default 0")
	If Err Then FAIL = FAIL & "错误：" & Err.Description & "(" & Now() & ")<br />": Err.Clear
	
	Conn.execute("UPDATE [PicColumn] SET Sort = 0")
	If Err Then FAIL = FAIL & "错误：" & Err.Description & "(" & Now() & ")<br />": Err.Clear
	
	If FAIL = "" Then
		SUCCESS = "恭喜，升级成功！请务必立刻把本升级文件(install/update.asp)删除！(" & Now() & ")"
	End If
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>系统升级</title>
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
	if(!confirm('请先备份好您网站的全部数据，然后再升级。\n\n我已经备份好所有数据了，现在进行升级?')){
		return;
	}
	form.submit();
}
</script>
</head>

<body>
<div class="wapper">

    	<div class="title">E酷CMS升级到EekkuCMS V1.0.3</div>
        <div class="content">
        	<b>注意事项</b>：<br />
            <ol>
            	<li>如果您网站数据量小，建议您直接安装，然后直接使用旧模板即可。</li>
        		<li>本次升级系统适合<span class="blue">EekkuCMS V1.0.0</span>升级到 <span class="blue">EekkuCMS V1.0.3</span>，请检查您的系统是否合适。</li>
                <li>本文件升级只是对数据库增加字段，其余升级请看升级说明。</li>
                <li><span class="red">请先备份您网站的所有数据。</span></li>
                <li>系统检测您的系统版本为：<span class="blue"><%=Sys%></span></li>
                <li>升级完成之后，请立刻把<span class="blue">本升级文件（update.asp）</span>删除！</li>
                <li>如果有任何升级不成功或者升级出错，请到官方：http://www.eekku.com论坛进行反馈。</li>
           	</ol>
            <div class="result">
            	<div class="green"><%=success%></div>
                <div class="red"><%=fail%></div>
            </div>
        </div>
        <div class="btn">
        	<form action="" method="post">
            	<input type="hidden" name="action" value="update" />
                <input type="button" value="升级"  onclick="update(this.form);"/>
            </form>
        </div>
        
        <div class="footer">
                <p>版权所有 (c)2009-2010，E酷工作室 (www.eekku.com) 保留所有权利。 </p>
                <p>本系统由Foolin(负零)独立开发。Author: Foolin &nbsp;&nbsp; QQ: 970026999 E-mail: Foolin@126.com </p>
        </div>

</div>
</body>
</html>
