<!--#include file="inc/admin.include.asp"-->
<%
 Call ChkLogin()
 Call ChkPower("template","all")
 Dim extCode: extCode = Request("ext")
 Dim MainStatus, SubStatus: MainStatus = "<a href='?'>管理模板</a>"
 If LCase(extCode) = "js" Then extCode = "javascript"
 Dim strTips: strTips = ""
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>后台管理 - 配置管理 - <%=SYS%></title>
<link href="images/common.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body{
	font-size:14px;
}
.css_table{ background:#F4FAFF; background:#F2F2F2; border:#999 1px solid;}
.css_list{ background:#FFF;}
.list { margin: 0px; padding: 0px;}
.list li { list-style-image: none; list-style-type: none; float: left; text-align: center; padding:15px 15px 10px 15px; }
.list a{color:#000;}
.img { border: 1px solid #ECF7D9; padding: 20px; }
.img img { border: 0px; height:60px; width:60px; }
.img a{ }
.img a:hover{ background:#C00;}
.txt { width: 100px; white-space:nowrap; text-overflow:ellipsis; overflow: hidden; padding-top: 5px; }
.btn{ padding:3px; margin:5px; margin-left:20px;}
-->
</style>
<script type="text/javascript" src="inc/base.js"></script>
<script type="text/javascript">
//li滑动效果 : 在<ul>中加入class='li-slide'即可。
function liEffect(){
	var nUl = document.getElementsByTagName("ul");
	for(var i = 0; i < nUl.length; i++){
			if(nUl[i].className != 'list') 	continue;
			var nLi = nUl[i].childNodes;
			for(var j = 0; j < nLi.length; j++){
					nLi[j].onmouseover = function(){ this.style.background = '#b5e28d'}    
					nLi[j].onmouseout = function(){ this.style.background = '#FFF';}
			}
	}
}
window.onload = liEffect; //网页载入运行
</script>
</head>

<body>
<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td id="topNav">
			<%Call TopNav("template")%>
        </td></tr>
        <tr>
            <td id="content" valign="top" height="100%">	
            	<div class="content">
                	    

<%
Dim Url,Urli,Urlr
Url = CStr(Request.ServerVariables("SCRIPT_NAME"))
Urli = Request("urli")
Urlr = Request("urlr")
If len(Urli) = 0 Then Urli =  "../template/" &  Urli 
If len(Urlr) = 0 Then Urlr = "../template/"
If InStr(Urli, "/template/") = 0 Then Urli = "../template/"
If InStr(Urli, "/template/") = 0 Then Urli = "../template/"
Urli = replace(Urli,"//","/")
Urlr = replace(Urlr,"//","/")

Dim ArrowEXT,Href
ArrowEXT = "htm/html/txt/js/css"


If request("act") = "save" Then
	Dim SaveFile,SaveContent
	SaveFile = Request("file")
	SaveContent = Request("content")
	Call CreateFile(SaveContent,SaveFile)
	Call WebLog("修改模板["& SaveFile &"]成功！", "SESSION")
	'Call MsgAndGo("修改自定义标签[id:"& id &"]成功！", "BACK")
	strTips = "温馨提示：保存文件成功！" & Now() & ""
End If 

Dim Fso: Set Fso = CreateObject("Scripting.FileSystemObject")

Dim Fileurl
Fileurl =  replace(request("file"),"//","/")
%>
<div class="status">您的位置：<a href="index.asp">管理首页</a> → <%=MainStatus%> → <%If Len(Fileurl) > 0 Then Response.write Fileurl Else Response.Write Urli%>  <a name="webftp" id="webftp"></a></div> 
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="css_table">
	<tr>
		<td class='css_top'>
        <a href='?'><img src="inc/webftp/images/home.gif" border="0" width="28"/>目录</a>
        |  <a href="#" onclick="history.go(-1);"><img src="inc/webftp/images/back.gif" border="0"/>后退</a>
        |  <a href="<%=Url%>?urli=<%=Urlr%>"><img src="inc/webftp/images/upper.gif" border="0"/>上一级</a>
        | 地址：<input type="text" value="<%If Len(Fileurl) > 0 Then Response.write Fileurl Else Response.Write Urli%>"  style="width:400px; height:18px; line-height:18px;" readonly="readonly" /> 
        <span style="color:#F00; font-weight:bold; padding:5px;"><%=strTips%></span>
        <br /></td>
	</tr>
	<%
	Dim Root,F,Ext,Extarr,Extimg
	Set Root = Fso.GetFolder(Server.Mappath(Urli))
	%>
	<%If Len(Fileurl) > 0 Then%>
	<form id="form1" name="form1" method="post" action="<%=URL%>?act=save&file=<%=Server.URLEncode(Fileurl)%>&ext=<%=Request("ext")%>#webftp" onsubmit="return confirm('确定保存文件？');">
    	<input type="hidden" name="urli" value="<%=Urlr%>" />
        <input type="hidden" name="urlr" value="<%=Urlr%>" />
	<tr>
		<td class='css_list'>
			<textarea name="content" id="content1" style="width:100%;height:550px; border:#999 1px solid;"><%=Server.HTMLEncode(getfile(Fileurl))%></textarea>
			<input class="btn" type="submit" value="保存文件"  /></td>
	</tr>
	</form>
	 <%End If%>
	<tr>
		<td class='css_list'><ul class='list'>
			<%
			For Each F In Root.SubFolders
				Response.write "<li><a href=""?urli=" & server.urlencode(urli) & "/" & server.urlencode(F.name) & "&urlr=" & server.urlencode(Urli) & "#webftp""><div class='img'><img src=""inc/webftp/images/folder.gif""></div><div class='txt'>" & F.Name & "</div></a></li>"& chr(10) & chr(10) & chr(9)
			Next
			For Each F In Root.files
				Extarr = Split(F.Name,".")
				Ext = LCase(Extarr(Ubound(Extarr)))
				Href = "#"
				If Instr(LCase(ArrowEXT),LCase(Ext)) > 0 Then Href="?urli=" & server.urlencode(urli) & "&urlr="&server.urlencode(urli) & "&file=" & server.urlencode(urli) & "/" & server.urlencode(F.name)
				if Instr("/png/jsp/asa/bat/rm/mp3/pdf/wma/rmvb/asp/html/htm/shtm/shtml/php/css/js/txt/gif/jpeg/jpg/bmp/swf/mdb/doc/xls/rar/zip/exe/xml/xsl/vbs/","/" & Ext & "/") > 0 Then Extimg = Ext Else Extimg = "file"
				Response.write "<li><a href=""" & href &"&ext="& Ext&"#webftp"" title=""NAME: " & F.Name & "; TYPE: " & F.Type & ";""><div class='img'><img src=""inc/webftp/images/" & Extimg & ".gif""></div><div class='txt'>" & F.Name & "</div></a></li>" & chr(10) & chr(10) & chr(9)
			Next
			%></ul>
		</td>
	</tr>
   
	<%
	set Root = nothing
	%>
</table>

					<div style="color:green; padding:10px;">温馨提示：如果要选择模板，请到【<a href="admin_config.asp">系统配置 → 模板目录 </a>】进行设置！</div>
               </div>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
