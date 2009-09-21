<!--#include file="inc/admin.include.asp"-->
<%
 ChkLogin()
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
<link href="css/common.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body{
	font-size:14px;
}
.css_table{ background:#F4FAFF; background:#F2F2F2; border:#68B4FF 1px solid;}
.css_list{ background:#FFF;}
.list { margin: 0px; padding: 0px;}
.list li { list-style-image: none; list-style-type: none; float: left; text-align: center; padding:15px 15px 10px 15px; }
.list a{color:#000;}
.img { border: 1px solid #ECF7D9; padding: 20px; }
.img img { border: 0px; height:60px; width:60px; }
.img a{ }
.img a:hover{ background:#C00;}
.txt { width: 100px; white-space:nowrap; text-overflow:ellipsis; overflow: hidden; padding-top: 5px; }
-->
</style>
<script type="text/javascript" src="inc/base.js"></script>
<script src="webftp/codepress/codepress.js" type="text/javascript"></script> 
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

function saveCode(form){
	var input1 = document.createElement('input');
	input1.type = 'hidden';
	input1.name = 'content';
	input1.value = content1.getCode();
	form.insertBefore(input1);
	//alert(input1.value);
	form.submit(); 
}
</script>
</head>

<body>
<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td id="topNav">
			<%Call TopNav("webftp")%>
        </td></tr>
        <tr>
            <td id="content" valign="top" height="100%">	
            	<div class="content">
                	    

<%
Dim Url,Urli,Urlr
Url = CStr(Request.ServerVariables("SCRIPT_NAME"))
Urli = Request("urli")
Urlr = Request("urlr")
if len(Urli) = 0 then Urli =  "/" & INSTALLDIR  & "/template/" &  Urli 
if len(Urlr) = 0 then Urlr = "/" & INSTALLDIR  & "/template/" 
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
	strTips = "温馨提示：保存文件成功！" & Now()
End If 

Dim Fso: Set Fso = CreateObject("Scripting.FileSystemObject")

Dim Fileurl
Fileurl =  replace(request("file"),"//","/")
%>
<div class="status"> <a name="webftp" id="webftp"></a>您的位置：<a href="index.asp">管理首页</a> → <%=MainStatus%> → <%If Len(Fileurl) > 0 Then Response.write Fileurl Else Response.Write Urli%>  </div> 
<table width="100%" border="0" cellpadding="1" cellspacing="1" class="css_table">
	<tr class="css_menu">
		<td>
		</td>
	</tr>
	<tr>
		<td class='css_top'>
        <a href='?'><img src="webftp/images/home.gif" border="0" width="28"/>首页</a>
        |  <a href="#" onclick="history.go(-1);"><img src="webftp/images/back.gif" border="0"/>后退</a>
        |  <a href="<%=Url%>?urli=<%=Urlr%>"><img src="webftp/images/upper.gif" border="0"/>上一级</a>
        | 地址：<input type="text" value="当前目录: <%If Len(Fileurl) > 0 Then Response.write Fileurl Else Response.Write Urli%>"  style="width:450px; height:18px; line-height:18px;" readonly="readonly" /> 
        <span style="color:#F00; font-weight:bold;"><%=strTips%></span>
        <br /></td>
	</tr>
	<%
	Dim Root,F,Ext,Extarr,Extimg
	Set Root = Fso.GetFolder(Server.Mappath(Urli))
	%>
	<%If Len(Fileurl) > 0 Then%>
	<form id="form1" name="form1" method="post" action="<%=URL%>?act=save&file=<%=Server.HTMLEncode(Fileurl)%>&ext=<%=Request("ext")%>">
	<tr>
		<td class='css_list'>
			<textarea name="content1" id="content1" class="codepress <%=extCode%> linenumbers-on" style="width:99%;height:550px; border:#09F 1px solid;"><%=getfile(Server.HTMLEncode(Fileurl))%></textarea>
			<input type="button" name="save" onclick="saveCode(this.form)" value="保存该文件" /> </td>
	</tr>
	</form>
	 <%End If%>
	<tr>
		<td class='css_list'><ul class='list'>
			<%
			For Each F In Root.SubFolders
				Response.write "<li><a href=""?urli=" & server.urlencode(urli) & "/" & server.urlencode(F.name) & "&urlr=" & server.urlencode(Urli) & "#webftp""><div class='img'><img src=""webftp/images/folder.gif""></div><div class='txt'>" & F.Name & "</div></a></li>"& chr(10) & chr(10) & chr(9)
			Next
			For Each F In Root.files
				Extarr = Split(F.Name,".")
				Ext = LCase(Extarr(Ubound(Extarr)))
				Href = "#"
				If Instr(LCase(ArrowEXT),LCase(Ext)) > 0 Then Href="?urli=" & server.urlencode(urli) & "&urlr="&server.urlencode(urli) & "&file=" & server.urlencode(urli) & "/" & server.urlencode(F.name)
				if Instr("/png/jsp/asa/bat/rm/mp3/pdf/wma/rmvb/asp/html/htm/shtm/shtml/php/css/js/txt/gif/jpeg/jpg/bmp/swf/mdb/doc/xls/rar/zip/exe/xml/xsl/vbs/","/" & Ext & "/") > 0 Then Extimg = Ext Else Extimg = "file"
				Response.write "<li><a href=""" & href &"&ext="& Ext&"#webftp"" title=""NAME: " & F.Name & "; TYPE: " & F.Type & ";""><div class='img'><img src=""webftp/images/" & Extimg & ".gif""></div><div class='txt'>" & F.Name & "</div></a></li>" & chr(10) & chr(10) & chr(9)
			Next
			%></ul>
		</td>
	</tr>
   
	<%
	set Root = nothing
	%>
</table>

					
               </div>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
