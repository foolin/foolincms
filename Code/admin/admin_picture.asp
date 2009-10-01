<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_picture.asp"-->
<%
'=========================================================
' File Name：	admin_picolumn.asp
' Purpose：		图片栏目管理
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-9 10:27:17
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "管理图片"

Call ChkLogin()	'检查登录
Call ChkPower("picture","all")	'检查权限
Call Init()		'初始化页面

'初始化页面
Sub Init()

	Select Case LCase(act)
		Case "create"
			SubStatus = "创建图片"
			Call Main("create")
		Case "modify"
			SubStatus = "修改图片"
			Call Main("modify")
		Case "setstate"
			Call SetState()
		Case "list"
			SubStatus = "图片列表"
			Call Main("list")
		Case "docreate"
			Call DoCreate()
		Case "domodify"
			Call DoModify()
		Case "dodelete"
			Call DoDelete()
		Case "dobatch"
			Call DoBatch()
		Case Else
			SubStatus = "图片列表"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'添加图片
Function DoCreate()
	Dim objC: Set objC = New ClassPicture
	If objC.BatCreate Then
		Call WebLog("添加图片[Title:"&objC.Title&"]成功！", "SESSION")
		Call MsgAndGo("添加图片[Title:"&objC.Title&"]成功！", "BACK")
	Else
		Call MsgBox("错误：" & objC.LastError, "BACK")
	End If
	Set objC = Nothing
End Function

'编辑图片信息
Sub DoModify()
	Dim objC: Set objC = New ClassPicture
	objC.ID = id
	If objC.SetValue And objC.Modify Then
		Call WebLog("修改图片[id:"& id &"]成功！", "SESSION")
		Call MsgAndGo("修改图片[id:"& id &"]成功！", "BACK")
	Else
		Call MsgBox("错误：" & objC.LastError, "BACK")
	End If
	Set objC = Nothing
End Sub

'删除图片
Sub DoDelete()
	Dim objC: Set objC = New ClassPicture
	objC.ID = id
	If objC.Delete Then
		Call WebLog("删除图片[id:"& id &"]成功！", "SESSION")
		Call MsgAndGo("删除图片[id:"& id &"]成功！", "REFRESH")
	Else
		Call MsgBox("错误：" & objC.LastError, "BACK")
	End If
	Set objC = Nothing
End Sub

Sub SetState()
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id参数错误", "BACK")
	Dim state: state = Request("state")
	Select Case LCase(state)
		Case "pass"
			Call DB("UPDATE Picture SET State = 1 WHERE ID = " & id, 0)
			Call WebLog("审核图片[id:"& id &"]成功!", "SESSION")
			Call MsgAndGo("审核图片[id:"& id &"]成功!", "REFRESH")
		Case "nopass"
			Call DB("UPDATE Picture SET State = 0 WHERE ID = " & id, 0)
			Call WebLog("取消审核图片[id:"& id &"]成功!", "SESSION")
			Call MsgAndGo("取消审核图片[id:"& id &"]成功!", "REFRESH")
		Case "delete"
			Call DB("UPDATE Picture SET State = -1 WHERE ID = " & id, 0)
			Call WebLog("删除图片[id:"& id &"]成功!", "SESSION")
			Call MsgAndGo("删除图片[id:"& id &"]成功!", "REFRESH")
		Case "nodelete"
			Call DB("UPDATE Picture SET State = 0 WHERE ID = " & id, 0)
			Call WebLog("还原图片[id:"& id &"]成功!", "SESSION")
			Call MsgAndGo("恭喜，还原图片[id:"& id &"]成功!", "REFRESH")
		Case Else
			Call MsgBox("对不起，您的错误操作错误！", "BACK")
	End Select
End Sub

'披处理操作
Sub DoBatch()
	Dim bat: bat = Request("batch")
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id参数错误", "BACK")
	Select Case LCase(bat)
		Case "pass"
			Call DB("UPDATE Picture SET State = 1 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量审核图片[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量审核图片[id:"& id &"]成功!", "REFRESH")
		Case "nopass"
			Call DB("UPDATE Picture SET State = 0 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量取消审核图片[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量取消审核图片[id:"& id &"]成功!", "REFRESH")
		Case "top"
			Call DB("UPDATE Picture SET IsTop = 1 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量置顶图片[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量置顶图片[id:"& id &"]成功!", "REFRESH")
		Case "notop"
			Call DB("UPDATE Picture SET IsTop = 0 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量取消置顶图片[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量取消置顶图片[id:"& id &"]成功!", "REFRESH")
		Case "trash"
			Call DB("UPDATE Picture SET State = -1 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量删除图片[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量删除图片[id:"& id &"]成功!", "REFRESH")
		Case "notrash"
			Call DB("UPDATE Picture SET State = 0 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量还原图片[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量还原图片[id:"& id &"]成功!", "REFRESH")
		Case "delete"
			Dim batRs
			Set batRs = DB("Select * From [Picture] Where [ID] IN (" & id & ")",1)
			If batRs.Eof Then batRs.Close : Set batRs = Nothing
			While Not batRs.Eof
				'删除文件
				If ExistFile("../"&batRs("SmallPicPath")) Then
					Call DeleteFile("../" & batRs("SmallPicPath"))
				End If
				If ExistFile("../"&batRs("PicPath")) Then
					Call DeleteFile("../" & batRs("PicPath"))
				End If
				DB "Delete From [Picture] Where [ID] = " & batRs("ID") ,0
				batRs.MoveNext
			Wend
			batRs.Close : Set batRs = Nothing
			Call WebLog("批量彻底删除图片[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量彻底删除图片[id:"& id &"]成功!", "REFRESH")
		Case Else
			Call MsgBox("操作错误", "BACK")
	End Select
End Sub

'主函数
Sub Main(ByVal picType)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>后台管理 - 图片管理 - <%=SYS%></title>
<link href="css/common.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="inc/base.js"></script>
<script language="javascript" type="text/javascript">
<!--
// 表单全选或者取消
function Checked(form, name, _this)
{
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if (name == '' || e.name == name) {
			if(_this.checked == false)
				e.checked = false;
			else
				e.checked = true;
		}
	}
}

// 表单全选
function SelectAll(form, name)
{
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if (name == '' || e.name == name) {
			e.checked = true;
		}
	}
}


// 表单反选
function SelectOthers(form, name){	
	
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if (name == '' || e.name == name) {
			e.checked = !e.checked;
		}
	}
}

//获取ID
function GetID(form){
	var name = "GroupID";
	var id = '';
	var intCount = 0;
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if ((name == '' || e.name == name) && e.checked) {
			id += id == '' ? e.value : ',' + e.value;
			intCount++;
		}
	}
	if(intCount==0){
		alert('您未选中任何选项！');
		return;
	}
	return id;
}

//批量处理审核
function BatchPass(form, isPass){
	var id = GetID(form);
	if(id){
		if(isPass){
			if (!confirm('是否把选中图片通过审核？')) return;
			form.action = '?action=dobatch&batch=pass&id=' + id;
			form.submit();  
		}
		else{
			if(!confirm('是否把选中图片取消审核？')) return;
			form.action  = '?action=dobatch&batch=nopass&id=' + id;
			form.submit();  
		}
	}
} 

//批量处理顶置
function BatchTop(form, isTop){
	var id = GetID(form);
	if(id){
		if(isTop){
			if(!confirm('是否把选中图片置顶？')) return;
			form.action  = '?action=dobatch&batch=top&id=' + id;
			form.submit();  
		}
		else{
			if(!confirm('是否把选中图片取消置顶？')) return;
			form.action  = '?action=dobatch&batch=notop&id=' + id;
			form.submit();  
		}
	}
} 

//批量移动到回收站
function BatchTrash(form, isTrash){
	var id = GetID(form);
	if(!id){return;}
	if (isTrash){
		if (confirm('是否把选中图片放到回收站？')){	
			form.action  = '?action=dobatch&batch=trash&id=' + id;
			form.submit();  
		}
	}
	else{
		if (confirm('是否把选中图片还原？')){	
			form.action  = '?action=dobatch&batch=notrash&id=' + id;
			form.submit();  
		}
	}
} 

//批量删除
function BatchDelete(form){
	var id = GetID(form);
	if(!id){return;}
	if(confirm('删除将不能恢复！\n\n是否真的删除？')){	
		form.action  = '?action=dobatch&batch=delete&id=' + id;
		form.submit(); 
	}
} 

//批处理操作
function Dobatch(objSel){
	switch(objSel.options[objSel.selectedIndex].value){
		case 'pass':
			BatchPass(objSel.form, true);
			break;
		case 'nopass':
			BatchPass(objSel.form, false);
			break;
		case 'top':
			BatchTop(objSel.form, true);
			break;
		case 'notop':
			BatchTop(objSel.form, false);
			break;
		case 'trash':
			BatchTrash(objSel.form, true);
			break;
		case 'notrash':
			BatchTrash(objSel.form, false);
			break;
		case 'delete':
			BatchDelete(objSel.form);
			break;
		default:
			return false;
	}
	objSel.selectedIndex = 0;
}
//-->
</script>
<style type="text/css">
<!--
.green {color:green;}
.red{ color:#F00;}
.blue{ color:blue;}
.btn{ padding:3px; background:#F7FBFF;}
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
.ke-content {
    font-family: Courier New;
    font-size: 12px;
    background-color: #ffffff;
}
#editor{ text-align:center; padding:2px;}
#editor table td{
	border:#6FB7FF 0px solid;
	padding:0px;
	line-height:normal;
}
.list{ font-size:13px;}
.img{ border:#CCC 1px solid; background:#FFF; padding:5px;}
-->
</style>
</head>

<body>
<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td colspan="2" id="topNav">
			<%Call TopNav("picture")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">管理首页</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== 图片管理 ==--</li>
                 <li <%If Request("action")="create" Then Echo("class=""on""")%>><a href="admin_picture.asp?action=create">上传图片</a></li>
                 <li <%If Request("action")<>"create" And Request("list")<>"trash" Then Echo("class=""on""")%>><a href="admin_picture.asp?action=list">管理图片</a></li>
                 <li <%If Request("list")="trash" Then Echo("class=""on""")%>><a href="admin_picture.asp?action=list&list=trash">图片回收站</a></li>
                 <li class="mTitle">--== 图片栏目 ==--</li>
                 <li><a href="admin_piccolumn.asp?action=create">添加栏目</a></li>
                 <li><a href="admin_piccolumn.asp">管理栏目</a></li>
                </ul>
				<%Call SysInfo()%>
            </td>
            <td id="content" valign="top">
            	<div class="content">
                	<div class="status"> 您的位置：<a href="index.asp">管理首页</a> → <%=MainStatus%> → <%=SubStatus%></div>
					<%
                        Select Case LCase(picType)
                            Case "create"
                                ColForm(0)
                            Case "modify"
                                ColForm(id)
							Case "list"
								List()
                            Case Else
                                List()
                        End Select
                    %>
                </div>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
<%End Sub%>

<%
'图片列表， mode - 模式
Sub List()
%>
	<form name="form2" action="" method="post">
	<table class="list">
    	<tr>
        	<td colspan="4">
                    <div style="font-size:14px; line-height:25px; padding-left:5px;">
                        <a href="?list=all">全部图片</a> | 
                        <a href="?list=pass">已经审核</a> | 
                        <a href="?list=nopass">未审核</a> |
                        <a href="?list=trash">回收站</a> |
                        <a href="?action=create">添加图片</a>
                    </div>
            </td>
        </tr>
        <tr >
	<%
		Dim strSql, Rs
		Dim mode: mode = LCase(Request("list"))
		Select Case mode
			Case "trash"
				strSql = "SELECT * FROM [Picture] WHERE State = -1 ORDER BY IsTop DESC,ID DESC"
			Case "nopass"
				strSql = "SELECT * FROM [Picture] WHERE State = 0 ORDER BY IsTop DESC,ID DESC"
			Case "pass"
				strSql = "SELECT * FROM [Picture] WHERE State = 1 ORDER BY IsTop DESC,ID DESC"
			Case "all"
				strSql = "SELECT * FROM [Picture] WHERE State > -1 ORDER BY IsTop DESC,ID DESC"
			Case Else
				strSql = "SELECT * FROM [Picture] WHERE State > -1 ORDER BY IsTop DESC,ID DESC"
		End Select
		Set Rs = New ClassPageList
		Rs.Result = 1
		Rs.Sql = strSql
		Rs.PageSize = 9
		Rs.AbsolutePage = page
		Rs.List()
		Dim i: i = 1
		For i = 1 To Rs.PageSize
			If Rs.Data.Eof Then Exit For
			Call LoopEcho(chr(9), 4)
			Echo("<td  align=""center""  style=""padding:5px""  onmouseover=""this.style.background='#FFFFFF';"" onmouseout=""this.style.background='#F0F8FF'""><a href=""?action=modify&id=" &Rs.Data("ID")& """><img src=""")
			If Rs.Data("SmallPicPath")<>"" Then
				Echo("../"&Rs.Data("SmallPicPath"))
			Else
				Echo("../"&Rs.Data("PicPath"))
			End If
			Echo """ width=""150"" height=""120"" class=""img""  /></a><br /><a href=""?action=modify&id=" &Rs.Data("ID")& """>"&Rs.Data("Title") & "</a>"
			If Rs.Data("IsTop")=1 Then Echo("<font color='red'>[顶]</font>")
			Echo "<hr style=""border:dotted 1px #B5DAFF""/> "
			Echo "<input type=""checkbox"" name=""GroupID"" value=" & Rs.Data("ID") & " />"
			If Rs.Data("State")=1 Then
				Echo " <a href=""?action=setstate&id="&Rs.Data("ID")&"&state=nopass""  title=""点击取消通过审核"">已审</a> |"
			ElseIf Rs.Data("State")=0 Then
				Echo " <a href=""?action=setstate&id="&Rs.Data("ID")&"&state=pass"" title=""点击通过审核""><font color=""red"">未审</font></a>  |"
			End If
			If Rs.Data("State")=-1 Then
				Echo " <a href=""?action=setstate&id=" &Rs.Data("ID")& "&state=nodelete"" onClick=""return confirm('确定还原照片？')"" title=""还原照片"">还原</a> | "
			Else
				Echo " <a href=""?action=modify&id=" &Rs.Data("ID")& """>编辑</a> | "
			End If
			If Rs.Data("State")=-1 Then
				Echo "<a href=""?action=dodelete&id=" &Rs.Data("ID")& """ onClick=""return confirm('确定把照片删除？删除将不能再恢复！')"" title=""删除"">删除</a></td>"
			Else
				Echo "<a href=""?action=setstate&id=" &Rs.Data("ID")& "&state=delete"" onClick=""return confirm('确定把照片放进回收站？')"" title=""删除"">删除</a></td>"& chr(10)
			End If

			If i Mod 3 = 0 Then
				Echo "</tr><tr>" & chr(10) & chr(10)
			End If	
			Rs.Data.MoveNext
		Next
	%>
    	 </tr>
       <tr>
        	<td colspan="4" style="padding:5px;">
  				<input type="button" onClick="SelectAll(this.form,'GroupID')" value="全选" /> 
                <input type="button" onClick="SelectOthers(this.form,'GroupID')" value="反选" /> 
                &nbsp;&nbsp;
                批量操作：
                <select name="name" onChange="Dobatch(this)" style="line-height:25px; padding:5px;">
                	<option value=""> 选择操作 </option>
                    <%If Request("list") = "trash" Then%>
                    <option value="notrash"> 还原 </option>
                    <%Else%>
                    <option value="pass"> 通过审核 </option>
                    <option value="nopass"> 取消审核 </option>
                    <option value="top"> 设置置顶 </option>
                    <option value="notop"> 取消顶置 </option>
                    <option value="trash"> 删除 </option>
                    <%End If%>
                    <option value="delete"> 彻底删除 </option>
                </select>
            </td>
        </tr>
    </table>
    </form>
    <div class="page"><%=Rs.Page%></div>
<%
	Rs.Data.Close: Set Rs = Nothing
End Sub
%>

<%
'图片表单
Sub ColForm(ByVal id)
	Dim objA: Set objA = New ClassPicture
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("对不起，你编辑的栏目不存在", "BACK")
	End If
%>

	<form action="?action=do<%If id > 0 Then Echo("modify") Else Echo("create")%>" id="form1" name="form1" method="post" onsubmit="return chkSubmit();">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <input type="hidden" name="fHits" value="<%=objA.Hits%>"/>
        <table class="form" style="border:1px #88C4FF solid;">

            <%If objA.PicPath<>"" Then %>
            <tr>
            	<td colspan="2"><div style="text-align:center; padding:5px;"><img class="img" src="../<%=objA.PicPath%>" width="500"  /></div></td>
            </tr>
            <%End If%>
                        <tr><th colspan="2">
				<%If id > 0 Then Echo("编辑") Else Echo("上传")%>图片
            </th></tr>
            <tr>
            	<td align="right" width="15%">标题：</td>
            	<td><input type="text" name="fTitle" value="<%=objA.Title%>" style="width:450px;"/> <span class="red">* 必填</span></td>
            </tr>
            <tr>
                <td align="right">栏目：</td>
                <td>
                	<select name="fColID">
                        	<option value="0"> => 请选择栏目 <= </option>
                    	<%If objA.ColID> 0 Then%>
                    		<option value="<%=objA.ColID%>" selected="selected"> => <%=GetColName(objA.ColID, "picture")%> <= </option>
                        <%End If%>
                    	<%Call MainColumn()%>
                    </select><span class="red">* 必选</span>（不选择，则作为父栏目）
                </td>
            </tr>
            <tr>
                <td align="right">作者：</td>
                <td><input type="text" name="fAuthor" value="<%=objA.Author%>" /></td>
            </tr>
            <tr>
                <td align="right">来源：</td>
                <td><input type="text" name="fSource" value="<%=objA.Source%>" /></td>
            </tr>
            <tr>
                <td align="right">选项：</td>
                <td>
                	置顶<input type="checkbox" name="fIsTop" value="1"  <%If objA.IsTop = 1 Then Echo("checked=""checked""")%> />  
                	通过审核<input type="checkbox" name="fState" value="1" <%If objA.State = 1 Then Echo("checked=""checked""")%> />
                </td>
            </tr>
            <tr>
            	<td align="right">图片介绍：</td>
                <td><textarea name="fIntro" style="width:99%;height:100px;"><%=objA.Intro%></textarea></td>
            </tr>
            <%If objA.PicPath="" Then %>
            <tr>
            	<td align="right">上传图片：</td>
                <td>
                    <div id="uploadFocusPic">
                    <iframe frameborder="0" src="inc/uploader/upload_picture.asp" width="80%" height="30"></iframe>
                     &nbsp;<span class="red">* 必填</span>
                    </div>
                </td>
            </tr>
            <%End If%>
            <tr>
            	<td align="right">缩略图路径：</td>
                <td><input type="text" name="fSmallPicPath" value="<%=objA.SmallPicPath%>" <%If objA.ID>0 Then Echo("readonly=""readonly""")%> style="width:450px;"/></td>
            </tr>
            <tr>
            	<td align="right">图片路径：</td>
                <td><input type="text" name="fPicPath" <%If objA.ID>0 Then Echo("readonly=""readonly""")%>  value="<%=objA.PicPath%>" style="width:450px;" /> <span class="red" id="PicNum">0</span>张</td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <input type="submit" class="btn" value="提交" />
                    <input type="reset" class="btn" value="重置" />
                </td>
            </tr>
        </table>
    </div>
    <div class="page">
    	<a href="javascript:history.go(-1)" onclick="return confirm('确定放弃修改?')"><< == 返回 << == </a>
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

function $F(el){return document.forms["form1"].elements[el];}
//判断几张图片
if ($F("fPicPath").value != ""){
	$("PicNum").innerHTML = '1';
}
//检查提交
function chkSubmit(){
	if (parseInt($F("fTitle").value.length) < 2 || parseInt($F("fTitle").value.length) > 20)
	{
		alert("标题长度控制在 2 至 20 字符之间");
		$F("fTitle").focus();
		return false;
	}
	if (parseInt($F("fColID").value) <= 0){
		alert("请选择栏目");
		return false;
	}
	if ($F("fPicPath").value == ""){
		alert("请上传图片");
		$F("fPicPath").focus();
		return false;
	}
	return true;
}
//-->
</script>
<%
	Set objA = Nothing
End Sub

'第一级栏目分类
Function MainColumn()
	Dim Rs
	Set Rs = DB( "SELECT * FROM PicColumn WHERE ParentID = 0", 1)
	If Not Rs.Eof Then
		Do While Not Rs.Eof
			If Rs("ID") <> Cint(Request("id")) Then
				Echo("<option value=""" & Rs("ID") & """>" & Rs("Name") & "</option>" & Chr(10) & Chr(9) & Chr(9))
			End If
			Call SubColumn(Rs("ID"),"|-") '循环子级分类
		Rs.MoveNext
		If Rs.Eof Then Exit Do '防上造成死循环
		Loop
	End If
	Rs.Close: Set Rs = Nothing
End Function
'子栏目分类
Function SubColumn(FID,StrDis)
	Dim Rs1
	Set Rs1 = DB("SELECT * FROM PicColumn WHERE ParentID = " & FID, 1)
	If Not Rs1.Eof Then
		Do While Not Rs1.Eof
			If Rs1("ID") <> Cint(Request("id")) Then
				Echo("<option value=""" & Rs1("ID") & """>" & StrDis & Rs1("Name") & "</option>" & Chr(10) & Chr(9))
			End If
			Call SubColumn(Trim(Rs1("ID")),"| " & Strdis) '递归子级分类
		Rs1.Movenext:Loop
		If Rs1.Eof Then
			Rs1.Close: Set Rs1 = Nothing
			Exit Function
		End If
	End If
	Rs1.Close: Set Rs1 = Nothing
End Function

%>
