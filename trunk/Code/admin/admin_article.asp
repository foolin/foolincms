<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_article.asp"-->
<%
'=========================================================
' File Name：	admin_article.asp
' Purpose：		文章管理
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-3 16:40:50
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "管理文章"

Call ChkLogin()	'检查登录
Call ChkPower("article","all")	'检查是否拥有管理权限
Call Init()		'初始化页面

'初始化页面
Sub Init()

	Select Case LCase(act)
		Case "create"
			SubStatus = "创建文章"
			If IsNullColumn = True Then
				Call MsgBox("尚未有任何栏目，请先添加栏目!","admin_artcolumn.asp?action=create")
			End If
			Call Main("create")
		Case "modify"
			SubStatus = "修改文章"
			Call Main("modify")
		Case "setstate"
			Call SetState()
		Case "list"
			SubStatus = "文章列表"
			Call Main("list")
		Case "docreate"
			Call DoCreate()
		Case "domodify"
			Call DoModify()
		Case "dodelete"
			Call DoDelete()
		case "dobatch"
			Call DoBatch()
		Case Else
			SubStatus = "文章列表"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'更新状态
Sub SetState()
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id参数错误", "BACK")
	Dim state: state = Request("state")
	Select Case LCase(state)
		Case "pass"
			Call DB("UPDATE Article SET State = 1 WHERE ID = " & id, 0)
			Call WebLog("审核文章[id:"& id &"]成功!", "SESSION")
			Call MsgAndGo("审核文章[id:"& id &"]成功!", "REFRESH")
		Case "nopass"
			Call DB("UPDATE Article SET State = 0 WHERE ID = " & id, 0)
			Call WebLog("取消审核文章[id:"& id &"]成功!", "SESSION")
			Call MsgAndGo("取消审核文章[id:"& id &"]成功!", "REFRESH")
		Case "delete"
			Call DB("UPDATE Article SET State = -1 WHERE ID = " & id, 0)
			Call WebLog("删除文章[id:"& id &"]成功!", "SESSION")
			Call MsgAndGo("删除文章[id:"& id &"]成功!", "REFRESH")
		Case "nodelete"
			Call DB("UPDATE Article SET State = 0 WHERE ID = " & id, 0)
			Call WebLog("还原文章[id:"& id &"]成功!", "SESSION")
			Call MsgAndGo("恭喜，还原文章[id:"& id &"]成功!", "REFRESH")
		Case Else
			Call MsgBox("对不起，您的错误操作错误！", "BACK")
	End Select
End Sub

'创建文章
Function DoCreate()
	Dim objA: Set objA = New ClassArticle
	If objA.SetValue = False Then
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	If objA.Create Then
		Call WebLog("发表文章[id:title:"&objA.Title&"]成功！", "SESSION")
		Call MsgAndGo("发表文章[id:title:"&objA.Title&"]成功！", "BACK")
	Else
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Function

'删除文章
Sub DoModify()
	Dim objA: Set objA = New ClassArticle
	objA.ID = id
	If objA.SetValue = False Then
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	If objA.Modify Then
		Call WebLog("修改文章[id:"& id &"]成功！", "SESSION")
		Call MsgAndGo("修改文章[id:"& id &"]成功！", "admin_article.asp")
	Else
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Sub

'删除文章
Sub DoDelete()
	Dim objA: Set objA = New ClassArticle
	objA.ID = id
	If objA.Delete Then
		Call WebLog("删除文章[id:"& id &"]成功！", "SESSION")
		Call MsgAndGo("删除文章[id:"& id &"]成功！", "REFRESH")
	Else
		Call MsgBox("错误：" & objC.LastError, "BACK")
	End If
	Set objA = Nothing
End Sub

'披处理操作
Sub DoBatch()
	Dim bat: bat = Request("batch")
	Dim colId: colId = Request("colid")
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id参数错误", "BACK")
	Select Case LCase(bat)
		Case "move"
			If Len(colId) = 0 Or Not IsNumeric(colId) Then Call MsgBox("栏目id参数错误", "BACK")
			Call DB("UPDATE Article SET ColID = "& colId &" WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量移动文章[id:"& id &"]To栏目["&colId&"]成功！", "SESSION")
			Call MsgAndGo("批量移动文章[id:"& id &"]To栏目["&colId&"]成功!", "REFRESH")
		Case "pass"
			Call DB("UPDATE Article SET State = 1 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量审核文章[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量审核文章[id:"& id &"]成功!", "REFRESH")
		Case "nopass"
			Call DB("UPDATE Article SET State = 0 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量取消审核文章[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量取消审核文章[id:"& id &"]成功!", "REFRESH")
		Case "top"
			Call DB("UPDATE Article SET IsTop = 1 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量置顶文章[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量置顶文章[id:"& id &"]成功!", "REFRESH")
		Case "notop"
			Call DB("UPDATE Article SET IsTop = 0 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量取消置顶文章[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量取消置顶文章[id:"& id &"]成功!", "REFRESH")
		Case "trash"
			Call DB("UPDATE Article SET State = -1 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量删除文章[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量删除文章[id:"& id &"]成功!", "REFRESH")
		Case "notrash"
			Call DB("UPDATE Article SET State = 0 WHERE ID IN (" & id & ")", 0)
			Call WebLog("批量还原文章[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量还原文章[id:"& id &"]成功!", "REFRESH")
		Case "delete"
			Call DB("Delete From [Article] Where [ID] In (" & id & ")", 0)
			Call WebLog("批量彻底删除文章[id:"& id &"]成功！", "SESSION")
			Call MsgAndGo("批量彻底删除文章[id:"& id &"]成功!", "REFRESH")
		Case Else
			Call MsgBox("操作错误", "BACK")
	End Select
End Sub

'检查栏目是否为空
Function IsNullColumn()
	Dim cRs,cFlag
	Set cRs = DB("SELECT * FROM ArtColumn", 1)
	If cRs.Eof Then
		cFlag = True
	Else
		cFlag = False
	End If
	Set cRs = Nothing
	IsNullColumn = cFlag
End Function


'主函数
Sub Main(ByVal artType)
Dim SubStatus2
Select Case LCase(Request("list"))
	Case "trash"
		SubStatus2 = " → 回收站"
	Case "nopass"
		SubStatus2 = " → 未审核"
	Case "pass"
		SubStatus2 = " → 已经审核"
	Case "all"
		SubStatus2 = ""
	Case Else
		SubStatus2 = ""
End Select
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>后台管理 - 文章管理 - Powered by eekku.com</title>
<link href="images/common.css" rel="stylesheet" type="text/css" />
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
			if (!confirm('是否把选中文章通过审核？')) return;
			form.action = '?action=dobatch&batch=pass&id=' + id;
			form.submit();  
		}
		else{
			if(!confirm('是否把选中文章取消审核？')) return;
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
			if(!confirm('是否把选中文章置顶？')) return;
			form.action  = '?action=dobatch&batch=top&id=' + id;
			form.submit();  
		}
		else{
			if(!confirm('是否把选中文章取消置顶？')) return;
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
		if (confirm('是否把选中文章放到回收站？')){	
			form.action  = '?action=dobatch&batch=trash&id=' + id;
			form.submit();  
		}
	}
	else{
		if (confirm('是否把选中文章还原？')){	
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

function BatchMove(form){
	var id = GetID(form);
	if(!id){return;}
	var colid = $("toColId").value;
	if ( parseInt(colid) == 0){
		alert('请选择栏目');
		return;
	}
	if (confirm('确定批量移动到该栏目？')){	
		form.action  = '?action=dobatch&batch=move&id=' + id + '&colid=' + colid;
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
		case 'move':
			$("batMove").style.display = "block";
			break;
		default:
			return false;
	}
	objSel.selectedIndex = 0;
}


//搜索文章
function soArticle(){
	var jumpUrl;
	jumpUrl = 'admin_article.asp?colid=' + $('sColId').value;
	if ($('sKeyword').value != "" && $('sKeyword').value !="请输入关键词"){
		jumpUrl =  jumpUrl + '&keyword=' + $('sKeyword').value;
	}
	this.location = jumpUrl;
	return false;
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

.openWin{
	position:fixed;
	left:30%;
	top:20%;
	width:350px;
	height:200px;
	border:#E3E3E3 5px solid;
	background:#FFF;
	overflow:auto;
}
.openWin .title{
	text-align:center;
	font-size:14px;
	font-weight:bold;
	line-height:35px;
	color:#666;
	border-bottom:#E3E3E3 2px solid;
	background:#F3F3F3;
}
.openWin .content{
	padding:5px;
	line-height:22px;
}
.openWin .close{
	text-align:center;
	padding:10px;
}
#batMove{ display:none;}
-->
</style>
</head>

<body>
<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td colspan="2" id="topNav">
			<%Call TopNav("article")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">管理首页</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== 文章管理 ==--</li>
                 <li <%If Request("action") = "create" Then Response.Write("class=""on""")%>><a href="admin_article.asp?action=create">添加文章</a></li>
                 <li <%If Request("action") <> "create" And Request("list") <> "trash"  Then Response.Write("class=""on""")%>><a href="admin_article.asp?action=list">管理文章</a></li>
                 <li <%If Request("list") = "trash" Then Response.Write("class=""on""")%>><a href="admin_article.asp?action=list&list=trash">文章回收站</a></li>
                 <li class="mTitle">--== 文章栏目 ==--</li>
                 <li><a href="admin_artcolumn.asp?action=create">添加栏目</a></li>
                 <li><a href="admin_artcolumn.asp">管理栏目</a></li>
                </ul>
				<%Call SysInfo()%>
            </td>
            <td id="content" valign="top">
            	<div class="content">
                	<div class="status"> 您的位置：<a href="index.asp">管理首页</a> → <%=MainStatus%> → <%=SubStatus%> <%=SubStatus2%> </div>
                    <div style="font-size:14px; line-height:25px; padding-left:5px;">
                        <a href="?list=all">全部文章</a> | 
                        <a href="?list=pass">已经审核</a> | 
                        <a href="?list=nopass">未审核</a> |
                        <a href="?list=trash">回收站</a> |
                        <a href="?action=create">添加文章</a>
                        

                    </div>
					<%
                        Select Case LCase(artType)
                            Case "create"
                                ArtForm(0)
                            Case "modify"
                                ArtForm(id)
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
'文章列表， mode - 模式
Sub List()
Dim mode: mode = LCase(Request("list"))
%>
	<form name="form2" action="" onsubmit="return soArticle();" method="post">
	<table class="list">
    	<tr>
        	<th><input type="checkbox" name="GroupID" value="" onClick="Checked(this.form,'GroupID',this)"/></th>
        	<th>ID</th>
            <th>标题</th>
            <th>栏目</th>
            <th>作者</th>
            <th>时间</th>
            <th>状态</th>
            <th>操作</th>
            <th>删除</th>
        </tr>
	<%
		Dim strSql, Rs
		Dim colId, sqlColId, strKeyword, sqlKeyword
		'栏目ID
		colId = Request("colid")
		If Len(Request("colid")) = 0 Then colId = 0
		If colId > 0 Then sqlColId = " And ColID IN ("& GetColIds(colId,"ARTICLE") &") "
		'搜索字符串
		strKeyword = Trim(Request("keyword"))
		If Len(strKeyword) > 0 Then sqlKeyword = " And Title LIKE '%"& strKeyword &"%' "
		'文章列表类型
		Select Case mode
			Case "trash"
				strSql = "SELECT * FROM [Article] WHERE State=-1 "& sqlColId & sqlKeyword &" ORDER BY IsTop DESC,ID DESC"
			Case "nopass"
				strSql = "SELECT * FROM [Article] WHERE State=0 "& sqlColId & sqlKeyword &" ORDER BY IsTop DESC,ID DESC"
			Case "pass"
				strSql = "SELECT * FROM [Article] WHERE State=1 "& sqlColId & sqlKeyword &" ORDER BY IsTop DESC,ID DESC"
			Case "all"
				strSql = "SELECT * FROM [Article] WHERE State>-1 "& sqlColId & sqlKeyword &" ORDER BY IsTop DESC,ID DESC"
			Case Else
				strSql = "SELECT * FROM [Article] WHERE State>-1 "& sqlColId & sqlKeyword &" ORDER BY IsTop DESC,ID DESC"
		End Select
		'Set Rs = DB(strSql, 1)
		Set Rs = New ClassPageList
		Rs.Result = 1
		Rs.Sql = strSql
		Rs.PageSize = 15
		Rs.AbsolutePage = page
		Rs.List()
		Dim i: i = 1
		For i = 1 To Rs.PageSize
			If Rs.Data.Eof Then Exit For
	%>
        <tr onMouseOver="this.style.background='#C8E3E2';" onMouseOut="this.style.background='#F0F8FF'">
        	<td><input type="checkbox" name="GroupID" value="<%=Rs.Data("ID")%>" /></td>
        	<td><%=Rs.Data("ID")%></td>
			<td>
            	<a href="admin_article.asp?action=modify&id=<%=Rs.Data("ID")%>">
                <%
					If Len(Request("keyword"))>0 Then
						Echo(Replace(Rs.Data("Title"),Request("keyword"),"<font color='red'>" & Request("keyword") & "</font>"))
					Else
						Echo(Rs.Data("Title"))
					End If
				%>
                </a>
				<%If Rs.Data("IsTop") = 1 Then Echo(" <font color=""red"">[顶]</font>")%>
                <%If Rs.Data("IsFocusPic") =1 And Rs.Data("FocusPic") <> "" Then Echo(" <font color=""red"">[图]</font>")%>
            </td>
            <td><a href="?colid=<%=Rs.Data("ColID")%>"><%=GetColName(Rs.Data("ColID"), "article")%></a></td>
            <td><%=Rs.Data("Author")%></td>
            <td><%=FDate(Rs.Data("CreateTime"), 2)%></td>
            <td>
            	<%If Rs.Data("State") = 1 Then%>
            		<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=nopass" title="点击取消审核" class="green">已审核</a>
                <%ElseIf Rs.Data("State") = 0 Then%>
                	<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=pass" title="点击通过审核" class="red">未审核</a>
                <%Else%>
					<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=nodelete" title="点击取消删除"  onclick="return confirm('确定还原数据？')" class="blue">已删除</a>
                <%End If%>
            </td>
            <td>
            	<%If Rs.Data("State") > -1 Then%>
            		<a href="?action=modify&id=<%=Rs.Data("ID")%>">编辑</a>
                <%Else%>
					            		<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=nodelete" onclick="return confirm('恢复该文章后为[未审核]状态，您确定还原数据？')">还原</a>
                <%End If%>
            </td>
            <td>
            	
            	<%If Rs.Data("State") > -1 Then%>
            		<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=delete" onclick="return confirm('确定把该文章放到回收站？')">删除</a>
                <%Else%>
					<a href="?action=dodelete&id=<%=Rs.Data("ID")%>" onclick="return confirm('删除将不用恢复数据！\n\n确定永久性删除该数据？')">删除</a>
                <%End If%>
            </td>
        </tr>
	<%
			Rs.Data.MoveNext
		Next
	%>
        <tr>
        	<td colspan="9" style="padding:5px;">
  				<input type="button" onClick="SelectAll(this.form,'GroupID')" value="全选" /> 
                <input type="button" onClick="SelectOthers(this.form,'GroupID')" value="反选" /> 
                &nbsp;&nbsp;
                批量操作：
                <select name="name" onChange="Dobatch(this)" style="line-height:25px; padding:5px;">
                	<option value=""> 选择操作 </option>
                    <option value="move"> 批量移动 </option>
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
                
                <div class="openWin" id="batMove">
                        <div class="title">请选择操作</div>
                        <div class="content">
                       
                            请选择的栏目：
                            <select name="toColId" id="toColId">
                                  <option value="0"> 请选择栏目 </option>
                                    <%Call MainColumn()%>
                            </select>
                            <input type="button" value="移动" onclick="BatchMove(this.form);" />
                            <br /> <br />
                        </div>
                        <div class="close"><a href="#" onclick="$('batMove').style.display='none';">[×] 关闭窗口</a></div>
                </div>
                
                 &nbsp; 搜索：<select name="sColId" id="sColId">
                 <option value="0"> 请选择栏目 </option>
                 <option value="0"> 全部栏目 </option>
                    <%Call MainColumn()%>
                 </select>
                 <input type="text" name="sKeyword" id="sKeyword" value="<%If Len(Request("keyword"))>0 Then Echo(Request("keyword")) Else Echo("请输入关键词")%>" onclick="if(this.value=='请输入关键词')this.value='';" />
                 <input type="button" value="搜索" onclick="soArticle();" />
                 
            </td>
        </tr>
    </table>
    </form>
    <div class="page"><%=Rs.Page%></div>
        
<%
	Rs.Data.Close: Set Rs = Nothing
End Sub%>

<%
'文章表单
Sub ArtForm(ByVal id)
	Dim objA: Set objA = New ClassArticle
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("对不起，你编辑的文章不存在", "BACK")
	Else
		objA.Author = Session("AdminName")
	End If
%>
	<form action="?action=do<%If id > 0 Then Echo("modify") Else Echo("create")%>" id="form1" name="form1" method="post">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <input type="hidden" name="Hits" value="<%=objA.Hits%>"/>
        <table class="form" style="border:1px #88C4FF solid;">
            <tr><th colspan="2">
				<%If id > 0 Then Echo("编辑") Else Echo("添加")%>文章
            </th></tr>
            <tr>
            	<td align="right" width="15%">标题：</td>
            	<td><input type="text" name="Title" value="<%=objA.Title%>" style="width:450px;"/> <span class="red">* 必填</span></td>
            </tr>
            <tr>
                <td align="right">栏目：</td>
                <td>
                	<select name="ColID">
                    	<%If id > 0 Then%>
                    		<option value="<%=objA.ColID%>"> => <%=GetColName(objA.ColID, "article")%> <= </option>
                        <%Else%>
                        	<option value="0"> => 请选择栏目 <= </option>
                        <%End If%>
                    	<%Call MainColumn()%>
                    </select>
                    <span class="red">* 必选</span>
                </td>
            </tr>
            <tr>
                <td align="right">作者：</td>
                <td><input type="text" name="Author" value="<%=objA.Author%>" /></td>
            </tr>
            <tr>
                <td align="right">来源：</td>
                <td><input type="text" name="Source" value="<%=objA.Source%>" /></td>
            </tr>
            <tr>
                <td align="right">焦点图片URL：</td>
                <td>
                	<input type="text" name="FocusPic" id="FocusPic" value="<%=objA.FocusPic%>" style="width:450px;" /> <a href="javascript:uploadFocusPic();">上传图片</a>
                    <div id="uploadFocusPic" style="display:none;">
                    <iframe frameborder="0" src="inc/upload_focuspic.asp" width="80%" height="30"></iframe>
                    </div>
                </td>
            </tr>
            <tr>
                <td align="right">关键词：</td>
                <td><input type="text" value="<%=objA.Keywords%>" name="Keywords"  style="width:450px;" /></td>
            </tr>
            <tr>
                <td align="right">选项：</td>
                <td>
                	置顶<input type="checkbox" name="IsTop" value="1"  <%If objA.IsTop = 1 Then Echo("checked=""checked""")%> />  
                	通过审核<input type="checkbox" name="State" value="1" <%If objA.State = 1 Then Echo("checked=""checked""")%> />
                    焦点图片<input type="checkbox" name="IsFocusPic" value="1" <%If objA.IsFocusPic = 1 Then Echo("checked=""checked""")%> id="IsFocusPic" onclick="chkFocusPic()"/>
                </td>
            </tr>
            <tr>
                <td align="right">跳转地址：</td>
                <td><input type="text" name="JumpUrl" id="JumpUrl" value="<%=objA.JumpUrl%>"  style="width:450px;" /></td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">内容：</td>
                <td ><span class="red">( <span class="green">内容</span> 和 <span class="green">跳转地址</span> 二者只能选其一  )</span></td>
            </tr>
            <tr>
                <td colspan="2">
                	<div id="editor">
                    <textarea id="content1" name="Content" style="width:100%;height:550px;visibility:hidden;"><%=objA.Content%></textarea>
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <input type="submit" class="btn" value="提交" />
                    <input type="reset" class="btn" value="重置" />
                </td>
            </tr>
        </table>
    </form>
    </div>
    <div class="page">
    	<a href="javascript:history.go(-1)" onclick="return confirm('确定放弃编辑文章?')"><< == 返回 << == </a>
    </div>
<script type="text/javascript" charset="utf-8" src="inc/editor/kindeditor.js"></script>
<script type="text/javascript">
//初始化编辑器
KE.show({
	id : 'content1',
	cssPath : 'inc/editor/editor.css',
	skinType: 'tinymce',
	items : [
		'source', 'preview',  'print', 'undo', 'redo', 'cut', 'copy', 'paste',
		'plainpaste', 'wordpaste', 'justifyleft', 'justifycenter', 'justifyright',
		'justifyfull', 'insertorderedlist', 'insertunorderedlist', 'indent', 'outdent', 'subscript',
		'superscript', 'date', 'time', 'specialchar', 'emoticons', 'link', 'unlink', '-',
		'title', 'fontname', 'fontsize', 'textcolor', 'bgcolor', 'bold',
		'italic', 'underline', 'strikethrough', 'removeformat', 'selectall', 'image',
		'flash', 'media', 'layer', 'table', 'hr', 'about'
	]
});
</script>
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
//chkJumpUrl(); 	//执行
function chkJumpUrl(){
	//alert(0)
	if ($("JumpUrl").value != ""){
		$("editor").style.display = "none";
	}
	else{
		$("editor").style.display = "block";
		//alert(2)
	}
}
function uploadFocusPic(){
	if ($("uploadFocusPic").style.display == "none"){
		$("uploadFocusPic").style.display = "block";
	}
	else{
		$("uploadFocusPic").style.display = "none";
	}
}
function chkFocusPic(){
	if ($("IsFocusPic").checked == true && $("FocusPic").value == ""){
		alert("您尚未填写焦点图片URL，请先上传图片！");
		$("IsFocusPic").checked = false;
		$("FocusPic").focus();
	}
}
//-->
</script>
<%
	Set objA = Nothing
End Sub

'第一级栏目分类
Function MainColumn()
	Dim Rs
	Set Rs = DB( "SELECT * FROM ArtColumn WHERE ParentID = 0 ORDER BY Sort DESC,ID", 1)
	If Not Rs.Eof Then
		Do While Not Rs.Eof
			Echo("<option value=""" & Rs("ID") & """>" & Rs("Name") & "</option>" & Chr(10) & Chr(9) & Chr(9))
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
	Set Rs1 = DB("SELECT * FROM ArtColumn WHERE ParentID = " & FID & " ORDER BY Sort DESC,ID", 1)
	If Not Rs1.Eof Then
		Do While Not Rs1.Eof
			Echo("<option value=""" & Rs1("ID") & """>" & StrDis & Rs1("Name") & "</option>" & Chr(10) & Chr(9))
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
