<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_guestbook.asp"-->
<%
'=========================================================
' File Name：	admin_guestbook.asp
' Purpose：		留言管理管理
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-27 9:31:00
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "<a href=""admin_guestbook.asp"">管理留言</a>"

Call ChkLogin()	'检查登录
Call Init()		'初始化页面

'初始化页面
Sub Init()

	Select Case LCase(act)
		Case "recomment"
			SubStatus = "回复留言"
			Call Main("recomment")
		Case "setstate"
			Call SetState()
		Case "dorecomment"
			Call DoRecomment()
		Case "dodelete"
			Call DoDelete()
		Case Else
			SubStatus = "留言列表"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'更新状态
Sub SetState()
	Dim st: st = Request("state")
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id参数错误", "REFRESH")
	Select Case LCase(st)
		Case "pass"
			Call DB("UPDATE [GuestBook] SET State=1 WHERE ID IN ("& id &")" ,0)
			Call WebLog("审核留言["& id &"]成功！", "SESSION")
			Call MsgAndGo("审核留言["& id &"]成功", "REFRESH")
		Case "nopass"
			Call DB("UPDATE [GuestBook] SET State=0 WHERE ID IN ("& id &")" ,0)
			Call WebLog("取消审核留言["& id &"]成功！", "SESSION")
			Call MsgAndGo("取消审核留言["& id &"]成功！", "REFRESH")
		Case Else
			Call MsgBox("参数错误！", "BACK")
	End Select
End Sub

'回复留言
Function DoRecomment()
	Dim objA: Set objA = New ClassGuestBook
	objA.ID = id
	If objA.Comment Then
		Call WebLog("回复留言[id:"& id &"]成功！", "SESSION")
		Call MsgAndGo("回复留言[id:"& id &"]成功！", "admin_guestbook.asp")
	Else
		Call MsgBox("错误：" & objA.LastError, "REFRESH")
	End If
	Set objA = Nothing
End Function

'删除留言
Sub DoDelete()
	Dim objA: Set objA = New ClassGuestBook
	objA.ID = id
	If objA.Delete Then
		Call WebLog("删除留言[id:"& id &"]成功！", "SESSION")
		Call MsgAndGo("删除留言[id:"& id &"]成功！", "admin_guestbook.asp")
	Else
		Call MsgBox("错误：" & objA.LastError, "REFRESH")
	End If
	Set objA = Nothing
End Sub

'主函数
Sub Main(ByVal strType)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>后台管理 - 留言管理 - <%=SYS%></title>
<script type="text/javascript" src="inc/base.js"></script>
<link href="css/common.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
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

.txtArea{
	margin:0px 5px 0px 5px;
	border:solid 1px #CCC;
	color:#333;
}
.txtArea .title{
	font-size:14px;
	font-weight:bold;
	height:30px;
	line-height:30px;
	padding-left:10px;
	color:#366d99;
	border-bottom:#CCC 1px dashed;
}
.txtArea .content{
	margin:0px;
	padding:10px;
	line-height:25px;
	font-size:13px;
}
.txtArea .info{
	height:28px;
	line-height:28px;
	padding-left:10px;
	border-top:#CCC 1px dashed;
	background:#F7F7F7;
	color:#AAA;
}
.txtArea  a{ color:#000; text-decoration:none;}
.txtArea .title a{color:#366d99; text-decoration:none;}
.txtArea a:hover{ color:#F00;}
.reComment{
	margin:5px;
	border:dashed 1px #CCC;
	padding:5px;
	line-height:22px;
	background:#F5F5F5;
	color:#090;
}
.manage{color:#000; margin-left:10px; padding-left:10px; border-left:solid 1px #09F;}
.manage a{ color:#069;}
.green, .green a{{color:green;}
.red, .red a{ color:#F00;}
.blue, .blue a{ color:blue;}
.gray, .gray a{ color:gray;}
.batCtrl { padding:5px;}
-->
</style>
</head>

<body>
<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td colspan="2" id="topNav">
			<%Call TopNav("guestbook")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">管理首页</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== 留言管理 ==--</li>
                 <li class="on"><a href="?action=list">管理留言</a></li>
                </ul>
				<%Call SysInfo()%>
            </td>
            <td id="content" valign="top">
            	<div class="content">
                	<div class="status"> 您的位置：<a href="index.asp">管理首页</a> → <%=MainStatus%> → <%=SubStatus%> </div>
					<%
                        Select Case LCase(strType)
                            Case "recomment"
                                FuncForm(id)
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
'自定义页面列表， mode - 模式
Sub List()
%>
	<form name="form2" action="" method="post">
	<%
		Dim strSql, Rs
		strSql = "SELECT * FROM [GuestBook] ORDER BY ID DESC"
		Set Rs = New ClassPageList
		Rs.Result = 1
		Rs.Sql = strSql
		Rs.PageSize = 10
		Rs.AbsolutePage = page
		Rs.List()
		Dim i: i = 1
		For i = 1 To Rs.PageSize
			If Rs.Data.Eof Then Exit For
	%>
            <div class="txtArea" onMouseOver="this.style.border='#F00 1px solid';" onMouseOut="this.style.border='#CCC 1px solid';">
                <div class="title"><%=Rs.Data("ID")%>: <%=Rs.Data("Title")%></div> 
                <div class="content"> 
                	<%=Rs.Data("Content")%>
                </div>
                <!-- 留言回复 -->
			
                <div class="reComment">
                <%If Len(Rs.Data("Recomment"))>0 Then
					Echo("<b>"&Rs.Data("ReUser")& "</b>回复：" &Rs.Data("Recomment")) 
				  Else
				  	Echo("暂无回复")
				  End If
				%>
                </div>
                
                <!-- 留言信息 -->
                <div class="info">留言者：<%=Rs.Data("User")%> E-mail:<%=Rs.Data("Email")%> 主页：<a href="<%=Rs.Data("HomePage")%>" target="_blank">浏览</a>　IP:<%=Rs.Data("IP")%>　发表：<%=Rs.Data("CreateTime")%>
                 <span class="manage">
                 	<input type="checkbox" name="GroupID" value="<%=Rs.Data("ID")%>" />
                     状态: 
                     <%If Rs.Data("State") = 1 Then%>
                     	<a href="?id=<%=Rs.Data("ID")%>&action=setstate&state=nopass" title="取消审核">已审核</a>
                     <%Else%>
                     	<span class="red"><a href="?id=<%=Rs.Data("ID")%>&action=setstate&state=pass" title="通过审核">未审核</a></span>
                     <%End If%>
                     <a href="?id=<%=Rs.Data("ID")%>&action=recomment">[回复] </a>
                     <a href="?id=<%=Rs.Data("ID")%>&action=dodelete">[删除]</a>
                </span>
                </div>
            </div> 
	<%
			Rs.Data.MoveNext
		Next
	%>	
			<div class="batCtrl">
  				<input type="button" onClick="selectAll(this.form,'GroupID')" value="全选" /> 
                <input type="button" onClick="selectOthers(this.form,'GroupID')" value="反选" /> 
                &nbsp;&nbsp;
                批量操作：
                <select name="name" onChange="dobatch(this)" style="line-height:25px; padding:5px;">
                	<option value=""> 选择操作 </option>
                    <option value="pass"> 通过审核 </option>
                    <option value="nopass"> 取消审核 </option>
                    <option value="delete"> 彻底删除 </option>
                </select>
            </div>
    </form>
    
    <div class="page"><%=Rs.Page%></div>
    
    
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
function selectAll(form, name)
{
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if (name == '' || e.name == name) {
			e.checked = true;
		}
	}
}


// 表单反选
function selectOthers(form, name){	
	
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if (name == '' || e.name == name) {
			e.checked = !e.checked;
		}
	}
}

//获取ID
function getID(form){
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

//批量显示
function batchPass(form){
	var id = getID(form);
	if(!id){return;}
	if(confirm('确定把选中选项通过审核？')){	
		form.action  = '?action=setstate&state=pass&id=' + id;
		form.submit(); 
	}
}

//批量隐藏
function batchNoPass(form){
	var id = getID(form);
	if(!id){return;}
	if(confirm('确定把选中选项取消审核？')){	
		form.action  = '?action=setstate&state=nopass&id=' + id;
		form.submit(); 
	}
} 

//批量删除
function batchDelete(form){
	var id = getID(form);
	if(!id){return;}
	if(confirm('删除将不能恢复！\n\n是否真的删除？')){	
		form.action  = '?action=dodelete&id=' + id;
		form.submit(); 
	}
} 

//批处理操作
function dobatch(objSel){
	switch(objSel.options[objSel.selectedIndex].value){
		case 'pass':
			batchPass(objSel.form);
			break;
		case 'nopass':
			batchNoPass(objSel.form);
			break;
		case 'delete':
			batchDelete(objSel.form);
			break;
		default:
			return false;
	}
	objSel.selectedIndex = 0;
}
//-->
</script>
<%
	Rs.Data.Close: Set Rs = Nothing
End Sub%>

<%
'自定义页面表单
Sub FuncForm(ByVal id)
	Dim objA: Set objA = New ClassGuestBook
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("对不起，你编辑的留言不存在", "REFRESH")
	End If
%>
	<form action="?action=dorecomment" id="form1" name="form1" method="post">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <table class="form" style="border:1px #88C4FF solid;">
            <tr><th colspan="2">
				回复留言
            </th></tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
            	<td align="right" width="15%">标题：</td>
            	<td><%=objA.Title%></td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">留言：</td>
                <td ><%=objA.Content%></td>
            </tr>
            <tr>
            	<td align="right" width="15%">回复内容：</td>
                <td>
                    <textarea name="fReComment" cols="50" rows="5"><%=objA.ReComment%></textarea> (250字符以内)
                </td>
            </tr>
            <tr>
            	<td align="right" width="15%">您的名字：</td>
                <td>
                    <input type="text" name="fReUser" value="<%=Session("AdminName")%>"/>
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
    	<a href="javascript:history.go(-1)" onclick="return confirm('确定放弃回复留言?')"><< == 返回 << == </a>
    </div>
<script type="text/javascript">
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

</script>
<%
	Set objA = Nothing
End Sub
%>
