<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_admin.asp"-->
<%
'=========================================================
' File Name：	admin_mytag.asp
' Purpose：		团队管理
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-3 16:40:50
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "<a href='?'>管理团队</a>"

Call ChkLogin()	'检查登录
Call ChkPower("admin","all") '检查权限
Call Init()		'初始化页面

'初始化页面
Sub Init()

	Select Case LCase(act)
		Case "create"
			SubStatus = "创建管理员"
			Call Main("create")
		Case "modify"
			SubStatus = "修改管理员资料"
			Call Main("modify")
		Case "dofreeze"
			Call DoFreeze()
		Case "docreate"
			Call DoCreate()
		Case "domodify"
			Call DoModify()
		Case "dodelete"
			Call DoDelete()
		Case Else
			SubStatus = "团队成员列表"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'冻结用户
Function DoFreeze()
	Dim st: st = Request("state")
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id参数错误", "REFRESH")
	Dim objA: Set objA = New ClassAdmin: objA.ID = id
	Select Case LCase(st)
		Case "freeze"
			If objA.Freeze Then
				Call WebLog("冻结用户["& id &"]成功！", "SESSION")
				Call MsgAndGo("冻结用户["& id &"]成功", "REFRESH")
			Else
				Call MsgBox("错误：" & objA.LastError, "BACK")
			End If
		Case "unfreeze"
			If objA.Unfreeze Then
				Call WebLog("解冻用户["& id &"]成功！", "SESSION")
				Call MsgAndGo("解冻用户["& id &"]成功", "REFRESH")
			Else
				Call MsgBox("错误：" & objA.LastError, "BACK")
			End If
		Case Else
			Call MsgBox("参数错误！", "BACK")
	End Select
End Function

'创建团队
Function DoCreate()
	If Len(Request("fUsername")) = 0 Then Call MsgBox("用户名不能为空!", "BACK")
	If Len(Request("fNickname")) = 0 Then Call MsgBox("昵称不能为空!", "BACK")
	If Len(Request("fPassword")) < 6 Then Call MsgBox("密码不能少于6位!", "BACK")
	If Request("fPassword")<>Request("fRePassword") Then Call MsgBox("密码不一致!", "BACK")
	Dim objA: Set objA = New ClassAdmin
	If objA.SetValue = False Then
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	If objA.Create Then
		Call WebLog("创建管理员["&objA.Username&"]成功！", "SESSION")
		Call MsgAndGo("创建管理员["&objA.Username&"]成功！", "BACK")
	Else
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Function

'删除团队
Sub DoModify()
	Dim objA: Set objA = New ClassAdmin
	objA.ID = id
	If Len(Request("fPassword"))>0 Then
		If Request("fPassword")<>Request("fRePassword") Then Call MsgBox("密码不一致!", "BACK")
		If objA.SetValue = False Then
			Call MsgBox("错误：" & objA.LastError, "BACK")
		End If
		If objA.ModifyPsw Then
			Call WebLog("修改管理员["& objA.Username &"]成功！", "SESSION")
			Call MsgAndGo("修改管理员["& objA.Username &"]成功！", "admin_user.asp")
		Else
			Call MsgBox("错误：" & objA.LastError, "BACK")
		End If
	Else
		If objA.SetValue = False Then
			Call MsgBox("错误：" & objA.LastError, "BACK")
		End If
		If objA.ModifyInfo Then
			Call WebLog("修改管理员["& objA.Username &"]信息成功，但没修改密码！", "SESSION")
			Call MsgAndGo("您修改管理员["& objA.Username &"]信息成功，但没修改密码！", "admin_user.asp")
		Else
			Call MsgBox("错误：" & objA.LastError, "BACK")
		End If
	End If
	Set objA = Nothing
End Sub

'删除团队
Sub DoDelete()
	Dim objA: Set objA = New ClassAdmin
	objA.ID = id
	If objA.Delete Then
		Call WebLog("删除管理员["& id &"]成功！", "SESSION")
		Call MsgAndGo("删除管理员["& id &"]成功！", "admin_user.asp")
	Else
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Sub

'管理员相应权限
Function GetLevel(Byval iLevel)
	Dim tLevel
	Select Case iLevel
		Case -1
			tLevel = "<font color='blue'>冻结用户</font>"
		Case 0
			tLevel = "普通管理员"
		Case 1
			tLevel = "中级管理员"
		Case 2
			tLevel = "高级管理员"
		Case 3
			tLevel = "超级管理员"
		Case Else
			tLevel = iLevel
	End Select
	GetLevel = tLevel
End Function

'主函数
Sub Main(ByVal strType)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>后台管理 - 团队管理 - <%=SYS%></title>
<script type="text/javascript" src="inc/base.js"></script>
<link href="images/common.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.green {color:green;}
.red{ color:#F00;}
.blue{ color:blue;}
.gray{ color:gray;}
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
-->
</style>
</head>

<body>
<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td colspan="2" id="topNav">
			<%Call TopNav("user")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">管理首页</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== 团队管理 ==--</li>
                 <li <%If Request("action") = "create" Then Response.Write("class=""on""")%>><a href="?action=create">添加管理员</a></li>
                 <li <%If Request("action") <> "create" And Request("list") <> "trash"  Then Response.Write("class=""on""")%>><a href="?action=list">管理员列表</a></li>
                </ul>
				<%Call SysInfo()%>
            </td>
            <td id="content" valign="top">
            	<div class="content">
                	<div class="status"> 您的位置：<a href="index.asp">管理首页</a> → <%=MainStatus%> → <%=SubStatus%> </div>
					<%
                        Select Case LCase(strType)
                            Case "create"
                                FuncForm(0)
                            Case "modify"
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
'团队列表， mode - 模式
Sub List()
%>
	<form name="form2" action="" method="post">
	<table class="list">
    	<tr>
        	<th><input type="checkbox" name="GroupID" value="" onClick="Checked(this.form,'GroupID',this)"/></th>
        	<th>ID</th>
            <th>用户名</th>
            <th>昵称</th>
            <th>等级</th>
            <th>登陆次数</th>
            <th>登录时间</th>
            <th>登录IP</th>
            <th>操作</th>
            <th>删除</th>
        </tr>
	<%
		Dim strSql, Rs
		strSql = "SELECT [ID],[Username],[Nickname],[Level],[LoginCount],[LoginTime],[LoginIP] FROM [Admin] ORDER BY [Level] DESC, ID"
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
            <td><%=Rs.Data("Username")%></td>
            <td><%=Rs.Data("Nickname")%></td>
            <td><%=GetLevel(Rs.Data("Level"))%></td>
            <td><%=Rs.Data("LoginCount")%></td>
            <td><%=Rs.Data("LoginTime")%></td>
            <td><%=Rs.Data("LoginIP")%></td>
            <td>
            	<%If Rs.Data("Level") = -1 Then%>
            		<a href="?action=dofreeze&state=unfreeze&id=<%=Rs.Data("ID")%>" onclick="return confirm('解冻用户将变成初级管理员！\n\n确定解冻该用户？')">解冻</a>
               <%Else%>
               		<a href="?action=modify&id=<%=Rs.Data("ID")%>">编辑</a> | 
                    <a href="?action=dofreeze&state=freeze&id=<%=Rs.Data("ID")%>" onclick="return confirm('冻结用户将不能登录后台管理！\n\n确定冻结该用户？')">冻结</a>
               <%End If%>
            </td>
            <td>
				<a href="?action=dodelete&id=<%=Rs.Data("ID")%>" onclick="return confirm('删除将不用恢复！\n\n确定永久性删除该用户？')">删除</a>
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
                    <option value="freeze"> 冻结用户 </option>
                    <option value="unfreeze"> 解冻用户 </option>
                    <option value="delete"> 彻底删除 </option>
                </select>
            </td>
        </tr>
    </table>
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
function BatchFreeze(form, isFreeze){
	var id = GetID(form);
	if(id){
		if(isFreeze){
			if (!confirm('冻结用户将不能进入后台！\n\n是否把选中用户进行冻结？')) return;
			form.action = '?action=dofreeze&state=freeze&id=' + id;
			form.submit();  
		}
		else{
			if(!confirm('解冻后用户将变成初级管理员！\n\n是否把选中用户进行解冻？')) return;
			form.action  = '?action=dofreeze&state=unfreeze&id=' + id;
			form.submit();  
		}
	}
} 

//批量删除
function BatchDelete(form){
	var id = GetID(form);
	if(!id){return;}
	if(confirm('删除将不能恢复！\n\n是否真的删除？')){	
		form.action  = '?action=dodelete&id=' + id;
		form.submit(); 
	}
} 

//批处理操作
function Dobatch(objSel){
	switch(objSel.options[objSel.selectedIndex].value){
		case 'freeze':
			BatchFreeze(objSel.form, true);
			break;
		case 'unfreeze':
			BatchFreeze(objSel.form, false);
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
<%
	Rs.Data.Close: Set Rs = Nothing
End Sub%>

<%
'团队表单
Sub FuncForm(ByVal id)
	Dim objA: Set objA = New ClassAdmin
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("对不起，你编辑的管理员不存在", "BACK")
	End If
%>
	<form action="?action=do<%If id > 0 Then Echo("modify") Else Echo("create")%>" id="form1" name="form1" method="post">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <table class="form" style="border:1px #88C4FF solid;">
            <tr><th colspan="2">
				<%If id > 0 Then Echo("编辑") Else Echo("添加")%>管理员
            </th></tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
            	<td align="right" width="15%">管理员：</td>
            	<td><input type="text" name="fUsername" value="<%=objA.Username%>" <%If id > 0 Then Echo("readonly=""readonly""")%> style="width:450px;"/> <span class="red">* 必填（只能英文和下划线，不能使用中文）</span></td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
            	<td align="right" width="15%">昵称：</td>
            	<td><input type="text" name="fNickname" value="<%=objA.Nickname%>" style="width:450px;"/> <span class="red">* 必填（可以使用中文）</span></td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right" width="15%">新密码：</td>
                <td><input type="password" name="fPassword" id="fPassword" value="" style="width:250px;"/> 
				<%If id > 0 Then%>
                	<span class="gray">不填，则不修改密码</span>
                <%Else%>
                	<span class="red"> * 必填</span>
                <%End If%>
                </td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right" width="15%">重复新密码：</td>
                <td><input type="password" name="fRePassword" id="fRePassword" value="" style="width:250px;"/> </td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right" width="15%">权限</td>
                <td>
                <select name="fLevel" style="line-height:25px; padding:5px;">
                	<option value="0"> 选择等级 </option>
                	<%If id > 0 Then%>
                    	<option value="<%=objA.Level%>" selected="selected"> <%=GetLevel(objA.Level)%> </option>
                    <%End If%>
                    <option value="0"> 普通管理员 </option>
                    <option value="1"> 中级管理员 </option>
                    <option value="2"> 高级管理员 </option>
                    <option value="3"> 超级管理员 </option>
                    <option value="4"> 冻结用户 </option>
                </select> <span class="red"> * 必选</span>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <input type="submit" class="btn" value="提交" />
                    <input type="reset" class="btn" value="重置" />
                </td>
            </tr>
        </table>
        <div style="border:dashed 1px #CCC; margin:10px 0px; padding:5px; line-height:22px;">
         <b>权限说明：</b><br />
         1、普通管理员权限：管理文章、图片、留言。<br />
         2、中级管理员权限：管理文章、图片、留言、文章栏目、图片栏目。<br />
         3、高级管理员权限：高级管理员权限 + 标签管理、DIY页面管理、系统配置、操作记录。<br />
         4、超级管理员权限：具有全部权限（即是：高级管理员权限 + 模板管理 + 团队管理）。<br />
         5、冻结用户：没有管理权限，并且不能登录本系统后台管理。
        </div>
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
    <div class="page">
    	<a href="javascript:history.go(-1)" onclick="return confirm('确定放弃编辑管理员?')"><< == 返回 << == </a>
    </div>
<%
	Set objA = Nothing
End Sub
%>
