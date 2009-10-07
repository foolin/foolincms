<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_mytag.asp"-->
<%
'=========================================================
' File Name：	admin_mytag.asp
' Purpose：		操作日志管理
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-3 16:40:50
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "<a href='?'>管理操作日志</a>"

Call ChkLogin()	'检查登录

Call Init()		'初始化页面

'初始化页面
Sub Init()
	Select Case LCase(act)
		Case "doclear"
			Call ChkPower("weblog","delete") '检查权限
			Call DoClear()
		Case "dodelete"
			Call ChkPower("weblog","delete") '检查权限
			Call DoDelete()
		Case Else
			SubStatus = "操作日志列表"
			Call Main()
	End Select
	Call ConnClose()
End Sub

'清空操作日志
Sub DoClear()
	If ClearWebLog() = True Then
		'Call WebLog("清空全部操作日志成功！", "SESSION")
		Call MsgAndGo("清空全部操作日志成功！", "admin_weblog.asp")
	End If
End Sub

'删除操作日志
Sub DoDelete()
	If DelWebLog(id) = True Then
		'Call WebLog("删除操作日志[id:"& id &"]成功！", "SESSION")
		Call MsgAndGo("删除操作日志[id:"& id &"]成功！", "admin_weblog.asp")
	End If
End Sub

'主函数
Sub Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>后台管理 - 操作日志管理 - <%=SYS%></title>
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
			<%Call TopNav("weblog")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">管理首页</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== 操作日志管理 ==--</li>
                 <li <%If Request("action") <> "create" And Request("list") <> "trash"  Then Response.Write("class=""on""")%>><a href="?action=list">管理操作日志</a></li>
                 <li><a href="?action=doclear" onclick="return confirm('清空全部日志将不用恢复数据！\n\n确定永久性清空全部日志？')">清空操作日志</a></li>
                </ul>
				<%Call SysInfo()%>
            </td>
            <td id="content" valign="top">
            	<div class="content">
                	<div class="status"> 您的位置：<a href="index.asp">管理首页</a> → <%=MainStatus%> → <%=SubStatus%> </div>
					
                    
	<form name="form2" action="" method="post">
	<table class="list">
    	<tr>
        	<th width="5%"><input type="checkbox" name="GroupID" value="" onClick="Checked(this.form,'GroupID',this)"/></th>
            <th width="30%">操作</th>
            <th width="10%">用户</th>
            <th width="12%">IP</th>
            <th width="20%">页面</th>
            <th>时间</th>
            <th width="8%">删除</th>
        </tr>
	<%
		Dim strSql, Rs
		strSql = "SELECT * FROM [WebLog] ORDER BY ID DESC"
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
            <td><%=Rs.Data("UserAction")%></td>
            <td><%=Rs.Data("Username")%></td>
            <td><%=Rs.Data("UserIP")%></td>
            <td><%=Rs.Data("ActionUrl")%></td>
            <td><%=Rs.Data("CreateTime")%></td>
            <td>
            	<a href="?action=dodelete&id=<%=Rs.Data("ID")%>" onclick="return confirm('删除将不用恢复数据！\n\n确定永久性删除该数据？')">删除</a>
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
                操作：
                <input type="button" onClick="BatchDelete(this.form)" value="删除选中项" />
                <input type="button" onClick="ClearAll(this.form)" value="清空全部日志" />
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

//批量删除
function BatchDelete(form){
	var id = GetID(form);
	if(!id){return;}
	if(confirm('删除将不能恢复！\n\n是否真的删除选中的记录？')){	
		form.action  = '?action=dodelete&id=' + id;
		form.submit(); 
	}
}
//批量删除
function ClearAll(form){
	if(confirm('清空全部日志将不能恢复！\n\n是否真的清空所有日志记录？')){	
		form.action  = '?action=doclear';
		form.submit(); 
	}
}
//-->
</script>
<%	Rs.Data.Close: Set Rs = Nothing %>
                    
                    
                    
                </div>
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>
<%End Sub%>
