<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_mytag.asp"-->
<%
'=========================================================
' File Name：	admin_mytag.asp
' Purpose：		自定义标签管理
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-3 16:40:50
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "<a href='?'>管理自定义标签</a>"

Call ChkLogin()	'检查登录
Call ChkPower("mytag","all") '检查权限
Call Init()		'初始化页面

'初始化页面
Sub Init()

	Select Case LCase(act)
		Case "create"
			SubStatus = "创建自定义标签"
			Call Main("create")
		Case "modify"
			SubStatus = "修改自定义标签"
			Call Main("modify")
		Case "docreate"
			Call DoCreate()
		Case "domodify"
			Call DoModify()
		Case "dodelete"
			Call DoDelete()
		Case Else
			SubStatus = "自定义标签列表"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'创建自定义标签
Function DoCreate()
	Dim objA: Set objA = New ClassMyTag
	If objA.SetValue = False Then
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	If objA.Create Then
		Call WebLog("创建自定义标签["&objA.Name&"]成功！", "SESSION")
		Call MsgAndGo("创建自定义标签["&objA.Name&"]成功！", "BACK")
	Else
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Function

'删除自定义标签
Sub DoModify()
	Dim objA: Set objA = New ClassMyTag
	objA.ID = id
	If objA.SetValue = False Then
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	If objA.Modify Then
		Call WebLog("修改自定义标签[id:"& id &"]成功！", "SESSION")
		Call MsgAndGo("修改自定义标签[id:"& id &"]成功！", "admin_mytag.asp")
	Else
		Call MsgBox("错误：" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Sub

'删除自定义标签
Sub DoDelete()
	Dim objA: Set objA = New ClassMyTag
	objA.ID = id
	If objA.Delete Then
		Call WebLog("删除自定义标签[id:"& id &"]成功！", "SESSION")
		Call MsgAndGo("删除自定义标签[id:"& id &"]成功！", "admin_mytag.asp")
	Else
		Call MsgBox("错误：" & objA.LastError, "BACK")
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
<title><%=SITENAME%>后台管理 - 自定义标签管理 - Powered by eekku.com</title>
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
			<%Call TopNav("mytag")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">管理首页</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== 自定义标签管理 ==--</li>
                 <li <%If Request("action") = "create" Then Response.Write("class=""on""")%>><a href="?action=create">添加自定义标签</a></li>
                 <li <%If Request("action") <> "create" And Request("list") <> "trash"  Then Response.Write("class=""on""")%>><a href="?action=list">管理自定义标签</a></li>
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
'自定义标签列表， mode - 模式
Sub List()
%>
	<form name="form2" action="" method="post">
	<table class="list">
    	<tr>
        	<th><input type="checkbox" name="GroupID" value="" onClick="Checked(this.form,'GroupID',this)"/></th>
        	<th>ID</th>
            <th>标签</th>
            <th>说明</th>
            <th>用法</th>
            <th>操作</th>
            <th>删除</th>
        </tr>
	<%
		Dim strSql, Rs
		strSql = "SELECT ID,Name,Info FROM [MyTags] ORDER BY ID"
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
            <td><%=Rs.Data("Name")%></td>
            <td><%=Rs.Data("Info")%></td>
            <td>{my:<%=Rs.Data("Name")%> /}</td>
            <td><a href="?action=modify&id=<%=Rs.Data("ID")%>">编辑</a></td>
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
                批量操作：
                <select name="name" onChange="Dobatch(this)" style="line-height:25px; padding:5px;">
                	<option value=""> 选择操作 </option>
                    <option value="delete"> 彻底删除 </option>
                </select>
            </td>
        </tr>
    </table>
    </form>
    <div class="page"><%=Rs.Page%></div>
    <div style="font-size:14px; background:#EEC; padding:5px; border:#993 1px dashed;">帮助：在您网站的模板页面中加入<span class="red">{my:<u class="blue">标签名</u> /}</span>即可引用该标签。</div>
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
	if(confirm('删除将不能恢复！\n\n是否真的删除？')){	
		form.action  = '?action=dodelete&id=' + id;
		form.submit(); 
	}
} 

//批处理操作
function Dobatch(objSel){
	switch(objSel.options[objSel.selectedIndex].value){
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
'自定义标签表单
Sub FuncForm(ByVal id)
	Dim objA: Set objA = New ClassMyTag
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("对不起，你编辑的自定义标签不存在", "BACK")
	End If
%>
	<form action="?action=do<%If id > 0 Then Echo("modify") Else Echo("create")%>" id="form1" name="form1" method="post">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <table class="form" style="border:1px #88C4FF solid;">
            <tr><th colspan="2">
				<%If id > 0 Then Echo("编辑") Else Echo("添加")%>自定义标签
            </th></tr>
            <tr>
            	<td align="right" width="15%">标签名：</td>
            	<td><input type="text" name="fName" value="<%=objA.Name%>" style="width:450px;"/> <span class="red">* 必填（只能英文和下划线，不能使用中文）</span></td>
            </tr>
            <tr>
                <td align="right">标签说明：</td>
                <td><input type="text" name="fInfo" value="<%=objA.Info%>" style="width:450px;"/>（可以中文）</td>
            </tr>
            <tr>
                <td align="right">代码：</td>
                <td ></td>
            </tr>
            <tr>
                <td colspan="2">
                	<div id="editor">
                    <textarea name="fCode" id="content1" style="width:100%;height:400px;"><%=Server.HTMLEncode(objA.Code)%></textarea>
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <input type="submit" class="btn" value="提交" />
                    <input type="reset" class="btn" value="重置" />
                    <input type="button" name="button" value="加载编辑器" onclick="javascript:KE.create('content1');" />
                </td>
            </tr>
        </table>
    </form>
    </div>
<script type="text/javascript" charset="utf-8" src="inc/editor/kindeditor.js"></script>
<script type="text/javascript">
//初始化编辑器
KE.init({
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
    //-->
    </script>
    <div class="page">
    	<a href="javascript:history.go(-1)" onclick="return confirm('确定放弃编辑自定义标签?')"><< == 返回 << == </a>
    </div>
<%
	Set objA = Nothing
End Sub
%>
