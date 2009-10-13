<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_mytag.asp"-->
<%
'=========================================================
' File Name��	admin_mytag.asp
' Purpose��		�Զ����ǩ����
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-3 16:40:50
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "<a href='?'>�����Զ����ǩ</a>"

Call ChkLogin()	'����¼
Call ChkPower("mytag","all") '���Ȩ��
Call Init()		'��ʼ��ҳ��

'��ʼ��ҳ��
Sub Init()

	Select Case LCase(act)
		Case "create"
			SubStatus = "�����Զ����ǩ"
			Call Main("create")
		Case "modify"
			SubStatus = "�޸��Զ����ǩ"
			Call Main("modify")
		Case "docreate"
			Call DoCreate()
		Case "domodify"
			Call DoModify()
		Case "dodelete"
			Call DoDelete()
		Case Else
			SubStatus = "�Զ����ǩ�б�"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'�����Զ����ǩ
Function DoCreate()
	Dim objA: Set objA = New ClassMyTag
	If objA.SetValue = False Then
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	If objA.Create Then
		Call WebLog("�����Զ����ǩ["&objA.Name&"]�ɹ���", "SESSION")
		Call MsgAndGo("�����Զ����ǩ["&objA.Name&"]�ɹ���", "BACK")
	Else
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Function

'ɾ���Զ����ǩ
Sub DoModify()
	Dim objA: Set objA = New ClassMyTag
	objA.ID = id
	If objA.SetValue = False Then
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	If objA.Modify Then
		Call WebLog("�޸��Զ����ǩ[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("�޸��Զ����ǩ[id:"& id &"]�ɹ���", "admin_mytag.asp")
	Else
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Sub

'ɾ���Զ����ǩ
Sub DoDelete()
	Dim objA: Set objA = New ClassMyTag
	objA.ID = id
	If objA.Delete Then
		Call WebLog("ɾ���Զ����ǩ[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("ɾ���Զ����ǩ[id:"& id &"]�ɹ���", "admin_mytag.asp")
	Else
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Sub

'������
Sub Main(ByVal strType)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>��̨���� - �Զ����ǩ���� - Powered by eekku.com</title>
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
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== �Զ����ǩ���� ==--</li>
                 <li <%If Request("action") = "create" Then Response.Write("class=""on""")%>><a href="?action=create">����Զ����ǩ</a></li>
                 <li <%If Request("action") <> "create" And Request("list") <> "trash"  Then Response.Write("class=""on""")%>><a href="?action=list">�����Զ����ǩ</a></li>
                </ul>
				<%Call SysInfo()%>
            </td>
            <td id="content" valign="top">
            	<div class="content">
                	<div class="status"> ����λ�ã�<a href="index.asp">������ҳ</a> �� <%=MainStatus%> �� <%=SubStatus%> </div>
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
'�Զ����ǩ�б� mode - ģʽ
Sub List()
%>
	<form name="form2" action="" method="post">
	<table class="list">
    	<tr>
        	<th><input type="checkbox" name="GroupID" value="" onClick="Checked(this.form,'GroupID',this)"/></th>
        	<th>ID</th>
            <th>��ǩ</th>
            <th>˵��</th>
            <th>�÷�</th>
            <th>����</th>
            <th>ɾ��</th>
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
            <td><a href="?action=modify&id=<%=Rs.Data("ID")%>">�༭</a></td>
            <td>
            	<a href="?action=dodelete&id=<%=Rs.Data("ID")%>" onclick="return confirm('ɾ�������ûָ����ݣ�\n\nȷ��������ɾ�������ݣ�')">ɾ��</a>
            </td>
        </tr>
	<%
			Rs.Data.MoveNext
		Next
	%>
        <tr>
        	<td colspan="9" style="padding:5px;">
  				<input type="button" onClick="SelectAll(this.form,'GroupID')" value="ȫѡ" /> 
                <input type="button" onClick="SelectOthers(this.form,'GroupID')" value="��ѡ" /> 
                &nbsp;&nbsp;
                ����������
                <select name="name" onChange="Dobatch(this)" style="line-height:25px; padding:5px;">
                	<option value=""> ѡ����� </option>
                    <option value="delete"> ����ɾ�� </option>
                </select>
            </td>
        </tr>
    </table>
    </form>
    <div class="page"><%=Rs.Page%></div>
    <div style="font-size:14px; background:#EEC; padding:5px; border:#993 1px dashed;">������������վ��ģ��ҳ���м���<span class="red">{my:<u class="blue">��ǩ��</u> /}</span>�������øñ�ǩ��</div>
<script language="javascript" type="text/javascript">
<!--
// ��ȫѡ����ȡ��
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

// ��ȫѡ
function SelectAll(form, name)
{
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if (name == '' || e.name == name) {
			e.checked = true;
		}
	}
}


// ����ѡ
function SelectOthers(form, name){	
	
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if (name == '' || e.name == name) {
			e.checked = !e.checked;
		}
	}
}

//��ȡID
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
		alert('��δѡ���κ�ѡ�');
		return;
	}
	return id;
}

//����ɾ��
function BatchDelete(form){
	var id = GetID(form);
	if(!id){return;}
	if(confirm('ɾ�������ָܻ���\n\n�Ƿ����ɾ����')){	
		form.action  = '?action=dodelete&id=' + id;
		form.submit(); 
	}
} 

//���������
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
'�Զ����ǩ��
Sub FuncForm(ByVal id)
	Dim objA: Set objA = New ClassMyTag
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("�Բ�����༭���Զ����ǩ������", "BACK")
	End If
%>
	<form action="?action=do<%If id > 0 Then Echo("modify") Else Echo("create")%>" id="form1" name="form1" method="post">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <table class="form" style="border:1px #88C4FF solid;">
            <tr><th colspan="2">
				<%If id > 0 Then Echo("�༭") Else Echo("���")%>�Զ����ǩ
            </th></tr>
            <tr>
            	<td align="right" width="15%">��ǩ����</td>
            	<td><input type="text" name="fName" value="<%=objA.Name%>" style="width:450px;"/> <span class="red">* ���ֻ��Ӣ�ĺ��»��ߣ�����ʹ�����ģ�</span></td>
            </tr>
            <tr>
                <td align="right">��ǩ˵����</td>
                <td><input type="text" name="fInfo" value="<%=objA.Info%>" style="width:450px;"/>���������ģ�</td>
            </tr>
            <tr>
                <td align="right">���룺</td>
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
                    <input type="submit" class="btn" value="�ύ" />
                    <input type="reset" class="btn" value="����" />
                    <input type="button" name="button" value="���ر༭��" onclick="javascript:KE.create('content1');" />
                </td>
            </tr>
        </table>
    </form>
    </div>
<script type="text/javascript" charset="utf-8" src="inc/editor/kindeditor.js"></script>
<script type="text/javascript">
//��ʼ���༭��
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
    	<a href="javascript:history.go(-1)" onclick="return confirm('ȷ�������༭�Զ����ǩ?')"><< == ���� << == </a>
    </div>
<%
	Set objA = Nothing
End Sub
%>
