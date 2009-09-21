<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_diypage.asp"-->
<%
'=========================================================
' File Name��	admin_diypage.asp
' Purpose��		�Զ���ҳ�����
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-3 16:40:50
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "�����Զ���ҳ��"

Call ChkLogin()	'����¼
Call Init()		'��ʼ��ҳ��

'��ʼ��ҳ��
Sub Init()

	Select Case LCase(act)
		Case "create"
			SubStatus = "�����Զ���ҳ��"
			Call Main("create")
		Case "modify"
			SubStatus = "�޸��Զ���ҳ��"
			Call Main("modify")
		Case "setstate"
			Call SetState()
		Case "docreate"
			Call DoCreate()
		Case "domodify"
			Call DoModify()
		Case "dodelete"
			Call DoDelete()
		Case Else
			SubStatus = "�Զ���ҳ���б�"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'����״̬
Sub SetState()
	Dim st: st = Request("state")
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id��������", "REFRESH")
	Select Case LCase(st)
		Case "show"
			Call DB("UPDATE DiyPage SET State=1 WHERE ID IN ("& id &")" ,0)
			Call WebLog("����["& id &"]״̬Ϊ[��ʾ]�ɹ���", "SESSION")
			Call MsgAndGo("����["& id &"]״̬Ϊ[��ʾ]�ɹ�", "REFRESH")
		Case "hide"
			Call DB("UPDATE DiyPage SET State=0 WHERE ID IN ("& id &")" ,0)
			Call WebLog("����["& id &"]״̬Ϊ[����]�ɹ���", "SESSION")
			Call MsgAndGo("����["& id &"]״̬Ϊ[����]�ɹ�", "REFRESH")
		Case Else
			Call MsgBox("��������", "BACK")
	End Select
End Sub

'�����Զ���ҳ��
Function DoCreate()
	Dim objA: Set objA = New ClassDiyPage
	If objA.SetValue And objA.Create Then
		Call WebLog("�����Զ���ҳ��[title:"&objA.Title&"]�ɹ���", "SESSION")
		Call MsgAndGo("�����Զ���ҳ��[title:"&objA.Title&"]�ɹ���", "REFRESH")
	Else
		Call MsgBox("����" & objA.LastError, "REFRESH")
	End If
	Set objA = Nothing
End Function

'ɾ���Զ���ҳ��
Sub DoModify()
	Dim objA: Set objA = New ClassDiyPage
	objA.ID = id
	If objA.SetValue And objA.Modify Then
		Call WebLog("�޸��Զ���ҳ��[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("�޸��Զ���ҳ��[id:"& id &"]�ɹ���", "admin_diypage.asp")
	Else
		Call MsgBox("����" & objA.LastError, "REFRESH")
	End If
	Set objA = Nothing
End Sub

'ɾ���Զ���ҳ��
Sub DoDelete()
	Dim objA: Set objA = New ClassDiyPage
	objA.ID = id
	If objA.Delete Then
		Call WebLog("ɾ���Զ���ҳ��[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("ɾ���Զ���ҳ��[id:"& id &"]�ɹ���", "admin_diypage.asp")
	Else
		Call MsgBox("����" & objA.LastError, "REFRESH")
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
<title><%=SITENAME%>��̨���� - �Զ���ҳ����� - <%=SYS%></title>
<script type="text/javascript" src="inc/base.js"></script>
<link href="css/common.css" rel="stylesheet" type="text/css" />
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
			<%Call TopNav("diypage")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== �Զ���ҳ����� ==--</li>
                 <li <%If Request("action") = "create" Then Response.Write("class=""on""")%>><a href="?action=create">����Զ���ҳ��</a></li>
                 <li <%If Request("action") <> "create" And Request("list") <> "trash"  Then Response.Write("class=""on""")%>><a href="?action=list">�����Զ���ҳ��</a></li>
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
'�Զ���ҳ���б� mode - ģʽ
Sub List()
%>
	<form name="form2" action="" method="post">
	<table class="list">
    	<tr>
        	<th><input type="checkbox" name="GroupID" value="" onClick="Checked(this.form,'GroupID',this)"/></th>
        	<th>ID</th>
            <th>����</th>
            <th>�ļ���</th>
            <th>ģ��</th>
            <th>״̬</th>
            <th>ϵͳ</th>
            <th>���ʵ�ַ</th>
            <th>����</th>
            <th>ɾ��</th>
        </tr>
	<%
		Dim strSql, Rs
		strSql = "SELECT * FROM [DiyPage] ORDER BY IsSystem DESC,ID"
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
            <td><%=Rs.Data("Title")%></td>
			<td><%=Rs.Data("PageName")%></td>
            <td><%=Rs.Data("Template")%></td>
            <td>
            	<%If Rs.Data("State") = 1 Then%>
            		<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=hide" title="�������Ϊ����" class="green">��ʾ</a>
                <%Else%>
					<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=show" title="�������Ϊ��ʾ"  onclick="return confirm('ȷ����ʾ��ҳ�棿')" class="blue">����</a>
                <%End If%>
            </td>
            <td>
            	<%If Rs.Data("IsSystem") = 1 Then%>
            		<span class="red">��</span>
                <%Else%>
					��
                <%End If%>
            </td>
            <td>diypage.asp?id=<%=Rs.Data("ID")%> [<a href="../diypage.asp?id=<%=Rs.Data("ID")%>" target="_blank">���]</a></td>
            <td><a href="?action=modify&id=<%=Rs.Data("ID")%>">�༭</a></td>
            <td>
				<%If Rs.Data("IsSystem") = 1 Then%>
                	<span class="gray">ɾ��</span>
                <%Else%>
                    <a href="?action=dodelete&id=<%=Rs.Data("ID")%>" onclick="return confirm('ɾ�������ûָ����ݣ�\n\nȷ��������ɾ�������ݣ�')">ɾ��</a>
                <%End If%>
            </td>
        </tr>
	<%
			Rs.Data.MoveNext
		Next
	%>
        <tr>
        	<td colspan="9" style="padding:5px;">
  				<input type="button" onClick="selectAll(this.form,'GroupID')" value="ȫѡ" /> 
                <input type="button" onClick="selectOthers(this.form,'GroupID')" value="��ѡ" /> 
                &nbsp;&nbsp;
                ����������
                <select name="name" onChange="dobatch(this)" style="line-height:25px; padding:5px;">
                	<option value=""> ѡ����� </option>
                    <option value="show"> ��ʾ </option>
                    <option value="hide"> ���� </option>
                    <option value="delete"> ����ɾ�� </option>
                </select>
            </td>
        </tr>
    </table>
    </form>
    <div class="page"><%=Rs.Page%></div>
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
function selectAll(form, name)
{
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if (name == '' || e.name == name) {
			e.checked = true;
		}
	}
}


// ����ѡ
function selectOthers(form, name){	
	
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if (name == '' || e.name == name) {
			e.checked = !e.checked;
		}
	}
}

//��ȡID
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
		alert('��δѡ���κ�ѡ�');
		return;
	}
	return id;
}

//������ʾ
function batchShow(form){
	var id = getID(form);
	if(!id){return;}
	if(confirm('ȷ����ѡ��ѡ������Ϊ��ʾ��')){	
		form.action  = '?action=setstate&state=show&id=' + id;
		form.submit(); 
	}
}

//��������
function batchHide(form){
	var id = getID(form);
	if(!id){return;}
	if(confirm('ȷ����ѡ��ѡ������Ϊ���أ�')){	
		form.action  = '?action=setstate&state=hide&id=' + id;
		form.submit(); 
	}
} 

//����ɾ��
function batchDelete(form){
	var id = getID(form);
	if(!id){return;}
	if(confirm('ɾ�������ָܻ���\n\n�Ƿ����ɾ����')){	
		form.action  = '?action=dodelete&id=' + id;
		form.submit(); 
	}
} 

//���������
function dobatch(objSel){
	switch(objSel.options[objSel.selectedIndex].value){
		case 'show':
			batchShow(objSel.form);
			break;
		case 'hide':
			batchHide(objSel.form);
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
'�Զ���ҳ���
Sub FuncForm(ByVal id)
	Dim objA: Set objA = New ClassDiyPage
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("�Բ�����༭���Զ���ҳ�治����", "REFRESH")
	End If
%>
	<form action="?action=do<%If id > 0 Then Echo("modify") Else Echo("create")%>" id="form1" name="form1" method="post">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <table class="form" style="border:1px #88C4FF solid;">
            <tr><th colspan="2">
				<%If id > 0 Then Echo("�༭") Else Echo("���")%>�Զ���ҳ��
            </th></tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
            	<td align="right" width="15%">���⣺</td>
            	<td><input type="text" name="fTitle" value="<%=objA.Title%>" style="width:450px;"/> <span class="red">* ����</span></td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">�ļ�����</td>
                <td><input type="text" name="fPageName" value="<%=objA.PageName%>" />
                    
                </td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">ģ������</td>
                <td><input type="text" name="fTemplate" value="<%=objA.Template%>" /></td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">�ؼ��ʣ�</td>
                <td><input type="text" name="fKeywords" value="<%=objA.Keywords%>" style="width:450px;"/></td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">ѡ�</td>
                <td>
                	��ʾ<input type="checkbox" name="fState" value="1" <%If objA.State = 1 Then Echo("checked=""checked""")%> />
                    <%If objA.IsSystem = 1 Then%>
						<span class="blue">����ҳ����ϵͳҳ�棬ֻ�ܱ༭������ɾ������</span><input type="hidden" name="fIsSystem" value="1"/> 
                    <%Else%>
						ϵͳҳ��<input type="checkbox" name="fIsSystem" id="fIsSystem" value="1" onclick="chkIsSystem()" /> 
                    <%End If%>
                	  
                </td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">���룺</td>
                <td ></td>
            </tr>
            <tr>
                <td colspan="2">
                	<div id="editor">
                    <textarea id="Content1" name="fCode" style="width:100%;height:550px;visibility:hidden;">
                    	<%=objA.Code%>
                    </textarea>
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <input type="submit" class="btn" value="�ύ" />
                    <input type="reset" class="btn" value="����" />
                </td>
            </tr>
        </table>
    </form>
    </div>
    <div class="page">
    	<a href="javascript:history.go(-1)" onclick="return confirm('ȷ�������༭�Զ���ҳ��?')"><< == ���� << == </a>
    </div>
<script type="text/javascript" charset="utf-8" src="./../inc/editor/kindeditor.js"></script>
<script type="text/javascript">
//��ʼ���༭��
KE.show({
	id : 'Content1',
	cssPath : './../inc/editor/editor.css'
});
function chkIsSystem(){
	if($("fIsSystem").checked == true){
		if(!confirm("����Ϊϵͳҳ�潫����ɾ���������ز�����\n\nȷ������ϵͳҳ�棿")){
			$("fIsSystem").checked = false;
		}
	}
}
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
