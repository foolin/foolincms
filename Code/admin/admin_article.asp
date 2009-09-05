<!--#include file="inc/admin.include.asp"-->
<!--#include file="lib/class_article.asp"-->
<%
'=========================================================
' File Name��	admin_article.asp
' Purpose��		���¹���
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-3 16:40:50
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "��������"

Call ChkLogin()	'����¼
Call Init()		'��ʼ��ҳ��

'��ʼ��ҳ��
Sub Init()

	Select Case LCase(act)
		Case "create"
			SubStatus = "��������"
			Call Main("create")
		Case "modify"
			SubStatus = "�޸�����"
			Call Main("modify")
		Case "setstate"
			Call SetState()
		Case "list"
			SubStatus = "�����б�"
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
			SubStatus = "�����б�"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'����״̬
Sub SetState()
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id��������", "BACK")
	Dim state: state = Request("state")
	Select Case LCase(state)
		Case "pass"
			Call DB("UPDATE Article SET State = 1 WHERE ID = " & id, 0)
			Call WebLog("�������[id:"& id &"]�ɹ�!", "SESSION")
			Call MsgAndGo("�������[id:"& id &"]�ɹ�!", "REFRESH")
		Case "nopass"
			Call DB("UPDATE Article SET State = 0 WHERE ID = " & id, 0)
			Call WebLog("ȡ���������[id:"& id &"]�ɹ�!", "SESSION")
			Call MsgAndGo("ȡ���������[id:"& id &"]�ɹ�!", "REFRESH")
		Case "delete"
			Call DB("UPDATE Article SET State = -1 WHERE ID = " & id, 0)
			Call WebLog("ɾ������[id:"& id &"]�ɹ�!", "SESSION")
			Call MsgAndGo("ɾ������[id:"& id &"]�ɹ�!", "REFRESH")
		Case "nodelete"
			Call DB("UPDATE Article SET State = 0 WHERE ID = " & id, 0)
			Call WebLog("��ԭ����[id:"& id &"]�ɹ�!", "SESSION")
			Call MsgAndGo("��ϲ����ԭ����[id:"& id &"]�ɹ�!", "REFRESH")
		Case Else
			Call MsgBox("�Բ������Ĵ����������", "BACK")
	End Select
End Sub

'��������
Function DoCreate()
	Dim objA: Set objA = New ClassArticle
	If objA.Create Then
		Call WebLog("��������[id:title:"&objA.Title&"]�ɹ���", "SESSION")
		Call MsgAndGo("��������[id:title:"&objA.Title&"]�ɹ���", "BACK")
	Else
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Function

'ɾ������
Sub DoModify()
	Dim objA: Set objA = New ClassArticle
	objA.ID = id
	If objA.SetValue And objA.Modify Then
		Call WebLog("�޸�����[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("�޸�����[id:"& id &"]�ɹ���", "admin_article.asp")
	Else
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Sub

'ɾ������
Sub DoDelete()
	DB "Delete From [Article] Where [ID] In (" & id & ")" ,0
	Call WebLog("ɾ������[id:"& id &"]�ɹ���", "SESSION")
	Call MsgAndGo("ɾ������[id:"& id &"]�ɹ�!", "REFRESH")
End Sub

'���������
Sub DoBatch()
	Dim bat: bat = Request("batch")
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id��������", "BACK")
	Select Case LCase(bat)
		Case "pass"
			Call DB("UPDATE Article SET State = 1 WHERE ID IN (" & id & ")", 0)
			Call WebLog("�����������[id:"& id &"]�ɹ���", "SESSION")
			Call MsgAndGo("�����������[id:"& id &"]�ɹ�!", "REFRESH")
		Case "nopass"
			Call DB("UPDATE Article SET State = 0 WHERE ID IN (" & id & ")", 0)
			Call WebLog("����ȡ���������[id:"& id &"]�ɹ���", "SESSION")
			Call MsgAndGo("����ȡ���������[id:"& id &"]�ɹ�!", "REFRESH")
		Case "top"
			Call DB("UPDATE Article SET IsTop = 1 WHERE ID IN (" & id & ")", 0)
			Call WebLog("�����ö�����[id:"& id &"]�ɹ���", "SESSION")
			Call MsgAndGo("�����ö�����[id:"& id &"]�ɹ�!", "REFRESH")
		Case "notop"
			Call DB("UPDATE Article SET IsTop = 0 WHERE ID IN (" & id & ")", 0)
			Call WebLog("����ȡ���ö�����[id:"& id &"]�ɹ���", "SESSION")
			Call MsgAndGo("����ȡ���ö�����[id:"& id &"]�ɹ�!", "REFRESH")
		Case "trash"
			Call DB("UPDATE Article SET State = -1 WHERE ID IN (" & id & ")", 0)
			Call WebLog("����ɾ������[id:"& id &"]�ɹ���", "SESSION")
			Call MsgAndGo("����ɾ������[id:"& id &"]�ɹ�!", "REFRESH")
		Case "notrash"
			Call DB("UPDATE Article SET State = 0 WHERE ID IN (" & id & ")", 0)
			Call WebLog("������ԭ����[id:"& id &"]�ɹ���", "SESSION")
			Call MsgAndGo("������ԭ����[id:"& id &"]�ɹ�!", "REFRESH")
		Case "delete"
			Call DB("Delete From [Article] Where [ID] In (" & id & ")", 0)
			Call WebLog("��������ɾ������[id:"& id &"]�ɹ���", "SESSION")
			Call MsgAndGo("��������ɾ������[id:"& id &"]�ɹ�!", "REFRESH")
		Case Else
			Call MsgBox("��������", "BACK")
	End Select
End Sub


'������
Sub Main(ByVal artType)
Dim SubStatus2
Select Case LCase(Request("list"))
	Case "trash"
		SubStatus2 = " �� ����վ"
	Case "nopass"
		SubStatus2 = " �� δ���"
	Case "pass"
		SubStatus2 = " �� �Ѿ����"
	Case "all"
		SubStatus2 = " �� ȫ������"
	Case Else
		SubStatus2 = " �� ȫ������"
End Select
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>��̨���� - ���¹��� - <%=SYS%></title>
<link href="css/common.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../inc/ckeditor/ckeditor.js"></script>
<style type="text/css">
<!--
.green {color:green;}
.red{ color:#F00;}
.blue{ color:blue;}
input{ background:#FFFFFF; border:#C4E1FF #84C1FF 1px solid; padding:3px;}
.btn{ padding:3px; background:#F7FBFF;}
form{ margin:0px;}
-->
</style>
<script type="text/javascript" src="inc/base.js"></script>
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

//�����������
function BatchPass(form, isPass){
	var id = GetID(form);
	if(id){
		if(isPass){
			if (!confirm('�Ƿ��ѡ������ͨ����ˣ�')) return;
			form.action = '?action=dobatch&batch=pass&id=' + id;
			form.submit();  
		}
		else{
			if(!confirm('�Ƿ��ѡ������ȡ����ˣ�')) return;
			form.action  = '?action=dobatch&batch=nopass&id=' + id;
			form.submit();  
		}
	}
} 

//����������
function BatchTop(form, isTop){
	var id = GetID(form);
	if(id){
		if(isTop){
			if(!confirm('�Ƿ��ѡ�������ö���')) return;
			form.action  = '?action=dobatch&batch=top&id=' + id;
			form.submit();  
		}
		else{
			if(!confirm('�Ƿ��ѡ������ȡ���ö���')) return;
			form.action  = '?action=dobatch&batch=notop&id=' + id;
			form.submit();  
		}
	}
} 

//�����ƶ�������վ
function BatchTrash(form, isTrash){
	var id = GetID(form);
	if(!id){return;}
	if (isTrash){
		if (confirm('�Ƿ��ѡ�����·ŵ�����վ��')){	
			form.action  = '?action=dobatch&batch=trash&id=' + id;
			form.submit();  
		}
	}
	else{
		if (confirm('�Ƿ��ѡ�����»�ԭ��')){	
			form.action  = '?action=dobatch&batch=notrash&id=' + id;
			form.submit();  
		}
	}
} 

//����ɾ��
function BatchDelete(form){
	var id = GetID(form);
	if(!id){return;}
	if(confirm('ɾ�������ָܻ���\n\n�Ƿ����ɾ����')){	
		form.action  = '?action=dobatch&batch=delete&id=' + id;
		form.submit(); 
	}
} 

//���������
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
                 <li class="mTitle">--== ���¹��� ==--</li>
                 <li <%If Request("action") = "create" Then Response.Write("class=""on""")%>><a href="admin_article.asp?action=create">�������</a></li>
                 <li <%If Request("action") <> "create" And Request("list") <> "trash"  Then Response.Write("class=""on""")%>><a href="admin_article.asp?action=list">��������</a></li>
                 <li <%If Request("list") = "trash" Then Response.Write("class=""on""")%>><a href="admin_article.asp?action=list&list=trash">���»���վ</a></li>
                 <li class="mTitle">--== ������Ŀ ==--</li>
                 <li><a href="admin_artcolumn.asp">�����Ŀ</a></li>
                 <li><a href="admin_artcolumn.asp">������Ŀ</a></li>
                 <li><a href="admin_artcolumn.asp">��Ŀ����վ</a></li>
                </ul>
				<%Call SysInfo()%>
            </td>
            <td id="content" valign="top">
            	<div class="content">
                	<div class="status"> ����λ�ã�<a href="index.asp">������ҳ</a> �� <%=MainStatus%> �� <%=SubStatus%> <%=SubStatus2%> </div>
                    <div style="font-size:14px; line-height:25px; padding-left:5px;">
                        <a href="?list=all">ȫ������</a> | 
                        <a href="?list=pass">�Ѿ����</a> | 
                        <a href="?list=nopass">δ���</a> |
                        <a href="?list=trash">����վ</a> |
                        <a href="?action=create">�������</a>
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
'�����б� mode - ģʽ
Sub List()
Dim mode: mode = LCase(Request("list"))
%>
	<form name="form1" action="" method="post">
	<table class="list">
    	<tr>
        	<th><input type="checkbox" name="GroupID" value="" onClick="Checked(this.form,'GroupID',this)"/></th>
        	<th>ID</th>
            <th>��Ŀ</th>
            <th>����</th>
            <th>����</th>
            <th>ʱ��</th>
            <th>״̬</th>
            <th>����</th>
            <th>ɾ��</th>
        </tr>
	<%
		Dim strSql, Rs
		Select Case mode
			Case "trash"
				strSql = "SELECT * FROM [Article] WHERE State = -1 ORDER BY IsTop DESC,ID DESC"
			Case "nopass"
				strSql = "SELECT * FROM [Article] WHERE State = 0 ORDER BY IsTop DESC,ID DESC"
			Case "pass"
				strSql = "SELECT * FROM [Article] WHERE State = 1 ORDER BY IsTop DESC,ID DESC"
			Case "all"
				strSql = "SELECT * FROM [Article] WHERE State > -1 ORDER BY IsTop DESC,ID DESC"
			Case Else
				strSql = "SELECT * FROM [Article] WHERE State > -1 ORDER BY IsTop DESC,ID DESC"
		End Select
		Set Rs = DB(strSql, 1)
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
            <td><%=GetColName(Rs.Data("ColID"), "article")%></td>
			<td>
            	<a href="admin_article.asp?action=modify&id=<%=Rs.Data("ID")%>"><%=Rs.Data("Title")%></a>
				<%If Rs.Data("IsTop") = 1 Then Response.Write(" <font color=""red"">[��]</font>")%>
            </td>
            <td><%=Rs.Data("Author")%></td>
            <td><%=FDate(Rs.Data("CreateTime"), 2)%></td>
            <td>
            	<%If Rs.Data("State") = 1 Then%>
            		<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=nopass" title="���ȡ�����" class="green">�����</a>
                <%ElseIf Rs.Data("State") = 0 Then%>
                	<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=pass" title="���ͨ�����" class="red">δ���</a>
                <%Else%>
					<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=nodelete" title="���ȡ��ɾ��"  onclick="return confirm('ȷ����ԭ���ݣ�')" class="blue">��ɾ��</a>
                <%End If%>
            </td>
            <td>
            	<%If Rs.Data("State") > -1 Then%>
            		<a href="?action=modify&id=<%=Rs.Data("ID")%>">�༭</a>
                <%Else%>
					            		<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=nodelete" onclick="return confirm('�ָ������º�Ϊ[δ���]״̬����ȷ����ԭ���ݣ�')">��ԭ</a>
                <%End If%>
            </td>
            <td>
            	
            	<%If Rs.Data("State") > -1 Then%>
            		<a href="?action=setstate&id=<%=Rs.Data("ID")%>&state=delete" onclick="return confirm('ȷ���Ѹ����·ŵ�����վ��')">ɾ��</a>
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
  				<input type="button" onClick="SelectAll(this.form,'GroupID')" value="ȫѡ" /> 
                <input type="button" onClick="SelectOthers(this.form,'GroupID')" value="��ѡ" /> 
                &nbsp;&nbsp;
                ����������
                <select name="name" onChange="Dobatch(this)" style="line-height:25px; padding:5px;">
                	<option value=""> ѡ����� </option>
                    <%If Request("list") = "trash" Then%>
                    <option value="notrash"> ��ԭ </option>
                    <%Else%>
                    <option value="pass"> ͨ����� </option>
                    <option value="nopass"> ȡ����� </option>
                    <option value="top"> �����ö� </option>
                    <option value="notop"> ȡ������ </option>
                    <option value="trash"> ɾ�� </option>
                    <%End If%>
                    <option value="delete"> ����ɾ�� </option>
                </select>
            </td>
        </tr>
    </table>
    </form>
    <div class="page"><%=Rs.Page%></div>
<%
	Rs.Data.Close: Set Rs = Nothing
End Sub%>

<%
'���±�
Sub ArtForm(ByVal id)
	Dim objA: Set objA = New ClassArticle
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("�Բ�����༭�����²�����", "BACK")
	End If
%>
	<form action="?action=do<%If id > 0 Then Echo("modify") Else Echo("create")%>" method="post" class="form">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <table class="list">
            <tr><th colspan="2">
				<%If id > 0 Then Echo("�༭") Else Echo("���")%>����
            </th></tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
            	<td align="right">���⣺</td>
            	<td><input type="text" name="Title" value="<%=objA.Title%>" style="width:450px;"/> <span class="red">* ����</span></td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">��Ŀ��</td>
                <td>
                	<select name="ColID">
                    	<%If id > 0 Then%>
                    		<option value="<%=objA.ColID%>"> => <%=GetColName(objA.ColID, "article")%> <= </option>
                        <%Else%>
                        	<option value="0"> => ��ѡ����Ŀ <= </option>
                        <%End If%>
                    	<%Call MainColumn()%>
                    </select>
                    <span class="red">* ��ѡ</span>
                </td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">���ߣ�</td>
                <td><input type="text" name="Author" value="<%=objA.Author%>" /></td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">��Դ��</td>
                <td><input type="text" name="Source" value="<%=objA.Source%>" /></td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">����ͼƬ��</td>
                <td><input type="text" name="PicPath" value="<%=objA.PicPath%>" style="width:450px;" />�ϴ��ļ�</td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">�ؼ��ʣ�</td>
                <td><input type="text" value="<%=objA.Keywords%>" name="Keywords"  style="width:450px;" /></td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">ѡ�</td>
                <td>
                	�ö�<input type="checkbox" name="IsTop" value="1"  <%If objA.IsTop = 1 Then Echo("checked=""checked""")%> />  
                	ͨ�����<input type="checkbox" name="State" value="1" <%If objA.State = 1 Then Echo("checked=""checked""")%> />
                </td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">��ת��ַ��</td>
                <td><input type="text" name="JumpUrl" id="JumpUrl" onfocus="chkJumpUrl()" onblur="chkJumpUrl()" value="<%=objA.JumpUrl%>"  style="width:450px;" /><span class="blue">��[����] �� [��ת��ַ] ����ֻ��ѡ��һ��</span></td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">���ݣ�</td>
                <td></td>
            </tr>
            <tr  id="editor">
                <td colspan="2">
                    <textarea class="ckeditor" cols="80" id="Content" name="Content" rows="50">
                        <%=objA.Content%>
                    </textarea>
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
    <div class="page">
    	<a href="javascript:history.go(-1)" onclick="return confirm('ȷ�������༭����?')"><< == ���� << == </a>
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
	if(oInputs.item(i).name == "JumpUrl")
	{
		oInputs.item(i).onfocus = function(){ $("editor").style.display = "none";}
		oInputs.item(i).onblur = chkJumpUrl;
	}
}
chkJumpUrl(); 	//ִ��
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
//-->
</script>
<%
	Set objA = Nothing
End Sub

'��һ����Ŀ����
Function MainColumn()
	Dim Rs
	Set Rs = DB( "SELECT * FROM ArtColumn WHERE ParentID = 0", 1)
	If Not Rs.Eof Then
		Do While Not Rs.Eof
			Echo("<option value=""" & Rs("ID") & """>" & Rs("Name") & "</option>" & Chr(10) & Chr(9) & Chr(9))
			Call SubColumn(Rs("ID"),"|-") 'ѭ���Ӽ�����
		Rs.MoveNext
		If Rs.Eof Then Exit Do '���������ѭ��
		Loop
	End If
	Rs.Close: Set Rs = Nothing
End Function
'����Ŀ����
Function SubColumn(FID,StrDis)
	Dim Rs1
	Set Rs1 = DB("SELECT * FROM ArtColumn WHERE ParentID = " & FID, 1)
	If Not Rs1.Eof Then
		Do While Not Rs1.Eof
			Echo("<option value=""" & Rs1("ID") & """>" & StrDis & Rs1("Name") & "</option>" & Chr(10) & Chr(9))
			Call SubColumn(Trim(Rs1("ID")),"| " & Strdis) '�ݹ��Ӽ�����
		Rs1.Movenext:Loop
		If Rs1.Eof Then
			Rs1.Close: Set Rs1 = Nothing
			Exit Function
		End If
	End If
	Rs1.Close: Set Rs1 = Nothing
End Function
%>
