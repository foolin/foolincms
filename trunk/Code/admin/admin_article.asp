<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_article.asp"-->
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
Call ChkPower("article","all")	'����Ƿ�ӵ�й���Ȩ��
Call Init()		'��ʼ��ҳ��

'��ʼ��ҳ��
Sub Init()

	Select Case LCase(act)
		Case "create"
			SubStatus = "��������"
			If IsNullColumn = True Then
				Call MsgBox("��δ���κ���Ŀ�����������Ŀ!","admin_artcolumn.asp?action=create")
			End If
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
	If objA.SetValue = False Then
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
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
	If objA.SetValue = False Then
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	If objA.Modify Then
		Call WebLog("�޸�����[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("�޸�����[id:"& id &"]�ɹ���", "admin_article.asp")
	Else
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Sub

'ɾ������
Sub DoDelete()
	Dim objA: Set objA = New ClassArticle
	objA.ID = id
	If objA.Delete Then
		Call WebLog("ɾ������[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("ɾ������[id:"& id &"]�ɹ���", "REFRESH")
	Else
		Call MsgBox("����" & objC.LastError, "BACK")
	End If
	Set objA = Nothing
End Sub

'���������
Sub DoBatch()
	Dim bat: bat = Request("batch")
	Dim colId: colId = Request("colid")
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id��������", "BACK")
	Select Case LCase(bat)
		Case "move"
			If Len(colId) = 0 Or Not IsNumeric(colId) Then Call MsgBox("��Ŀid��������", "BACK")
			Call DB("UPDATE Article SET ColID = "& colId &" WHERE ID IN (" & id & ")", 0)
			Call WebLog("�����ƶ�����[id:"& id &"]To��Ŀ["&colId&"]�ɹ���", "SESSION")
			Call MsgAndGo("�����ƶ�����[id:"& id &"]To��Ŀ["&colId&"]�ɹ�!", "REFRESH")
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

'�����Ŀ�Ƿ�Ϊ��
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
		SubStatus2 = ""
	Case Else
		SubStatus2 = ""
End Select
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>��̨���� - ���¹��� - Powered by eekku.com</title>
<link href="images/common.css" rel="stylesheet" type="text/css" />
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

function BatchMove(form){
	var id = GetID(form);
	if(!id){return;}
	var colid = $("toColId").value;
	if ( parseInt(colid) == 0){
		alert('��ѡ����Ŀ');
		return;
	}
	if (confirm('ȷ�������ƶ�������Ŀ��')){	
		form.action  = '?action=dobatch&batch=move&id=' + id + '&colid=' + colid;
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
		case 'move':
			$("batMove").style.display = "block";
			break;
		default:
			return false;
	}
	objSel.selectedIndex = 0;
}


//��������
function soArticle(){
	var jumpUrl;
	jumpUrl = 'admin_article.asp?colid=' + $('sColId').value;
	if ($('sKeyword').value != "" && $('sKeyword').value !="������ؼ���"){
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
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== ���¹��� ==--</li>
                 <li <%If Request("action") = "create" Then Response.Write("class=""on""")%>><a href="admin_article.asp?action=create">�������</a></li>
                 <li <%If Request("action") <> "create" And Request("list") <> "trash"  Then Response.Write("class=""on""")%>><a href="admin_article.asp?action=list">��������</a></li>
                 <li <%If Request("list") = "trash" Then Response.Write("class=""on""")%>><a href="admin_article.asp?action=list&list=trash">���»���վ</a></li>
                 <li class="mTitle">--== ������Ŀ ==--</li>
                 <li><a href="admin_artcolumn.asp?action=create">�����Ŀ</a></li>
                 <li><a href="admin_artcolumn.asp">������Ŀ</a></li>
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
	<form name="form2" action="" onsubmit="return soArticle();" method="post">
	<table class="list">
    	<tr>
        	<th><input type="checkbox" name="GroupID" value="" onClick="Checked(this.form,'GroupID',this)"/></th>
        	<th>ID</th>
            <th>����</th>
            <th>��Ŀ</th>
            <th>����</th>
            <th>ʱ��</th>
            <th>״̬</th>
            <th>����</th>
            <th>ɾ��</th>
        </tr>
	<%
		Dim strSql, Rs
		Dim colId, sqlColId, strKeyword, sqlKeyword
		'��ĿID
		colId = Request("colid")
		If Len(Request("colid")) = 0 Then colId = 0
		If colId > 0 Then sqlColId = " And ColID IN ("& GetColIds(colId,"ARTICLE") &") "
		'�����ַ���
		strKeyword = Trim(Request("keyword"))
		If Len(strKeyword) > 0 Then sqlKeyword = " And Title LIKE '%"& strKeyword &"%' "
		'�����б�����
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
				<%If Rs.Data("IsTop") = 1 Then Echo(" <font color=""red"">[��]</font>")%>
                <%If Rs.Data("IsFocusPic") =1 And Rs.Data("FocusPic") <> "" Then Echo(" <font color=""red"">[ͼ]</font>")%>
            </td>
            <td><a href="?colid=<%=Rs.Data("ColID")%>"><%=GetColName(Rs.Data("ColID"), "article")%></a></td>
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
                    <option value="move"> �����ƶ� </option>
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
                
                <div class="openWin" id="batMove">
                        <div class="title">��ѡ�����</div>
                        <div class="content">
                       
                            ��ѡ�����Ŀ��
                            <select name="toColId" id="toColId">
                                  <option value="0"> ��ѡ����Ŀ </option>
                                    <%Call MainColumn()%>
                            </select>
                            <input type="button" value="�ƶ�" onclick="BatchMove(this.form);" />
                            <br /> <br />
                        </div>
                        <div class="close"><a href="#" onclick="$('batMove').style.display='none';">[��] �رմ���</a></div>
                </div>
                
                 &nbsp; ������<select name="sColId" id="sColId">
                 <option value="0"> ��ѡ����Ŀ </option>
                 <option value="0"> ȫ����Ŀ </option>
                    <%Call MainColumn()%>
                 </select>
                 <input type="text" name="sKeyword" id="sKeyword" value="<%If Len(Request("keyword"))>0 Then Echo(Request("keyword")) Else Echo("������ؼ���")%>" onclick="if(this.value=='������ؼ���')this.value='';" />
                 <input type="button" value="����" onclick="soArticle();" />
                 
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
	Else
		objA.Author = Session("AdminName")
	End If
%>
	<form action="?action=do<%If id > 0 Then Echo("modify") Else Echo("create")%>" id="form1" name="form1" method="post">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <input type="hidden" name="Hits" value="<%=objA.Hits%>"/>
        <table class="form" style="border:1px #88C4FF solid;">
            <tr><th colspan="2">
				<%If id > 0 Then Echo("�༭") Else Echo("���")%>����
            </th></tr>
            <tr>
            	<td align="right" width="15%">���⣺</td>
            	<td><input type="text" name="Title" value="<%=objA.Title%>" style="width:450px;"/> <span class="red">* ����</span></td>
            </tr>
            <tr>
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
            <tr>
                <td align="right">���ߣ�</td>
                <td><input type="text" name="Author" value="<%=objA.Author%>" /></td>
            </tr>
            <tr>
                <td align="right">��Դ��</td>
                <td><input type="text" name="Source" value="<%=objA.Source%>" /></td>
            </tr>
            <tr>
                <td align="right">����ͼƬURL��</td>
                <td>
                	<input type="text" name="FocusPic" id="FocusPic" value="<%=objA.FocusPic%>" style="width:450px;" /> <a href="javascript:uploadFocusPic();">�ϴ�ͼƬ</a>
                    <div id="uploadFocusPic" style="display:none;">
                    <iframe frameborder="0" src="inc/upload_focuspic.asp" width="80%" height="30"></iframe>
                    </div>
                </td>
            </tr>
            <tr>
                <td align="right">�ؼ��ʣ�</td>
                <td><input type="text" value="<%=objA.Keywords%>" name="Keywords"  style="width:450px;" /></td>
            </tr>
            <tr>
                <td align="right">ѡ�</td>
                <td>
                	�ö�<input type="checkbox" name="IsTop" value="1"  <%If objA.IsTop = 1 Then Echo("checked=""checked""")%> />  
                	ͨ�����<input type="checkbox" name="State" value="1" <%If objA.State = 1 Then Echo("checked=""checked""")%> />
                    ����ͼƬ<input type="checkbox" name="IsFocusPic" value="1" <%If objA.IsFocusPic = 1 Then Echo("checked=""checked""")%> id="IsFocusPic" onclick="chkFocusPic()"/>
                </td>
            </tr>
            <tr>
                <td align="right">��ת��ַ��</td>
                <td><input type="text" name="JumpUrl" id="JumpUrl" value="<%=objA.JumpUrl%>"  style="width:450px;" /></td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">���ݣ�</td>
                <td ><span class="red">( <span class="green">����</span> �� <span class="green">��ת��ַ</span> ����ֻ��ѡ��һ  )</span></td>
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
                    <input type="submit" class="btn" value="�ύ" />
                    <input type="reset" class="btn" value="����" />
                </td>
            </tr>
        </table>
    </form>
    </div>
    <div class="page">
    	<a href="javascript:history.go(-1)" onclick="return confirm('ȷ�������༭����?')"><< == ���� << == </a>
    </div>
<script type="text/javascript" charset="utf-8" src="inc/editor/kindeditor.js"></script>
<script type="text/javascript">
//��ʼ���༭��
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
//chkJumpUrl(); 	//ִ��
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
		alert("����δ��д����ͼƬURL�������ϴ�ͼƬ��");
		$("IsFocusPic").checked = false;
		$("FocusPic").focus();
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
	Set Rs = DB( "SELECT * FROM ArtColumn WHERE ParentID = 0 ORDER BY Sort DESC,ID", 1)
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
	Set Rs1 = DB("SELECT * FROM ArtColumn WHERE ParentID = " & FID & " ORDER BY Sort DESC,ID", 1)
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
