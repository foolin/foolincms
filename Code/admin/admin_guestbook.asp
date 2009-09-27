<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_guestbook.asp"-->
<%
'=========================================================
' File Name��	admin_guestbook.asp
' Purpose��		���Թ������
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-27 9:31:00
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "<a href=""admin_guestbook.asp"">��������</a>"

Call ChkLogin()	'����¼
Call Init()		'��ʼ��ҳ��

'��ʼ��ҳ��
Sub Init()

	Select Case LCase(act)
		Case "recomment"
			SubStatus = "�ظ�����"
			Call Main("recomment")
		Case "setstate"
			Call SetState()
		Case "dorecomment"
			Call DoRecomment()
		Case "dodelete"
			Call DoDelete()
		Case Else
			SubStatus = "�����б�"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'����״̬
Sub SetState()
	Dim st: st = Request("state")
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id��������", "REFRESH")
	Select Case LCase(st)
		Case "pass"
			Call DB("UPDATE [GuestBook] SET State=1 WHERE ID IN ("& id &")" ,0)
			Call WebLog("�������["& id &"]�ɹ���", "SESSION")
			Call MsgAndGo("�������["& id &"]�ɹ�", "REFRESH")
		Case "nopass"
			Call DB("UPDATE [GuestBook] SET State=0 WHERE ID IN ("& id &")" ,0)
			Call WebLog("ȡ���������["& id &"]�ɹ���", "SESSION")
			Call MsgAndGo("ȡ���������["& id &"]�ɹ���", "REFRESH")
		Case Else
			Call MsgBox("��������", "BACK")
	End Select
End Sub

'�ظ�����
Function DoRecomment()
	Dim objA: Set objA = New ClassGuestBook
	objA.ID = id
	If objA.Comment Then
		Call WebLog("�ظ�����[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("�ظ�����[id:"& id &"]�ɹ���", "admin_guestbook.asp")
	Else
		Call MsgBox("����" & objA.LastError, "REFRESH")
	End If
	Set objA = Nothing
End Function

'ɾ������
Sub DoDelete()
	Dim objA: Set objA = New ClassGuestBook
	objA.ID = id
	If objA.Delete Then
		Call WebLog("ɾ������[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("ɾ������[id:"& id &"]�ɹ���", "admin_guestbook.asp")
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
<title><%=SITENAME%>��̨���� - ���Թ��� - <%=SYS%></title>
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
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== ���Թ��� ==--</li>
                 <li class="on"><a href="?action=list">��������</a></li>
                </ul>
				<%Call SysInfo()%>
            </td>
            <td id="content" valign="top">
            	<div class="content">
                	<div class="status"> ����λ�ã�<a href="index.asp">������ҳ</a> �� <%=MainStatus%> �� <%=SubStatus%> </div>
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
'�Զ���ҳ���б� mode - ģʽ
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
                <!-- ���Իظ� -->
			
                <div class="reComment">
                <%If Len(Rs.Data("Recomment"))>0 Then
					Echo("<b>"&Rs.Data("ReUser")& "</b>�ظ���" &Rs.Data("Recomment")) 
				  Else
				  	Echo("���޻ظ�")
				  End If
				%>
                </div>
                
                <!-- ������Ϣ -->
                <div class="info">�����ߣ�<%=Rs.Data("User")%> E-mail:<%=Rs.Data("Email")%> ��ҳ��<a href="<%=Rs.Data("HomePage")%>" target="_blank">���</a>��IP:<%=Rs.Data("IP")%>������<%=Rs.Data("CreateTime")%>
                 <span class="manage">
                 	<input type="checkbox" name="GroupID" value="<%=Rs.Data("ID")%>" />
                     ״̬: 
                     <%If Rs.Data("State") = 1 Then%>
                     	<a href="?id=<%=Rs.Data("ID")%>&action=setstate&state=nopass" title="ȡ�����">�����</a>
                     <%Else%>
                     	<span class="red"><a href="?id=<%=Rs.Data("ID")%>&action=setstate&state=pass" title="ͨ�����">δ���</a></span>
                     <%End If%>
                     <a href="?id=<%=Rs.Data("ID")%>&action=recomment">[�ظ�] </a>
                     <a href="?id=<%=Rs.Data("ID")%>&action=dodelete">[ɾ��]</a>
                </span>
                </div>
            </div> 
	<%
			Rs.Data.MoveNext
		Next
	%>	
			<div class="batCtrl">
  				<input type="button" onClick="selectAll(this.form,'GroupID')" value="ȫѡ" /> 
                <input type="button" onClick="selectOthers(this.form,'GroupID')" value="��ѡ" /> 
                &nbsp;&nbsp;
                ����������
                <select name="name" onChange="dobatch(this)" style="line-height:25px; padding:5px;">
                	<option value=""> ѡ����� </option>
                    <option value="pass"> ͨ����� </option>
                    <option value="nopass"> ȡ����� </option>
                    <option value="delete"> ����ɾ�� </option>
                </select>
            </div>
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
function batchPass(form){
	var id = getID(form);
	if(!id){return;}
	if(confirm('ȷ����ѡ��ѡ��ͨ����ˣ�')){	
		form.action  = '?action=setstate&state=pass&id=' + id;
		form.submit(); 
	}
}

//��������
function batchNoPass(form){
	var id = getID(form);
	if(!id){return;}
	if(confirm('ȷ����ѡ��ѡ��ȡ����ˣ�')){	
		form.action  = '?action=setstate&state=nopass&id=' + id;
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
'�Զ���ҳ���
Sub FuncForm(ByVal id)
	Dim objA: Set objA = New ClassGuestBook
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("�Բ�����༭�����Բ�����", "REFRESH")
	End If
%>
	<form action="?action=dorecomment" id="form1" name="form1" method="post">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <table class="form" style="border:1px #88C4FF solid;">
            <tr><th colspan="2">
				�ظ�����
            </th></tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
            	<td align="right" width="15%">���⣺</td>
            	<td><%=objA.Title%></td>
            </tr>
            <tr onmouseover="this.style.background='#FFFFFF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right">���ԣ�</td>
                <td ><%=objA.Content%></td>
            </tr>
            <tr>
            	<td align="right" width="15%">�ظ����ݣ�</td>
                <td>
                    <textarea name="fReComment" cols="50" rows="5"><%=objA.ReComment%></textarea> (250�ַ�����)
                </td>
            </tr>
            <tr>
            	<td align="right" width="15%">�������֣�</td>
                <td>
                    <input type="text" name="fReUser" value="<%=Session("AdminName")%>"/>
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
    	<a href="javascript:history.go(-1)" onclick="return confirm('ȷ�������ظ�����?')"><< == ���� << == </a>
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
