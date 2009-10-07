<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_admin.asp"-->
<%
'=========================================================
' File Name��	admin_mytag.asp
' Purpose��		�Ŷӹ���
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-3 16:40:50
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "<a href='?'>�����Ŷ�</a>"

Call ChkLogin()	'����¼
Call ChkPower("admin","all") '���Ȩ��
Call Init()		'��ʼ��ҳ��

'��ʼ��ҳ��
Sub Init()

	Select Case LCase(act)
		Case "create"
			SubStatus = "��������Ա"
			Call Main("create")
		Case "modify"
			SubStatus = "�޸Ĺ���Ա����"
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
			SubStatus = "�Ŷӳ�Ա�б�"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'�����û�
Function DoFreeze()
	Dim st: st = Request("state")
	If Len(id) = 0 Or Not IsNumeric(id) Then Call MsgBox("id��������", "REFRESH")
	Dim objA: Set objA = New ClassAdmin: objA.ID = id
	Select Case LCase(st)
		Case "freeze"
			If objA.Freeze Then
				Call WebLog("�����û�["& id &"]�ɹ���", "SESSION")
				Call MsgAndGo("�����û�["& id &"]�ɹ�", "REFRESH")
			Else
				Call MsgBox("����" & objA.LastError, "BACK")
			End If
		Case "unfreeze"
			If objA.Unfreeze Then
				Call WebLog("�ⶳ�û�["& id &"]�ɹ���", "SESSION")
				Call MsgAndGo("�ⶳ�û�["& id &"]�ɹ�", "REFRESH")
			Else
				Call MsgBox("����" & objA.LastError, "BACK")
			End If
		Case Else
			Call MsgBox("��������", "BACK")
	End Select
End Function

'�����Ŷ�
Function DoCreate()
	If Len(Request("fUsername")) = 0 Then Call MsgBox("�û�������Ϊ��!", "BACK")
	If Len(Request("fNickname")) = 0 Then Call MsgBox("�ǳƲ���Ϊ��!", "BACK")
	If Len(Request("fPassword")) < 6 Then Call MsgBox("���벻������6λ!", "BACK")
	If Request("fPassword")<>Request("fRePassword") Then Call MsgBox("���벻һ��!", "BACK")
	Dim objA: Set objA = New ClassAdmin
	If objA.SetValue = False Then
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	If objA.Create Then
		Call WebLog("��������Ա["&objA.Username&"]�ɹ���", "SESSION")
		Call MsgAndGo("��������Ա["&objA.Username&"]�ɹ���", "BACK")
	Else
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Function

'ɾ���Ŷ�
Sub DoModify()
	Dim objA: Set objA = New ClassAdmin
	objA.ID = id
	If Len(Request("fPassword"))>0 Then
		If Request("fPassword")<>Request("fRePassword") Then Call MsgBox("���벻һ��!", "BACK")
		If objA.SetValue = False Then
			Call MsgBox("����" & objA.LastError, "BACK")
		End If
		If objA.ModifyPsw Then
			Call WebLog("�޸Ĺ���Ա["& objA.Username &"]�ɹ���", "SESSION")
			Call MsgAndGo("�޸Ĺ���Ա["& objA.Username &"]�ɹ���", "admin_user.asp")
		Else
			Call MsgBox("����" & objA.LastError, "BACK")
		End If
	Else
		If objA.SetValue = False Then
			Call MsgBox("����" & objA.LastError, "BACK")
		End If
		If objA.ModifyInfo Then
			Call WebLog("�޸Ĺ���Ա["& objA.Username &"]��Ϣ�ɹ�����û�޸����룡", "SESSION")
			Call MsgAndGo("���޸Ĺ���Ա["& objA.Username &"]��Ϣ�ɹ�����û�޸����룡", "admin_user.asp")
		Else
			Call MsgBox("����" & objA.LastError, "BACK")
		End If
	End If
	Set objA = Nothing
End Sub

'ɾ���Ŷ�
Sub DoDelete()
	Dim objA: Set objA = New ClassAdmin
	objA.ID = id
	If objA.Delete Then
		Call WebLog("ɾ������Ա["& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("ɾ������Ա["& id &"]�ɹ���", "admin_user.asp")
	Else
		Call MsgBox("����" & objA.LastError, "BACK")
	End If
	Set objA = Nothing
End Sub

'����Ա��ӦȨ��
Function GetLevel(Byval iLevel)
	Dim tLevel
	Select Case iLevel
		Case -1
			tLevel = "<font color='blue'>�����û�</font>"
		Case 0
			tLevel = "��ͨ����Ա"
		Case 1
			tLevel = "�м�����Ա"
		Case 2
			tLevel = "�߼�����Ա"
		Case 3
			tLevel = "��������Ա"
		Case Else
			tLevel = iLevel
	End Select
	GetLevel = tLevel
End Function

'������
Sub Main(ByVal strType)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>��̨���� - �Ŷӹ��� - <%=SYS%></title>
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
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== �Ŷӹ��� ==--</li>
                 <li <%If Request("action") = "create" Then Response.Write("class=""on""")%>><a href="?action=create">��ӹ���Ա</a></li>
                 <li <%If Request("action") <> "create" And Request("list") <> "trash"  Then Response.Write("class=""on""")%>><a href="?action=list">����Ա�б�</a></li>
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
'�Ŷ��б� mode - ģʽ
Sub List()
%>
	<form name="form2" action="" method="post">
	<table class="list">
    	<tr>
        	<th><input type="checkbox" name="GroupID" value="" onClick="Checked(this.form,'GroupID',this)"/></th>
        	<th>ID</th>
            <th>�û���</th>
            <th>�ǳ�</th>
            <th>�ȼ�</th>
            <th>��½����</th>
            <th>��¼ʱ��</th>
            <th>��¼IP</th>
            <th>����</th>
            <th>ɾ��</th>
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
            		<a href="?action=dofreeze&state=unfreeze&id=<%=Rs.Data("ID")%>" onclick="return confirm('�ⶳ�û�����ɳ�������Ա��\n\nȷ���ⶳ���û���')">�ⶳ</a>
               <%Else%>
               		<a href="?action=modify&id=<%=Rs.Data("ID")%>">�༭</a> | 
                    <a href="?action=dofreeze&state=freeze&id=<%=Rs.Data("ID")%>" onclick="return confirm('�����û������ܵ�¼��̨����\n\nȷ��������û���')">����</a>
               <%End If%>
            </td>
            <td>
				<a href="?action=dodelete&id=<%=Rs.Data("ID")%>" onclick="return confirm('ɾ�������ûָ���\n\nȷ��������ɾ�����û���')">ɾ��</a>
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
                    <option value="freeze"> �����û� </option>
                    <option value="unfreeze"> �ⶳ�û� </option>
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
function BatchFreeze(form, isFreeze){
	var id = GetID(form);
	if(id){
		if(isFreeze){
			if (!confirm('�����û������ܽ����̨��\n\n�Ƿ��ѡ���û����ж��᣿')) return;
			form.action = '?action=dofreeze&state=freeze&id=' + id;
			form.submit();  
		}
		else{
			if(!confirm('�ⶳ���û�����ɳ�������Ա��\n\n�Ƿ��ѡ���û����нⶳ��')) return;
			form.action  = '?action=dofreeze&state=unfreeze&id=' + id;
			form.submit();  
		}
	}
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
'�Ŷӱ�
Sub FuncForm(ByVal id)
	Dim objA: Set objA = New ClassAdmin
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("�Բ�����༭�Ĺ���Ա������", "BACK")
	End If
%>
	<form action="?action=do<%If id > 0 Then Echo("modify") Else Echo("create")%>" id="form1" name="form1" method="post">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <table class="form" style="border:1px #88C4FF solid;">
            <tr><th colspan="2">
				<%If id > 0 Then Echo("�༭") Else Echo("���")%>����Ա
            </th></tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
            	<td align="right" width="15%">����Ա��</td>
            	<td><input type="text" name="fUsername" value="<%=objA.Username%>" <%If id > 0 Then Echo("readonly=""readonly""")%> style="width:450px;"/> <span class="red">* ���ֻ��Ӣ�ĺ��»��ߣ�����ʹ�����ģ�</span></td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
            	<td align="right" width="15%">�ǳƣ�</td>
            	<td><input type="text" name="fNickname" value="<%=objA.Nickname%>" style="width:450px;"/> <span class="red">* �������ʹ�����ģ�</span></td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right" width="15%">�����룺</td>
                <td><input type="password" name="fPassword" id="fPassword" value="" style="width:250px;"/> 
				<%If id > 0 Then%>
                	<span class="gray">������޸�����</span>
                <%Else%>
                	<span class="red"> * ����</span>
                <%End If%>
                </td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right" width="15%">�ظ������룺</td>
                <td><input type="password" name="fRePassword" id="fRePassword" value="" style="width:250px;"/> </td>
            </tr>
            <tr onmouseover="this.style.background='#51C7FF';" onmouseout="this.style.background='#F0F8FF'">
                <td align="right" width="15%">Ȩ��</td>
                <td>
                <select name="fLevel" style="line-height:25px; padding:5px;">
                	<option value="0"> ѡ��ȼ� </option>
                	<%If id > 0 Then%>
                    	<option value="<%=objA.Level%>" selected="selected"> <%=GetLevel(objA.Level)%> </option>
                    <%End If%>
                    <option value="0"> ��ͨ����Ա </option>
                    <option value="1"> �м�����Ա </option>
                    <option value="2"> �߼�����Ա </option>
                    <option value="3"> ��������Ա </option>
                    <option value="4"> �����û� </option>
                </select> <span class="red"> * ��ѡ</span>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <input type="submit" class="btn" value="�ύ" />
                    <input type="reset" class="btn" value="����" />
                </td>
            </tr>
        </table>
        <div style="border:dashed 1px #CCC; margin:10px 0px; padding:5px; line-height:22px;">
         <b>Ȩ��˵����</b><br />
         1����ͨ����ԱȨ�ޣ��������¡�ͼƬ�����ԡ�<br />
         2���м�����ԱȨ�ޣ��������¡�ͼƬ�����ԡ�������Ŀ��ͼƬ��Ŀ��<br />
         3���߼�����ԱȨ�ޣ��߼�����ԱȨ�� + ��ǩ����DIYҳ�����ϵͳ���á�������¼��<br />
         4����������ԱȨ�ޣ�����ȫ��Ȩ�ޣ����ǣ��߼�����ԱȨ�� + ģ����� + �Ŷӹ�����<br />
         5�������û���û�й���Ȩ�ޣ����Ҳ��ܵ�¼��ϵͳ��̨����
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
    	<a href="javascript:history.go(-1)" onclick="return confirm('ȷ�������༭����Ա?')"><< == ���� << == </a>
    </div>
<%
	Set objA = Nothing
End Sub
%>
