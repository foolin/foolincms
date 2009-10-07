<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_mytag.asp"-->
<%
'=========================================================
' File Name��	admin_mytag.asp
' Purpose��		������־����
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-3 16:40:50
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "<a href='?'>���������־</a>"

Call ChkLogin()	'����¼

Call Init()		'��ʼ��ҳ��

'��ʼ��ҳ��
Sub Init()
	Select Case LCase(act)
		Case "doclear"
			Call ChkPower("weblog","delete") '���Ȩ��
			Call DoClear()
		Case "dodelete"
			Call ChkPower("weblog","delete") '���Ȩ��
			Call DoDelete()
		Case Else
			SubStatus = "������־�б�"
			Call Main()
	End Select
	Call ConnClose()
End Sub

'��ղ�����־
Sub DoClear()
	If ClearWebLog() = True Then
		'Call WebLog("���ȫ��������־�ɹ���", "SESSION")
		Call MsgAndGo("���ȫ��������־�ɹ���", "admin_weblog.asp")
	End If
End Sub

'ɾ��������־
Sub DoDelete()
	If DelWebLog(id) = True Then
		'Call WebLog("ɾ��������־[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("ɾ��������־[id:"& id &"]�ɹ���", "admin_weblog.asp")
	End If
End Sub

'������
Sub Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>��̨���� - ������־���� - <%=SYS%></title>
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
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== ������־���� ==--</li>
                 <li <%If Request("action") <> "create" And Request("list") <> "trash"  Then Response.Write("class=""on""")%>><a href="?action=list">���������־</a></li>
                 <li><a href="?action=doclear" onclick="return confirm('���ȫ����־�����ûָ����ݣ�\n\nȷ�����������ȫ����־��')">��ղ�����־</a></li>
                </ul>
				<%Call SysInfo()%>
            </td>
            <td id="content" valign="top">
            	<div class="content">
                	<div class="status"> ����λ�ã�<a href="index.asp">������ҳ</a> �� <%=MainStatus%> �� <%=SubStatus%> </div>
					
                    
	<form name="form2" action="" method="post">
	<table class="list">
    	<tr>
        	<th width="5%"><input type="checkbox" name="GroupID" value="" onClick="Checked(this.form,'GroupID',this)"/></th>
            <th width="30%">����</th>
            <th width="10%">�û�</th>
            <th width="12%">IP</th>
            <th width="20%">ҳ��</th>
            <th>ʱ��</th>
            <th width="8%">ɾ��</th>
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
                ������
                <input type="button" onClick="BatchDelete(this.form)" value="ɾ��ѡ����" />
                <input type="button" onClick="ClearAll(this.form)" value="���ȫ����־" />
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

//����ɾ��
function BatchDelete(form){
	var id = GetID(form);
	if(!id){return;}
	if(confirm('ɾ�������ָܻ���\n\n�Ƿ����ɾ��ѡ�еļ�¼��')){	
		form.action  = '?action=dodelete&id=' + id;
		form.submit(); 
	}
}
//����ɾ��
function ClearAll(form){
	if(confirm('���ȫ����־�����ָܻ���\n\n�Ƿ�������������־��¼��')){	
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
