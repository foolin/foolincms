<!--#include file="inc/admin.include.asp"-->
<!--#include file="inc/lib/class_piccolumn.asp"-->
<%
'=========================================================
' File Name��	admin_picolumn.asp
' Purpose��		ͼƬ��Ŀ����
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-9 10:27:17
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Dim act: act = Request("action")
Dim id: id = Request("id")
If Len(id) = 0 Then id = 0
Dim page: page = Request("page")
Dim MainStatus, SubStatus: MainStatus = "����ͼƬ��Ŀ"

Call ChkLogin()	'����¼
Call Init()		'��ʼ��ҳ��

'��ʼ��ҳ��
Sub Init()

	Select Case LCase(act)
		Case "create"
			SubStatus = "������Ŀ"
			Call Main("create")
		Case "modify"
			SubStatus = "�޸���Ŀ"
			Call Main("modify")
		Case "list"
			SubStatus = "��Ŀ�б�"
			Call Main("list")
		Case "docreate"
			Call DoCreate()
		Case "domodify"
			Call DoModify()
		Case "dodelete"
			Call DoDelete()
		Case Else
			SubStatus = "��Ŀ�б�"
			Call Main("list")
	End Select
	Call ConnClose()
End Sub

'������Ŀ
Function DoCreate()
	Dim objC: Set objC = New ClassPicColumn
	If objC.SetValue Then
		If objC.Create Then
			Call WebLog("�����Ŀ[Name:"&objC.Name&"]�ɹ���", "SESSION")
			Call MsgAndGo("�����Ŀ[Name:"&objC.Name&"]�ɹ���", "BACK")
		Else
			Call MsgBox("����" & objC.LastError, "BACK")
		End If
	Else
		Call MsgBox("����" & objC.LastError, "BACK")
	End If
	Set objC = Nothing
End Function

'ɾ����Ŀ
Sub DoModify()
	Dim objC: Set objC = New ClassPicColumn
	objC.ID = id
	If objC.SetValue Then
		If objC.Modify Then
			Call WebLog("�޸���Ŀ[id:"& id &"]�ɹ���", "SESSION")
			Call MsgAndGo("�޸���Ŀ[id:"& id &"]�ɹ���", "admin_piccolumn.asp")
		Else
			Call MsgBox("����" & objC.LastError, "BACK")
		End If
	Else
		Call MsgBox("����" & objC.LastError, "BACK")
	End If
	Set objC = Nothing
End Sub

'ɾ����Ŀ
Sub DoDelete()
	Dim objC: Set objC = New ClassPicColumn
	objC.ID = id
	If objC.Delete Then
		Call WebLog("ɾ����Ŀ[id:"& id &"]�ɹ���", "SESSION")
		Call MsgAndGo("ɾ����Ŀ[id:"& id &"]�ɹ���", "admin_piccolumn.asp")
	Else
		Call MsgBox("����" & objC.LastError, "BACK")
	End If
	Set objC = Nothing
End Sub

'������
Sub Main(ByVal picType)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SITENAME%>��̨���� - ͼƬ���� - <%=SYS%></title>
<link href="css/common.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="inc/base.js"></script>
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
.list{ font-size:13px;}
-->
</style>
</head>

<body>
<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td colspan="2" id="topNav">
			<%Call TopNav("picture")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                 <li class="mTitle">--== ͼƬ���� ==--</li>
                 <li><a href="admin_picture.asp?action=create">�ϴ�ͼƬ</a></li>
                 <li><a href="admin_picture.asp?action=list">����ͼƬ</a></li>
                 <li><a href="admin_picture.asp?action=list&list=trash">ͼƬ����վ</a></li>
                 <li class="mTitle">--== ͼƬ��Ŀ ==--</li>
                 <li <%If Request("action") = "create" Then Echo("class=""on""")%>><a href="admin_piccolumn.asp?action=create">�����Ŀ</a></li>
                 <li <%If Request("action") <> "create" Then Echo("class=""on""")%>><a href="admin_piccolumn.asp">������Ŀ</a></li>
                </ul>
				<%Call SysInfo()%>
            </td>
            <td id="content" valign="top">
            	<div class="content">
                	<div class="status"> ����λ�ã�<a href="index.asp">������ҳ</a> �� <%=MainStatus%> �� <%=SubStatus%></div>
					<%
                        Select Case LCase(picType)
                            Case "create"
                                ColForm(0)
                            Case "modify"
                                ColForm(id)
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
'ͼƬ�б� mode - ģʽ
Sub List()
%>
	<table class="list">
    	<tr>
        	<th>ID</th>
            <th>����</th>
            <th>��Ŀ����</th>
            <th>ģ��</th>
            <th>����</th>
            <th>ɾ��</th>
        </tr>
	<%
		Dim strSql, Rs
			'strSql = "SELECT * FROM [PicColumn]"
		strSql = "SELECT * FROM [PicColumn] WHERE ParentID = 0 ORDER BY ID"
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
        	<td><%=Rs.Data("ID")%></td>
			<td>
            	<a href="?action=modify&id=<%=Rs.Data("ID")%>"><%=Rs.Data("Name")%></a>
            </td>
            <td><%=Rs.Data("Info")%></td>
            <td><%=Rs.Data("Template")%></td>
            <td>
            	<a href="?action=modify&id=<%=Rs.Data("ID")%>">�༭</a>
            </td>
            <td>
				<a href="?action=dodelete&id=<%=Rs.Data("ID")%>" onclick="return confirm('��ɾ������Ŀ�Լ�����Ŀ������ͼƬ��\n\nɾ�������ָܻ���ȷ��ɾ�������ݣ�')">ɾ��</a>
            </td>
        </tr>
	<%		Call SubColumnTR(Rs.Data("ID"),"&nbsp;&nbsp; |-") 'ѭ���Ӽ�����
			Rs.Data.MoveNext
		Next
	%>
    </table>
    <div class="page"><%=Rs.Page%></div>
<%
	Rs.Data.Close: Set Rs = Nothing
End Sub

'����Ŀ����
Function SubColumnTR(FID,StrDis)
	Dim Rs1
	Set Rs1 = DB("SELECT * FROM PicColumn WHERE ParentID = " & FID, 1)
	If Not Rs1.Eof Then
		Do While Not Rs1.Eof
%>
        <tr onMouseOver="this.style.background='#C8E3E2';" onMouseOut="this.style.background='#F0F8FF'" style="color:#666;">
        	<td><%=Rs1("ID")%></td>
			<td>
            	<a href="?action=modify&id=<%=Rs1("ID")%>"> <%=StrDis%> <%=Rs1("Name")%></a>
            </td>
            <td><%=Rs1("Info")%></td>
            <td><%=Rs1("Template")%></td>
            <td>
            	<a href="?action=modify&id=<%=Rs1("ID")%>">�༭</a>
            </td>
            <td>
				<a href="?action=dodelete&id=<%=Rs1("ID")%>" onclick="return confirm('��ɾ������Ŀ�Լ�����Ŀ������ͼƬ��\n\nɾ�������ָܻ���ȷ��ɾ�������ݣ�')">ɾ��</a>
            </td>
        </tr>
<%
			Call SubColumnTr(Trim(Rs1("ID")),"&nbsp;&nbsp; " & Strdis) '�ݹ��Ӽ�����
		Rs1.Movenext:Loop
		If Rs1.Eof Then
			Rs1.Close: Set Rs1 = Nothing
			Exit Function
		End If
	End If
	Rs1.Close: Set Rs1 = Nothing
End Function
%>

<%
'ͼƬ��
Sub ColForm(ByVal id)
	Dim objA: Set objA = New ClassPicColumn
	If Cint(id) > 0 Then
		objA.ID = id
		If objA.LetValue = False Then Call MsgBox("�Բ�����༭����Ŀ������", "BACK")
	End If
%>
	<form action="?action=do<%If id > 0 Then Echo("modify") Else Echo("create")%>" id="form1" name="form1" method="post">
    	<input type="hidden" name="id" value="<%=objA.ID%>"/>
        <table class="form" style="border:1px #88C4FF solid;">
            <tr><th colspan="2">
				<%If id > 0 Then Echo("�༭") Else Echo("���")%>��Ŀ
            </th></tr>
            <tr>
            	<td align="right" width="15%">���ƣ�</td>
            	<td><input type="text" name="fName" value="<%=objA.Name%>" style="width:450px;"/> <span class="red">* ����</span></td>
            </tr>
            <tr>
                <td align="right">����Ŀ��</td>
                <td>
                	<select name="fParentID">
                        	<option value="0"> ��ѡ����Ŀ </option>         
                    	<%If objA.ParentID> 0 Then%>
                        	<option value="0"> ��һ����Ŀ </option>
                    		<option value="<%=objA.ID%>" selected="selected"> => <%=GetColName(objA.ParentID, "picture")%> <= </option>
                        <%Else%>
                        	 <option value="0"  selected="selected"> => ��һ����Ŀ <= </option>
                        <%End If%>
                    	<%Call MainColumn()%>
                    </select>����ѡ������Ϊ����Ŀ��
                </td>
            </tr>
            <tr>
            	<td align="right">��Ŀ���ܣ�</td>
                <td><textarea name="fInfo" cols="62" rows="3"><%=objA.Info%></textarea></td>
            </tr>
            <tr>
            	<td align="right" width="15%">ģ��·����</td>
            	<td><input type="text" name="fTemplate" value="<%=objA.Template%>" style="width:450px;"/> <span  style="color:gray;">(�����ݲ�֧��)</span></td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <input type="submit" class="btn" value="�ύ" />
                    <input type="reset" class="btn" value="����" />
                </td>
            </tr>
        </table>
    </div>
    <div class="page">
    	<a href="javascript:history.go(-1)" onclick="return confirm('ȷ�������޸�?')"><< == ���� << == </a>
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
<%
	Set objA = Nothing
End Sub

'��һ����Ŀ����
Function MainColumn()
	Dim Rs
	Set Rs = DB( "SELECT * FROM PicColumn WHERE ParentID = 0", 1)
	If Not Rs.Eof Then
		Do While Not Rs.Eof
			If Rs("ID") <> Cint(Request("id")) Then
				Echo("<option value=""" & Rs("ID") & """>" & Rs("Name") & "</option>" & Chr(10) & Chr(9) & Chr(9))
			End If
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
	Set Rs1 = DB("SELECT * FROM PicColumn WHERE ParentID = " & FID, 1)
	If Not Rs1.Eof Then
		Do While Not Rs1.Eof
			If Rs1("ID") <> Cint(Request("id")) Then
				Echo("<option value=""" & Rs1("ID") & """>" & StrDis & Rs1("Name") & "</option>" & Chr(10) & Chr(9))
			End If
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
