<!--#include file="inc/admin.include.asp"-->
<%
'===========================================
'File Name��	tags.asp
'Purpose��		��ȡ��ǩ�����ļ�
'Auhtor: 		Foolin
'Create on:		2009-9-30
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'===========================================
Call ChkLogin()
If Request("action") = "create" Then
	 Dim mode: mode = Request("ListMode")
	 Dim strCode: strCode = ""
	 Select Case LCase(mode)
		Case "default"
			strCode = DefaultCode()
		Case "table"
			strCode = TableCode()
		Case "sql"
			strCode = SQLCode()
		Case Else
	 End Select
End If

'Ĭ��ģʽ
Function DefaultCode()
	Dim vName, vSrc, vColumn, vRow, vCol, vWidth, vClass, vOrder, vIspage, tempCode, vInnerCode
	vName = Request("ListName"): If Len(vName) = 0 Then vName = "my"
	If Len(Request("ListSrc")) > 0 Then vSrc = " src=" & chr(34) & Request("ListSrc") & chr(34)
	If Len(Request("ListColumn")) > 0 Then vColumn = " column=" & chr(34) & Request("ListColumn") & chr(34)
	If Len(Request("ListRow")) > 0 Then vRow = " row=" & chr(34) & Request("ListRow") & chr(34)
	If Len(Request("ListCol")) > 0 Then vCol = " col=" & chr(34) & Request("ListCol") & chr(34)
	If Len(Request("ListWidth")) > 0 Then vWidth = " width=" & chr(34) & Request("ListWidth") & chr(34)
	If Len(Request("ListClass")) > 0 Then vClass = " class=" & chr(34) & Request("ListClass") & chr(34)
	If Len(Request("ListOrder")) > 0 Then vOrder = " order=" & chr(34) & Request("ListOrder") & chr(34)
	If LCase(Request("ListIspage")) = "true" Then vIspage = " ispage=" & chr(34) & Request("ListIspage") & chr(34)
	If LCase(Request("ListSrc")) = "picture" Then
		vInnerCode = PicTags(vName)
	Else
		vInnerCode = ArtTags(vName)
	End If
	If Cint(Request("ListCol")) > 1 Then
		tempCode = "{list:" & vName & vSrc & vColumn & vRow & vCol & vWidth & vClass & vOrder & vIspage & "}<br />"
	Else
		tempCode = "{list:" & vName & vSrc & vColumn & vRow & vCol & vOrder & vIspage & "}<br />"
	End If
	tempCode = tempCode & vInnerCode & "<br />"
	tempCode = tempCode & "{/list:" & vName &"}"
	DefaultCode = tempCode
End Function

'��ϱ�ǩģʽ
Function TableCode()
	Dim vName, vTable, vField, vWhere, vOrder, vRow, vCol, vWidth, vClass, vIspage, tempCode, vInnerCode
	vName = Request("ListName"): If Len(vName) = 0 Then vName = "my"
	If Len(Trim(Request("ListTable"))) = 0 Then vTable = "article" Else vTable = Request("ListTable")
	vTable = " table=" & chr(34) & vTable & chr(34)
	If Len(Request("ListField")) > 0 Then vField = " field=" & chr(34) & Request("ListField") & chr(34)
	If Len(Request("ListWhere")) > 0 Then vWhere = " where=" & chr(34) & Request("ListWhere") & chr(34)
	If Len(Request("ListOrder")) > 0 Then vOrder = " order=" & chr(34) & Request("ListOrder") & chr(34)
	If Len(Request("ListRow")) > 0 Then vRow = " row=" & chr(34) & Request("ListRow") & chr(34)
	If Len(Request("ListCol")) > 0 Then vCol = " col=" & chr(34) & Request("ListCol") & chr(34)
	If Len(Request("ListWidth")) > 0 Then vWidth = " width=" & chr(34) & Request("ListWidth") & chr(34)
	If Len(Request("ListClass")) > 0 Then vClass = " class=" & chr(34) & Request("ListClass") & chr(34)
	If LCase(Request("ListIspage")) = "true" Then vIspage = " ispage=" & chr(34) & Request("ListIspage") & chr(34)
	Select Case LCase(Request("ListTable"))
		Case "article"
			vInnerCode = ArtTags(vName)
		Case "picture"
			vInnerCode = PicTags(vName)
		Case "artcolumn", "piccolumn"
			vInnerCode = ColTags(vName)
		Case "guestbook"
			vInnerCode = GbookTags(vName)
		Case "mytags"
			vInnerCode = MyTags(vName)
		Case "diypage"
			vInnerCode = DiypageTags(vName)
		Case Else 
			vInnerCode = CommonTags(vName)
	End Select
	If Cint(Request("ListCol")) > 1 Then
		tempCode = "{list:" & vName & " mode=""table""" & vTable  & vField & vWhere & vOrder & vRow & vCol & vWidth & vClass & vIspage & "}<br />"
	Else
		tempCode = "{list:" & vName & " mode=""table""" & vTable  & vField & vWhere & vOrder & vRow & vCol & vIspage & "}<br />"
	End If
	tempCode = tempCode & vInnerCode & "<br />"
	tempCode = tempCode & "{/list:" & vName &"}"
	TableCode = tempCode
End Function

'SQL��ǩģʽ
Function SQLCode()
	Dim vName, vTable, vSQL, vRow, vCol, vWidth, vClass, vIspage, tempCode, vInnerCode
	vName = Request("ListName"): If Len(vName) = 0 Then vName = "MyList"
	vSQL = Trim(Request("ListSQL"))
	If Len(vSQL) = 0 Then Response.Write("<font color='red'>SQL����Ϊ�գ�������SQL��䣡</font>"): Response.End(): Exit Function
	If UCase(Left(vSQL,6)) <> "SELECT" Then Response.Write("<font color='red'>�Ƿ�SQL</font>"): Response.End(): Exit Function
	vSQL = " sql=" & chr(34) & vSQL & chr(34)
	If Len(Request("ListRow")) > 0 Then vRow = " row=" & chr(34) & Request("ListRow") & chr(34)
	If Len(Request("ListCol")) > 0 Then vCol = " col=" & chr(34) & Request("ListCol") & chr(34)
	If Len(Request("ListWidth")) > 0 Then vWidth = " width=" & chr(34) & Request("ListWidth") & chr(34)
	If Len(Request("ListClass")) > 0 Then vClass = " class=" & chr(34) & Request("ListClass") & chr(34)
	If LCase(Request("ListIspage")) = "true" Then vIspage = " ispage=" & chr(34) & Request("ListIspage") & chr(34)
	'������ʽ��ȡ���ݿ��
	Dim Reg, Match, Matches
	Set Reg = New RegExp
	Reg.Ignorecase = True
	Reg.Global = True
	Reg.Pattern = "from\s\[?([a-z]*)\]?(?:\swhere)?"
	Set Matches = Reg.Execute(vSql)
	For Each Match In Matches
		vTable = Match.SubMatches(0)
	Next
	Set Reg = Nothing
	'Response.Write("<font color='red'>SQL:" & vSql & " TABLE:" & vTable & "��</font>"): Response.End() 
	If Len(Trim(vTable)) = 0 Then Response.Write("<font color='red'>" & vSql & "����</font>"): Response.End() 
	Select Case LCase(Trim(vTable))
		Case "article"
			vInnerCode = ArtTags(vName)
		Case "picture"
			vInnerCode = PicTags(vName)
		Case "artcolumn", "piccolumn"
			vInnerCode = ColTags(vName)
		Case "guestbook"
			vInnerCode = GbookTags(vName)
		Case "mytags"
			vInnerCode = MyTags(vName)
		Case "diypage"
			vInnerCode = DiypageTags(vName)
		Case "weblog","admin","uploadfile","comment"
			vInnerCode = CommonTags(vName)
		Case Else 
			vInnerCode = "<br /><br /><font color='red'>SQL���󣬲��������ݿ��[<font color='blue'>" & vTable & "</font>]��</font><br /><br />"
	End Select
	If Cint(Request("ListCol")) > 1 Then
		tempCode = "{list:" & vName & " mode=""sql""" & vSql & vRow & vCol & vWidth & vClass & vIspage & "}<br />"
	Else
		tempCode = "{list:" & vName & " mode=""sql""" & vSql & vRow & vCol & vIspage & "}<br />"
	End If
	tempCode = tempCode & vInnerCode & "<br />"
	tempCode = tempCode & "{/list:" & vName &"}"
	SQLCode = tempCode
End Function

'���µײ��ǩ
Function ArtTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i���ʱ����ţ��Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- ��¼�������Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID��ʶ�����Զ����� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":url] &lt;!-- �������URL���Ǳ����ֶ�) --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":title] &lt;!-- ���±��� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":content] &lt;!-- �������� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colid] &lt;!-- ������ĿID --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colurl] &lt;!-- ������ĿURL���Ǳ����ֶ�) --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colname] &lt;!-- ������Ŀ���ƣ��Ǳ����ֶ�) --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":author] &lt;!-- ���� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":source] &lt;!-- ��Դ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":hits] &lt;!-- ����� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":focuspic] &lt;!-- ����ͼƬ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":keywords] &lt;!-- �ؼ��� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":createtime] &lt;!-- ����ʱ�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":modifytime] &lt;!-- �޸�ʱ�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":istop] &lt;!-- �Ƿ��ö���1 - �ö��� 0 - ���ö� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":isfocuspic] &lt;!-- �Ƿ񽹵�ͼƬ��1 - �ǣ� 0 - �� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":state] &lt;!-- ����״̬��1 - �Ѿ���ˣ� 0 - δ��ˣ� -1 - ��ɾ�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;"
	ArtTags = strTemp
End Function



'ͼƬ�ײ��ǩ
Function PicTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i���ʱ����ţ��Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- ��¼�������Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID��ʶ�����Զ����� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":url] &lt;!-- ���ͼƬURL���Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":title] &lt;!-- ���� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":smallpicpath] &lt;!-- ͼƬ����ͼ·��,��ʱΪ�գ�����ʹ��picpath --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":picpath] &lt;!-- ͼƬ·�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":intro] &lt;!-- ͼƬ���� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colid] &lt;!-- ������ĿID --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colurl] &lt;!-- ������ĿURL���Ǳ����ֶ�) --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colname] &lt;!-- ������Ŀ���ƣ��Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":author] &lt;!-- ���� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":source] &lt;!-- ��Դ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":hits] &lt;!-- ����� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":createtime] &lt;!-- ����ʱ�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":istop] &lt;!-- �Ƿ��ö���1 - �ö��� 0 - ���ö� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":state] &lt;!-- ͼƬ״̬��1 - �Ѿ���ˣ� 0 - δ��ˣ� -1 - ��ɾ�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;"
	PicTags = strTemp
End Function

'���¡�ͼƬ��Ŀ�ײ��ǩ
Function ColTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i���ʱ����ţ��Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- ��¼�������Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID��ʶ�����Զ����� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":url] &lt;!-- �����ĿURL���Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":name] &lt;!-- ��Ŀ���� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":info] &lt;!-- ��Ŀ��Ϣ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":parentid] &lt;!-- ����ĿID��������Ϊ����Ŀ��Ϊ0 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":template] &lt;!-- ��Ŀģ�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;"
	ColTags = strTemp
End Function


'���Եײ��ǩ
Function GbookTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i���ʱ����ţ��Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- ��¼�������Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID��ʶ�����Զ����� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":title] &lt;!-- ���Ա��� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":content] &lt;!-- �������� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":user] &lt;!-- ���������� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":email] &lt;!-- �����ߵ��ʼ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":homepage] &lt;!-- ��������ҳ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":ip] &lt;!-- ������IP --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":createtime] &lt;!-- ����ʱ�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":recomment] &lt;!-- �ظ��������� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":reuser] &lt;!-- �ظ����������� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":retime] &lt;!-- �ظ�ʱ�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":state] &lt;!-- ״̬�� 0 - δ��ˣ� 1 - �Ѿ���� --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;"
	GbookTags = strTemp
End Function

'�Զ����ǩ�ײ��ǩ
Function MyTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i���ʱ����ţ��Ǳ����ֶΣ����Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- ��¼�������Ǳ����ֶΣ����Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":url] &lt;!-- ���URL���Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID��ʶ�����Զ����� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":name] &lt;!-- ��ǩ�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":info] &lt;!-- ��ǩ������Ϣ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":code] &lt;!-- ��ǩ�Ĵ��� --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;"
	MyTags = strTemp
End Function


'�Զ���ҳ���ײ��ǩ
Function DiypageTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i���ʱ����ţ��Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- ��¼�������Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID��ʶ�����Զ����� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":title] &lt;!-- ҳ����� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":pagename] &lt;!-- ��ҳ���ļ��� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":keywords] &lt;!-- ҳ��ؼ��� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":template] &lt;!-- ҳ��ģ�� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":code] &lt;!-- ҳ����� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":state] &lt;!-- ״̬�� 0 - ���أ� 1 - ��ʾ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":issystem] &lt;!-- �Ƿ���ϵͳ����ҳ�棺0 - �� 1 - �� --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;"
	DiypageTags = strTemp
End Function

'��ͬ�ײ��ǩ
Function CommonTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i���ʱ����ţ��Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- ��¼�������Ǳ����ֶΣ� --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":�ֶ���] &lt;!-- �ֶ��� --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- �ڲ�ѭ����ǩ --&gt;"
	CommonTags = strTemp
End Function

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ǩ - Powered by eekku.com</title>
<style type="text/css">
<!--
body{
	font-size:14px;
	font-family:"Times New Roman", Times, serif;
	line-height:150%;
}
#showTags {
	color:red;
}
.desc{ color:#999; padding-left:5px; font-size:13px;}
-->
</style>
</head>

<body onload="parent.window.document.getElementById('showTags').height=document.body.scrollHeight"> 

 <div id="showTags">
 	<%=strCode%>
    <%	
		If LCase(Trim(Request("ListIsPage"))) = "true" Then
			Response.Write("<br /><br />&nbsp; {tag:page /} &lt;!-- ��ҳ��ǩ��ispage=��true��ʱ�� --&gt;<br /><br />") 
		End If
	%>
 </div>

<script type="text/javascript" language="javascript">
<!--
//��ȡIDԪ��
function $(o){ return typeof(o)=="string" ? document.getElementById(o) : o;}
//��ǩ����
function HightLightTag(){
	var strList = $("showTags").innerHTML;//�б�����
	
	var listName;	//�б�����
	var regExp;		//������ʽ
	//��ɫ����{tag:name }�е�name
	regExp = /\{(.+?):([^\s\]\}]+)/ig;
	strList = strList.replace(regExp, "{$1:<font color='blue'>$2</font>");
	//��ɫ����[list:name]�е�name
	//regExp = /\[(.+?):([^\s\]]+)\]/ig;
	regExp = /\[(.+?):(.+?)\]/ig;
	strList = strList.replace(regExp, " [<font color='blue'>$1</font>:<font color='green'>$2</font>\]");
	//��ɫ��������������ɫ��������ֵ
	//alert(strList);
	//regExp = /\s(\S+)=\"(\S*)\"/ig;
	regExp = /\s(\S+)=\"(.+?)\"/ig;
	strList = strList.replace(regExp, " <font color='blue'>$1</font>=\"<font color='green'>$2</font>\"");
	//alert(strList);
	//��ɫ����IF�����ײ��ǩ
	regExp = /\{if:(.+?)\}(.+?)/ig;
	strList = strList.replace(regExp, "{if:$1}<font color='green'>$2");
	regExp = /\{else\}/ig;
	strList = strList.replace(regExp, "</font>{else}<font color='green'>");
	regExp = /\{\/if\}/ig;
	strList = strList.replace(regExp, "</font>{/if}");
	//��ɫ��������˵��
	regExp = /&lt;!--(.+?)--&gt;/ig;
	strList = strList.replace(regExp, " &nbsp;<span class='desc'>&lt;!--$1--&gt;</span>");
	//alert(strList);
	$("showTags").innerHTML = strList;
	
}
HightLightTag();
-->
</script>
</body>
</html>
