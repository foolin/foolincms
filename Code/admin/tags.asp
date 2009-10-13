<!--#include file="inc/admin.include.asp"-->
<%
'===========================================
'File Name：	tags.asp
'Purpose：		获取标签帮助文件
'Auhtor: 		Foolin
'Create on:		2009-9-30
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
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

'默认模式
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

'组合标签模式
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

'SQL标签模式
Function SQLCode()
	Dim vName, vTable, vSQL, vRow, vCol, vWidth, vClass, vIspage, tempCode, vInnerCode
	vName = Request("ListName"): If Len(vName) = 0 Then vName = "MyList"
	vSQL = Trim(Request("ListSQL"))
	If Len(vSQL) = 0 Then Response.Write("<font color='red'>SQL不能为空，请输入SQL语句！</font>"): Response.End(): Exit Function
	If UCase(Left(vSQL,6)) <> "SELECT" Then Response.Write("<font color='red'>非法SQL</font>"): Response.End(): Exit Function
	vSQL = " sql=" & chr(34) & vSQL & chr(34)
	If Len(Request("ListRow")) > 0 Then vRow = " row=" & chr(34) & Request("ListRow") & chr(34)
	If Len(Request("ListCol")) > 0 Then vCol = " col=" & chr(34) & Request("ListCol") & chr(34)
	If Len(Request("ListWidth")) > 0 Then vWidth = " width=" & chr(34) & Request("ListWidth") & chr(34)
	If Len(Request("ListClass")) > 0 Then vClass = " class=" & chr(34) & Request("ListClass") & chr(34)
	If LCase(Request("ListIspage")) = "true" Then vIspage = " ispage=" & chr(34) & Request("ListIspage") & chr(34)
	'正则表达式获取数据库表
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
	'Response.Write("<font color='red'>SQL:" & vSql & " TABLE:" & vTable & "！</font>"): Response.End() 
	If Len(Trim(vTable)) = 0 Then Response.Write("<font color='red'>" & vSql & "出错！</font>"): Response.End() 
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
			vInnerCode = "<br /><br /><font color='red'>SQL错误，不存在数据库表[<font color='blue'>" & vTable & "</font>]！</font><br /><br />"
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

'文章底层标签
Function ArtTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- 内层循环标签 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i输出时的序号（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- 记录总数（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID标识符（自动排序） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":url] &lt;!-- 浏览文章URL（非表中字段) --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":title] &lt;!-- 文章标题 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":content] &lt;!-- 文章内容 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colid] &lt;!-- 所属栏目ID --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colurl] &lt;!-- 所属栏目URL（非表中字段) --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colname] &lt;!-- 所属栏目名称（非表中字段) --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":author] &lt;!-- 作者 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":source] &lt;!-- 来源 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":hits] &lt;!-- 点击率 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":focuspic] &lt;!-- 焦点图片 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":keywords] &lt;!-- 关键词 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":createtime] &lt;!-- 创建时间 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":modifytime] &lt;!-- 修改时间 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":istop] &lt;!-- 是否置顶：1 - 置顶， 0 - 不置顶 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":isfocuspic] &lt;!-- 是否焦点图片：1 - 是， 0 - 否 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":state] &lt;!-- 文章状态：1 - 已经审核， 0 - 未审核， -1 - 已删除 --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- 内层循环标签 --&gt;"
	ArtTags = strTemp
End Function



'图片底层标签
Function PicTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- 内层循环标签 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i输出时的序号（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- 记录总数（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID标识符（自动排序） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":url] &lt;!-- 浏览图片URL（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":title] &lt;!-- 标题 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":smallpicpath] &lt;!-- 图片缩略图路径,有时为空，建议使用picpath --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":picpath] &lt;!-- 图片路径 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":intro] &lt;!-- 图片介绍 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colid] &lt;!-- 所属栏目ID --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colurl] &lt;!-- 所属栏目URL（非表中字段) --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":colname] &lt;!-- 所属栏目名称（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":author] &lt;!-- 作者 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":source] &lt;!-- 来源 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":hits] &lt;!-- 点击率 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":createtime] &lt;!-- 创建时间 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":istop] &lt;!-- 是否置顶：1 - 置顶， 0 - 不置顶 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":state] &lt;!-- 图片状态：1 - 已经审核， 0 - 未审核， -1 - 已删除 --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- 内层循环标签 --&gt;"
	PicTags = strTemp
End Function

'文章、图片栏目底层标签
Function ColTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- 内层循环标签 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i输出时的序号（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- 记录总数（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID标识符（自动排序） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":url] &lt;!-- 浏览栏目URL（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":name] &lt;!-- 栏目名称 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":info] &lt;!-- 栏目信息 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":parentid] &lt;!-- 父栏目ID，若本身为父栏目则为0 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":template] &lt;!-- 栏目模板 --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- 内层循环标签 --&gt;"
	ColTags = strTemp
End Function


'留言底层标签
Function GbookTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- 内层循环标签 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i输出时的序号（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- 记录总数（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID标识符（自动排序） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":title] &lt;!-- 留言标题 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":content] &lt;!-- 留言内容 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":user] &lt;!-- 留言者名字 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":email] &lt;!-- 留言者的邮件 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":homepage] &lt;!-- 留言者主页 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":ip] &lt;!-- 留言者IP --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":createtime] &lt;!-- 留言时间 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":recomment] &lt;!-- 回复留言内容 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":reuser] &lt;!-- 回复留言者名字 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":retime] &lt;!-- 回复时间 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":state] &lt;!-- 状态： 0 - 未审核， 1 - 已经审核 --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- 内层循环标签 --&gt;"
	GbookTags = strTemp
End Function

'自定义标签底层标签
Function MyTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- 内层循环标签 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i输出时的序号（非表中字段）（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- 记录总数（非表中字段）（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":url] &lt;!-- 浏览URL（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID标识符（自动排序） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":name] &lt;!-- 标签名 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":info] &lt;!-- 标签描述信息 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":code] &lt;!-- 标签的代码 --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- 内层循环标签 --&gt;"
	MyTags = strTemp
End Function


'自定义页面表底层标签
Function DiypageTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- 内层循环标签 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i输出时的序号（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- 记录总数（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":id] &lt;!-- ID标识符（自动排序） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":title] &lt;!-- 页面标题 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":pagename] &lt;!-- 该页面文件名 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":keywords] &lt;!-- 页面关键词 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":template] &lt;!-- 页面模板 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":code] &lt;!-- 页面代码 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":state] &lt;!-- 状态： 0 - 隐藏， 1 - 显示 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":issystem] &lt;!-- 是否是系统定义页面：0 - 否， 1 - 是 --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- 内层循环标签 --&gt;"
	DiypageTags = strTemp
End Function

'共同底层标签
Function CommonTags(Byval ListName)
	Dim strTemp
	strTemp = "&nbsp; &lt;!-- 内层循环标签 --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":i] &lt;!-- i输出时的序号（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":num] &lt;!-- 记录总数（非表中字段） --&gt;<br />"
	strTemp = strTemp & "&nbsp; [" & ListName & ":字段名] &lt;!-- 字段名 --&gt;<br />"
	strTemp = strTemp & "&nbsp; &lt;!-- 内层循环标签 --&gt;"
	CommonTags = strTemp
End Function

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>标签 - Powered by eekku.com</title>
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
			Response.Write("<br /><br />&nbsp; {tag:page /} &lt;!-- 分页标签，ispage=“true”时用 --&gt;<br /><br />") 
		End If
	%>
 </div>

<script type="text/javascript" language="javascript">
<!--
//获取ID元素
function $(o){ return typeof(o)=="string" ? document.getElementById(o) : o;}
//标签高亮
function HightLightTag(){
	var strList = $("showTags").innerHTML;//列表内容
	
	var listName;	//列表名称
	var regExp;		//正则表达式
	//绿色高亮{tag:name }中的name
	regExp = /\{(.+?):([^\s\]\}]+)/ig;
	strList = strList.replace(regExp, "{$1:<font color='blue'>$2</font>");
	//绿色高亮[list:name]中的name
	//regExp = /\[(.+?):([^\s\]]+)\]/ig;
	regExp = /\[(.+?):(.+?)\]/ig;
	strList = strList.replace(regExp, " [<font color='blue'>$1</font>:<font color='green'>$2</font>\]");
	//蓝色高亮属性名，绿色高亮属性值
	//alert(strList);
	//regExp = /\s(\S+)=\"(\S*)\"/ig;
	regExp = /\s(\S+)=\"(.+?)\"/ig;
	strList = strList.replace(regExp, " <font color='blue'>$1</font>=\"<font color='green'>$2</font>\"");
	//alert(strList);
	//绿色高亮IF条件底层标签
	regExp = /\{if:(.+?)\}(.+?)/ig;
	strList = strList.replace(regExp, "{if:$1}<font color='green'>$2");
	regExp = /\{else\}/ig;
	strList = strList.replace(regExp, "</font>{else}<font color='green'>");
	regExp = /\{\/if\}/ig;
	strList = strList.replace(regExp, "</font>{/if}");
	//灰色高亮解析说明
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
