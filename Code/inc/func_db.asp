<%
'/*******************************************/
'/*	Purpose:	相关对数据库操作函数		*/
'/*	Author:		Foolin						*/
'/*	E-mail:		Foolin@126.com 				*/
'/*	Created on:	2009-7-20 23:37:37 			*/
'/*******************************************/


'获取数据库中MyTag的值
Function GetMyTag(ByVal tagName)
	GetMyTag = "1"
	Dim tagValue, strSql, objRs
	strSql = "SELECT Top 1 Code FROM  MyTags WHERE NAME = """ & tagName & """ ORDER BY ID DESC"
	'检查是否缓存
	If IsCache = 1 Then
		If ChkCache("MyTag_" & tagName) Then
			tagValue = GetCache("MyTag_" & tagName)
		Else
			Set objRs = DB(strSql, 1)
			If Not objRs.Eof Then
				tagValue = objRs("Code")
				Call SetCache("MyTag_" & tagName, tagValue)
			Else
				tagValue = Warn("未初始化标签{my:" & tagName & " /}")
			End If
			objRs.Close
			Set objRs = Nothing
		End If
	Else
		Set objRs = DB(strSql, 1)
		If Not objRs.Eof Then
			tagValue = objRs("Code")
		Else
			tagValue = Warn("未初始化标签{my:" & tagName & " /}")
		End If
		objRs.Close
		Set objRs = Nothing
	End If
	GetMyTag = tagValue
End Function

'获取数据库中栏目的名称
Function GetColName(ByVal ColId, ByVal strType)
	Dim objRs, strSql, strTempName
	ColId = CInt(ColId)
	If Not IsNumeric(ColId) Then GetColName = "": Exit Function
	If LCase(strType) = "picture" Or LCase(strType) = "pic" Then
		strSql = "SELECT Name FROM PicColumn WHERE ID = " & ColID
	Else
		strSql = "SELECT Name FROM ArtColumn WHERE ID = " & ColID
	End If
	Set objRs = DB(strSql, 1)
	If Not objRs.Eof Then
		strTempName = objRs("Name")
	Else
		strTempName = ""
	End If
	objRs.Close: Set objRs = Nothing
	GetColName = strTempName
End Function

'根据父栏目ID获取所有子栏目ID以及本身ID
Function GetColIds(ByVal FID, ByVal ColType)
	GetColIds = FID & GetSubColIds(FID, ColType)
End Function

'根据父栏目ID获取所有子栏目ID
Function GetSubColIds(ByVal FID, ByVal ColType)
	Dim Rs,strIds
	Select Case LCase(ColType)
		Case "picture","pic","image","img"
			Set Rs = DB("SELECT * FROM PicColumn WHERE ParentID IN (" & FID & ")", 1)
		Case Else
			Set Rs = DB("SELECT * FROM ArtColumn WHERE ParentID IN (" & FID & ")", 1)
	End Select
	'strIds = FID
	If Not Rs.Eof Then
		Do While Not Rs.Eof
			strIds = strIds &  "," & Rs("ID") & GetSubColIds(Rs("ID"), ColType) '递归子级分类
		Rs.Movenext:Loop
		If Rs.Eof Then
			Rs.Close: Set Rs = Nothing
			GetSubColIds = strIds
			Exit Function
		End If
	End If
	Rs.Close: Set Rs = Nothing
	GetSubColIds = strIds
End Function

'================================================================
'获取上一篇文章(图片)
'参数：	id -- 当前id， 
'		srcType -- 获取类型（Article|Picture)： 0 - 文章， 1 - 图片
'		getType -- 获取类型: 0 - id, 1 - Title, 2 - Url, 3 - Link
'================================================================
Function GetPreLink(ByVal id, ByVal srcType, ByVal getType)
	Dim Rs,strSql,strTemp, tempLinkUrl, titleLength: titleLength = 25		'标题长度
	If CInt(srcType) = 1  Then
		tempLinkUrl = "picture.asp"
		strSql = "Select top 1 ID,Title from Picture  where ID < " & id & " AND State = 1 order by ID desc"
	Else
		tempLinkUrl = "article.asp"
		strSql = "Select top 1 ID,Title from Article where ID < " & id & " AND State = 1  order by ID desc"
	End If
	Set Rs = DB(strSql,1)
	If Rs.Eof Then
		Select Case Int(getType)
		Case 0
			strTemp = "0"
		Case 1
			strTemp = "没有了"
		Case 2
			strTemp = "#"
		Case 3
			strTemp = "没有了"
		Case Else
			strTemp = "没有了"
		End Select
	Else
		Select Case Int(getType)
		Case 0
			strTemp = Rs("ID")
		Case 1
			strTemp = Rs("Title")
		Case 2
			strTemp = tempLinkUrl & "?id=" & Rs("ID")
		Case 3
			strTemp = "<a href=""" & tempLinkUrl & "?id=" & Rs("ID") & """ title=""" & Rs("Title") & """>" & CutStr(Rs("Title"), titleLength) & "</a>"
		Case Else
			strTemp = "<a href=""" & tempLinkUrl & "?id=" & Rs("ID") & """ title=""" & Rs("Title") & """>" & CutStr(Rs("Title"), titleLength) & "</a>"
		End Select
	End If
	Rs.Close: Set Rs = Nothing
	GetPreLink = strTemp
End Function

'================================================================
'获取下一篇文章（图片）
'参数：	id -- 当前id， 
'		srcType -- 获取类型（Article|Picture)： 0 - 文章， 1 - 图片
'		getType -- 获取类型: 0 - id, 1 - Title, 2 - Url, 3 - Link
'================================================================
Function GetNextLink(ByVal id, ByVal srcType, ByVal getType)
	Dim Rs,strSql,strTemp, tempLinkUrl, titleLength: titleLength = 25		'标题长度
	If CInt(srcType) = 1 Then
		tempLinkUrl = "picture.asp"
		strSql = "Select top 1 ID,Title from Picture where ID > " & ID & " AND State = 1  order by ID"
	Else
		tempLinkUrl = "article.asp"
		strSql = "Select top 1 ID,Title from Article where ID > " & ID & " AND State = 1  order by ID"
	End If
	Set Rs = DB(strSql,1)
	If Rs.Eof Then
		Select Case Int(getType)
		Case 0
			strTemp = "0"
		Case 1
			strTemp = "没有了"
		Case 2
			strTemp = "#"
		Case 3
			strTemp = "没有了"
		Case Else
			strTemp = "没有了"
		End Select
	Else
		Select Case Int(getType)
		Case 0
			strTemp = Rs("ID")
		Case 1
			strTemp = Rs("Title")
		Case 2
			strTemp = tempLinkUrl & "?id=" & Rs("ID")
		Case 3
			strTemp = "<a href=""" & tempLinkUrl & "?id=" & Rs("ID") & """ title=""" & Rs("Title") & """>" & CutStr(Rs("Title"), titleLength) & "</a>"
		Case Else
			strTemp = "<a href=""" & tempLinkUrl & "?id=" & Rs("ID") & """ title=""" & Rs("Title") & """>" & CutStr(Rs("Title"), titleLength) & "</a>"
		End Select
	End If
	Rs.Close: Set Rs = Nothing
	GetNextLink = strTemp
End Function

'---------------------------------------------------
'	函数：	GetColLink
'	功能：	获取栏目导航
'	参数：	id -- ColId, colType - 1 图片栏目， col - 0 文章栏目
'---------------------------------------------------
Function GetColLink(ByVal id, ByVal colType)
	Dim Rs, strSql, strLink, strUrl
	If Len(id) = 0 Or Not IsNumeric(id) Then GetColLink = "": Exit Function
	If Cint(colType) = 1 Then
		strSql = "SELECT ID, Name, ParentID FROM PicColumn WHERE ID= " & id
		strUrl = "piclist.asp"
	Else
		strSql = "SELECT ID, Name, ParentID FROM ArtColumn WHERE ID= " & id
		strUrl = "artlist.asp"
	End If
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strLink = "<a href=""" & strUrl & "?id=" & Rs("ID") & """>" & Rs("Name") & "</a>"
		'如果存在父ID，则进行递归调用
		If Rs("ParentID") <> 0 Then
			strLink =  GetColLink(Rs("ParentID"), colType) & SitePathSplit & strLink
		End If
	Else
		strLink = "<a href=""" & strUrl & """>列表</a>"
	End If
	Rs.Close: Set Rs = Nothing
	GetColLink = strLink
End Function

'获取栏目的名称
'artlist.asp中Title调用
Function GetNameOfColumn(Byval colid, Byval colType)
	Dim strName, Rs, strSql, strType
	If LCase(colType) = "picture" Then
		strType = "图片"
		strSql = "SELECT Name FROM PicColumn WHERE ID= " & colid
	Else
		strType = "文章"
		strSql = "SELECT Name FROM ArtColumn WHERE ID= " & colid
	End If
	If Len(id) = 0 Or Not IsNumeric(id) Then GetNameOfColumn = strType & "列表": Exit Function
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strName = Rs("Name")
	Else
		strName = strType & "列表"
	End If
	Rs.Close: Set Rs = Nothing
	GetNameOfColumn = strName
End Function

'获取文章、图片的标题
'article.asp(picture.asp)中Title调用
Function GetTitleOfArtOrPic(Byval id, Byval colType)
	Dim strTitle, Rs, strSql
	If LCase(colType) = "picture" Then
		strSql = "SELECT Title FROM Picture WHERE ID= " & id
	Else
		strSql = "SELECT Title FROM Article WHERE ID= " & id
	End If
	If Len(id) = 0 Or Not IsNumeric(id) Then GetTitleOfArtOrPic = "": Exit Function
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strTitle = Rs("Title")
	Else
		strTitle = "列表"
	End If
	Rs.Close: Set Rs = Nothing
	GetTitleOfArtOrPic = strTitle
End Function

'获取自定义页面的标题
'diypage.asp中Title调用
Function GetTitleOfDiypage(Byval param)
	Dim strTitle, Rs, strSql
	If Len(param) = 0 Then GetTitleOfDiypage= "自定义页面": Exit Function
	If IsNumeric(param) Then
		strSql = "SELECT Title FROM DiyPage WHERE ID = " & param
	Else
		strSql = "SELECT Title FROM DiyPage WHERE PageName = '" & param & "'"
	End If
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strTitle = Rs("Title")
	Else
		strTitle = "自定义页面"
	End If
	Rs.Close: Set Rs = Nothing
	GetTitleOfDiypage = strTitle
End Function
%>
