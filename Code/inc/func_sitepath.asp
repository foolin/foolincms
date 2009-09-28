<%
'---------------------------------------------------
'	函数：	ToPath
'	功能：	辅助，将两个连接串接起来
'	参数：	strLinkA - 第一个连接， strLinkB - 第二个链接
'---------------------------------------------------
Function ToPath(ByVal strLinkA, ByVal strLinkB)
	ToPath = strLinkA & SitePathSplit & strLinkB
End Function

'---------------------------------------------------
'	函数：	ToLink
'	功能：	辅助，转换成连接
'	参数：	strName - 名称， strUrl - 连接
'---------------------------------------------------
Function ToLink(ByVal strName, ByVal strUrl)
	ToLink = "<a href=""" & strUrl & """>" & strName & "</a>"
End Function



'---------------------------------------------------
'	函数：	IndexPath
'	功能：	首页链接导航
'	参数：	无
'---------------------------------------------------
Function IndexPath()
	IndexPath = ToLink("首页", "index.asp")
End Function

'---------------------------------------------------
'	函数：	ColPath
'	功能：	栏目链接导航
'	参数：	id - 栏目ID, srcType - 1 图片， 0 - 文章
'---------------------------------------------------
Function ColPath(ByVal id, ByVal colType)
	ColPath = ToPath(IndexPath, GetColLink(id, colType))
End Function

'---------------------------------------------------
'	函数：	ArtPath
'	功能：	栏目链接导航
'	参数：	id - 文章ID
'---------------------------------------------------
Function ArtPath(ByVal id)
	Dim Rs, strSql, strPath
	If Len(id) = 0 Or Not IsNumeric(id) Then PicPath = "": Exit Function
	strSql = "SELECT ColID FROM Article WHERE ID = " & id
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, GetColLink(Rs("ColID"), 0))
		strPath = ToPath(strPath, "文章浏览")
	Else
		strPath = "文章浏览"
	End If
	Rs.Close: Set Rs = Nothing
	ArtPath = strPath
End Function

'---------------------------------------------------
'	函数：	PicPath
'	功能：	栏目链接导航
'	参数：	id - 图片id
'---------------------------------------------------
Function PicPath(ByVal id)
	Dim Rs, strSql, strPath
	If Len(id) = 0 Or Not IsNumeric(id) Then PicPath = "": Exit Function
	strSql = "SELECT ColID FROM Picture WHERE ID = " & id
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, GetColLink(Rs("ColID"), 1))
		strPath = ToPath(strPath, "图片浏览")
	Else
		strPath = "图片浏览"
	End If
	Rs.Close: Set Rs = Nothing
	PicPath = strPath
End Function

'---------------------------------------------------
'	函数：	DiyPagePath
'	功能：	栏目链接导航
'	参数：	id - 页面id
'---------------------------------------------------
Function DiyPagePath(ByVal param)
	Dim Rs, strSql, strPath
	If Len(param) = 0 Then DiyPagePath = "": Exit Function
	If IsNumeric(param) Then
		strSql = "SELECT Title FROM DiyPage WHERE ID = " & param
	Else
		strSql = "SELECT Title FROM DiyPage WHERE PageName = '" & param & "'"
	End If
	
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, Rs("Title"))
	Else
		strPath = "自定义页面"
	End If
	Rs.Close: Set Rs = Nothing
	DiyPagePath = strPath
End Function

'---------------------------------------------------
'	函数：	GuestbookPath
'	功能：	栏目链接导航
'---------------------------------------------------
Function GuestbookPath()
	GuestbookPath = ToPath(IndexPath, ToLink("留言簿", "guestbook.asp"))
End Function
%>