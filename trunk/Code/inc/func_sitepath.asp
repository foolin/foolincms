<%
'---------------------------------------------------
'	������	ToPath
'	���ܣ�	���������������Ӵ�������
'	������	strLinkA - ��һ�����ӣ� strLinkB - �ڶ�������
'---------------------------------------------------
Function ToPath(ByVal strLinkA, ByVal strLinkB)
	ToPath = strLinkA & SitePathSplit & strLinkB
End Function

'---------------------------------------------------
'	������	ToLink
'	���ܣ�	������ת��������
'	������	strName - ���ƣ� strUrl - ����
'---------------------------------------------------
Function ToLink(ByVal strName, ByVal strUrl)
	ToLink = "<a href=""" & strUrl & """>" & strName & "</a>"
End Function



'---------------------------------------------------
'	������	IndexPath
'	���ܣ�	��ҳ���ӵ���
'	������	��
'---------------------------------------------------
Function IndexPath()
	IndexPath = ToLink("��ҳ", "index.asp")
End Function

'---------------------------------------------------
'	������	ColPath
'	���ܣ�	��Ŀ���ӵ���
'	������	id - ��ĿID, srcType - 1 ͼƬ�� 0 - ����
'---------------------------------------------------
Function ColPath(ByVal id, ByVal colType)
	ColPath = ToPath(IndexPath, GetColLink(id, colType))
End Function

'---------------------------------------------------
'	������	ArtPath
'	���ܣ�	��Ŀ���ӵ���
'	������	id - ����ID
'---------------------------------------------------
Function ArtPath(ByVal id)
	Dim Rs, strSql, strPath
	If Len(id) = 0 Or Not IsNumeric(id) Then PicPath = "": Exit Function
	strSql = "SELECT ColID FROM Article WHERE ID = " & id
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, GetColLink(Rs("ColID"), 0))
		strPath = ToPath(strPath, "�������")
	Else
		strPath = "�������"
	End If
	Rs.Close: Set Rs = Nothing
	ArtPath = strPath
End Function

'---------------------------------------------------
'	������	PicPath
'	���ܣ�	��Ŀ���ӵ���
'	������	id - ͼƬid
'---------------------------------------------------
Function PicPath(ByVal id)
	Dim Rs, strSql, strPath
	If Len(id) = 0 Or Not IsNumeric(id) Then PicPath = "": Exit Function
	strSql = "SELECT ColID FROM Picture WHERE ID = " & id
	Set Rs = DB(strSql, 1)
	If Not Rs.Eof Then
		strPath = ToPath(IndexPath, GetColLink(Rs("ColID"), 1))
		strPath = ToPath(strPath, "ͼƬ���")
	Else
		strPath = "ͼƬ���"
	End If
	Rs.Close: Set Rs = Nothing
	PicPath = strPath
End Function

'---------------------------------------------------
'	������	DiyPagePath
'	���ܣ�	��Ŀ���ӵ���
'	������	id - ҳ��id
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
		strPath = "�Զ���ҳ��"
	End If
	Rs.Close: Set Rs = Nothing
	DiyPagePath = strPath
End Function

'---------------------------------------------------
'	������	GuestbookPath
'	���ܣ�	��Ŀ���ӵ���
'---------------------------------------------------
Function GuestbookPath()
	GuestbookPath = ToPath(IndexPath, ToLink("���Բ�", "guestbook.asp"))
End Function
%>