<%
'=========================================================
' Class Name：	ClassPicture
' Purpose：		图片管理类
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-11 18:48:23
' Modify log:	
' Updated on: 	
' CopyRight (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Class ClassPicture

	'v前缀：Value，数据库字段的值（类成员）
	Private vID
	Private vColID
	Private vTitle
	Private vAuthor
	Private vSource
	Private vSmallPicPath
	Private vPicPath
	Private vIntro
	Private vIsTop
	Private vState
	Private vHits
	Private vCreateTime
	'm前缀：Menber，类成员
	Dim mLastError
	
	'ID
	Public Property Let ID(ByVal pID): vID = pID: End Property
	Public Property Get ID: ID = vID: End Property
	'ColID
	Public Property Let ColID(ByVal pColID): vColID = pColID: End Property
	Public Property Get ColID: ColID = vColID: End Property
	'Title
	Public Property Let Title(ByVal pTitle): vTitle = pTitle: End Property
	Public Property Get Title: Title = vTitle: End Property
	'Author
	Public Property Let Author(ByVal pAuthor): vAuthor = pAuthor: End Property
	Public Property Get Author: Author = vAuthor: End Property
	'Source
	Public Property Let Source(ByVal pSource): vSource = pSource: End Property
	Public Property Get Source: Source = vSource: End Property
	'SmallPicPath
	Public Property Let SmallPicPath(ByVal pSmallPicPath): vSmallPicPath = pSmallPicPath: End Property
	Public Property Get SmallPicPath: SmallPicPath = vSmallPicPath: End Property
	'PicPath
	Public Property Let PicPath(ByVal pPicPath): vPicPath = pPicPath: End Property
	Public Property Get PicPath: PicPath = vPicPath: End Property
	'Intro
	Public Property Let Intro(ByVal pIntro): vIntro = pIntro: End Property
	Public Property Get Intro: Intro = vIntro: End Property
	'IsTop
	Public Property Let IsTop(ByVal pIsTop): vIsTop = pIsTop: End Property
	Public Property Get IsTop: IsTop = vIsTop: End Property
	'State
	Public Property Let State(ByVal pState): vState = pState: End Property
	Public Property Get State: State = vState: End Property
	'Hits
	Public Property Let Hits(ByVal pHits): vHits = pHits: End Property
	Public Property Get Hits: Hits = vHits: End Property
	'CreateTime
	Public Property Let CreateTime(ByVal pCreateTime): vCreateTime = pCreateTime: End Property
	Public Property Get CreateTime: CreateTime = vCreateTime: End Property
	'LastError
	Public Property Let LastError(ByVal pLastError): mLastError = pLastError: End Property
	Public Property Get LastError: LastError = mLastError: End Property
	
	Private Sub Class_Initialize()
		Call ChkLogin()		'检查登录
		Call Initialize()	'初始化
	End Sub

	Private Sub Class_Terminate()
		Call Initialize()
	End Sub

	Public Function Initialize()
		vID = 0
		vColID = 0
		vTitle = ""
		vAuthor = ""
		vSource = "本站"
		vSmallPicPath = ""
		vPicPath = ""
		vIntro = "暂无介绍"
		vIsTop = 0
		vState = 0
		vHits = 0
		vCreateTime = Now()
	End Function
	
	'--------------------------------------------------------------
	' Function name：	SetValue
	' Description: 		从表单获取数据并赋值
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-11 19:09:31
	'--------------------------------------------------------------
	Public Function SetValue()
		vColID = Request.Form("fColID")
		vTitle = Request.Form("fTitle")
		vAuthor = Request.Form("fAuthor")
		vSource = Request.Form("fSource")
		vSmallPicPath = Request.Form("fSmallPicPath")
		vPicPath = Request.Form("fPicPath")
		vIntro = Request.Form("fIntro")
		vIsTop = Request.Form("fIsTop")
		vState = Request.Form("fState")
		vHits = Request.Form("fHits")
		vCreateTime = Now()
		If Len(vTitle) < 1 Or Len(vTitle) > 50 Then mLastError = "标题的长度请控制在 1 至 50 位" : SetValue = False : Exit Function
		If Not IsNumeric(ColID) Then mLastError = "栏目ID必须为数字" : SetValue = False : Exit Function
		If Cint(ColID) = 0 Then mLastError = "请选择栏目" : SetValue = False : Exit Function
		If Len(vAuthor) = 0 Then vAuthor = ""
		If Len(vSource) = 0 Then vSource = "不详"
		If Len(vPicPath) = 0 Then mLastError = "请先上传图片" : SetValue = False : Exit Function
		If Len(vSmallPicPath) = 0 Then vSmallPicPath = vPicPath
		If Len(vTitle) > 250 Then mLastError = "标题的长度请控制在 250 位" : SetValue = False : Exit Function
		If Len(vIntro) = 0 Then mLastError = "暂无介绍"
		If Len(vIsTop) = 0 Then vIsTop = 0
		If Len(vState) = 0 Then vState = 0
		If Len(vHits) = 0  Then vHits = 0
		SetValue = True
	End Function

	'--------------------------------------------------------------
	' Function name：	LetValue
	' Description: 		从数据库获取数据并赋值
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:36:16
	'--------------------------------------------------------------
	Public Function LetValue()
		Dim Rs
		Set Rs = DB("Select * From [Picture] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要查询的记录 " & vID & " 不存在!" : LetValue = False : Exit Function
		vColID = Rs("ColID")
		vTitle = Rs("Title")
		vAuthor = Rs("Author")
		vSource = Rs("Source")
		vSmallPicPath = Rs("SmallPicPath")
		vPicPath = Rs("PicPath")
		vIntro = Rs("Intro")
		vIsTop = Rs("IsTop")
		vState = Rs("State")
		vHits = Rs("Hits")
		vCreateTime = Rs("CreateTime")
		Rs.Close
		Set Rs = Nothing
		LetValue = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Create()
	' Description: 		创建一个记录
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-7 16:26:26
	'--------------------------------------------------------------
	Public Function Create()
		If SetValue = False Then Create = False: Exit Function
		Dim Rs
		Set Rs = DB("Select * From [Picture]",3)
		Rs.AddNew
		Rs("ColID") = vColID
		Rs("Title") = vTitle
		Rs("Author") = vAuthor
		Rs("Source") = vSource
		Rs("SmallPicPath") = vSmallPicPath
		Rs("PicPath") = vPicPath
		Rs("Intro") = vIntro
		Rs("IsTop") = vIsTop
		Rs("State") = vState
		Rs("Hits") = vHits
		Rs("CreateTime") = Now()
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Create = True
	End Function
	
	'--------------------------------------------------------------
	' Function name：	BatCreate()
	' Description: 		批量创建一个记录
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-12 11:17:02
	'--------------------------------------------------------------
	Public Function BatCreate()
		If SetValue = False Then BatCreate = False: Exit Function
		Dim arrPicPath, i
		arrPicPath = Split(vPicPath, "|")
		For i = 0 To UBound(arrPicPath)
			vPicPath = arrPicPath(i)
			Call DB("INSERT INTO Picture(ColID, Title, Author, Source, SmallPicPath, PicPath, Intro, IsTop, State, Hits, CreateTime) VALUES("&vColID&",'"&vTitle&"','"&vAuthor&"','"&vSource&"','"&vPicPath&"','"&vPicPath&"','"&vIntro&"',"&vIsTop&","&vState&","&vHits&",'"&vCreateTime&"')",0)
		Next
		BatCreate = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Modify()
	' Description: 		修改记录
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-7 16:26:35
	'--------------------------------------------------------------
	Public Function Modify()
		Dim Rs
		Set Rs = DB("Select * From [Picture] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要更新的记录 " & vID & " 不存在!" : Modify = False : Exit Function
		Rs("ColID") = vColID
		Rs("Title") = vTitle
		Rs("Author") = vAuthor
		Rs("Source") = vSource
		Rs("SmallPicPath") = vSmallPicPath
		Rs("PicPath") = vPicPath
		Rs("Intro") = vIntro
		Rs("IsTop") = vIsTop
		Rs("State") = vState
		Rs("Hits") = vHits
		'Rs("CreateTime") = Now()
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Modify = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Delete()
	' Description: 		删除记录
	' Params: 			none
	' Return:			True
	' Create on: 		2009-9-7 16:26:45
	'--------------------------------------------------------------
	Public Function Delete()
		Dim Rs
		Set Rs = DB("Select * From [Picture] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing: mLastError = "你所需要更新的记录 " & vID & " 不存在!" : Delete = False : Exit Function
		'删除文件
		If ExistFile("../"&Rs("SmallPicPath")) Then
			Call DeleteFile("../" & Rs("SmallPicPath"))
		End If
		If ExistFile("../"&Rs("PicPath")) Then
			Call DeleteFile("../" & Rs("PicPath"))
		End If
		Rs.Close : Set Rs = Nothing
		DB "Delete From [Picture] Where [ID] = " & vID ,0
		Delete = True
	End Function

End Class
%>