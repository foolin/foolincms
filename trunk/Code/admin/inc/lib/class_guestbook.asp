<%
'=========================================================
' Class Name：	ClassGuestBook
' Purpose：		留言类
' Auhtor: 		Foolin
' Email: 		Foolin@126.com
' Createed on: 	2009-9-27 8:30:47
' Modify log:	
' Updated on: 	
' CopyRight (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Class ClassGuestBook

	'v前缀：Value，数据库字段的值（类成员）
	Private vID
	Private vUser
	Private vTitle
	Private vContent
	Private vEmail
	Private vHomePage
	Private vIp
	Private vCreateTime
	Private vRecomment
	Private vReUser
	Private vReTime
	Private vState
	'm前缀：Menber，类成员
	Dim mLastError
	
	'ID
	Public Property Let ID(ByVal pID): vID = pID: End Property
	Public Property Get ID: ID = vID: End Property
	'User
	Public Property Let User(ByVal pUser): vUser = pUser: End Property
	Public Property Get User: User = vUser: End Property
	'Title
	Public Property Let Title(ByVal pTitle): vTitle = pTitle: End Property
	Public Property Get Title: Title = vTitle: End Property
	'Content
	Public Property Let Content(ByVal pContent): vContent = pContent: End Property
	Public Property Get Content: Content = vContent: End Property
	'Email
	Public Property Let Email(ByVal pEmail): vEmail = pEmail: End Property
	Public Property Get Email: Email = vEmail: End Property
	'HomePage
	Public Property Let HomePage(ByVal pHomePage): vHomePage = pHomePage: End Property
	Public Property Get HomePage: HomePage = vHomePage: End Property
	'Ip
	Public Property Let Ip(ByVal pIp): vIp = pIp: End Property
	Public Property Get Ip: Ip = vIp: End Property
	'CreateTime
	Public Property Let CreateTime(ByVal pCreateTime): vCreateTime = pCreateTime: End Property
	Public Property Get CreateTime: CreateTime = vCreateTime: End Property
	'Recomment
	Public Property Let Recomment(ByVal pRecomment): vRecomment = pRecomment: End Property
	Public Property Get Recomment: Recomment = vRecomment: End Property
	'ReUser
	Public Property Let ReUser(ByVal pReUser): vReUser = pReUser: End Property
	Public Property Get ReUser: ReUser = vReUser: End Property
	'ReTime
	Public Property Let ReTime(ByVal pReTime): vReTime = pReTime: End Property
	Public Property Get ReTime: ReTime = vReTime: End Property
	'State
	Public Property Let State(ByVal pState): vState = pState: End Property
	Public Property Get State: State = vState: End Property
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
		vID = -1
		vUser	= "匿名"
		vTitle = ""
		vContent = ""
		vEmail = "@"
		vHomePage	= "#"
		vIp = ""
		vCreateTime	= Now()
		vRecomment	= 0
		vReUser = ""
		vReTime = ""
		vState = 0
		mLastError = ""
	End Function
	
	'--------------------------------------------------------------
	' Function name：	SetValue
	' Description: 		从表单获取数据并赋值
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-27 8:41:23
	'--------------------------------------------------------------
	Public Function SetValue()
		vUser = Request.Form("fUser")
		vTitle = Request.Form("fTitle")
		vEmail = Request.Form("fEmail")
		vHomePage	= Request.Form("HfomePage")
		vIp = GetIP()
		vRecomment	= Request.Form("fRecomment")
		vReUser = Request.Form("fReUser")
		vContent = Request.Form("fContent")
		vState = Request.Form("fState")
		vCreateTime	= Now()
		vReTime = Now()
		If Len(vTitle) < 1 Or Len(vTitle) > 50 Then mLastError = "标题的长度请控制在 1 至 50 位" : SetValue = False : Exit Function
		If Len(User) = 0 Then User = "匿名"
		If Len(vContent)<1 Or Len(Content)>250 Then mLastError = "留言内容的长度请控制在 1 至 250 位" : SetValue = False : Exit Function
		If Len(vRecomment)<1 Or Len(vRecomment)>250 Then mLastError = "回复内容的长度请控制在 1 至 250 位" : SetValue = False : Exit Function
		If Len(vState) = 0 Then vState = 0
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
		Set Rs = DB("Select * From [GuestBook] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要查询的记录 " & vID & " 不存在!" : LetValue = False : Exit Function
		vUser = Rs("User")
		vTitle = Rs("Title")
		vContent = Rs("Content")
		vEmail = Rs("Email")
		vHomePage	= Rs("HomePage")
		vIp = Rs("Ip")
		vCreateTime	= Rs("CreateTime")
		vRecomment	= Rs("Recomment")
		vReUser = Rs("ReUser")
		vReTime = Rs("ReTime")
		vState = Rs("State")
		Rs.Close
		Set Rs = Nothing
		LetValue = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Create()
	' Description: 		创建记录
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-27 8:42:22
	'--------------------------------------------------------------
	Public Function Create()
		If SetValue = False Then Create = False: Exit Function
		Dim Rs
		Set Rs = DB("Select * From [GuestBook]",3)
		Rs.AddNew
		Rs("User") = vUser
		Rs("Title") = vTitle
		Rs("Content") = vContent
		Rs("Email") = vEmail
		Rs("HomePage") = vHomePage
		Rs("Ip") = vIp
		Rs("State") = vState
		Rs("CreateTime") = vCreateTime
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Create = True
	End Function
	
	'--------------------------------------------------------------
	' Function name：	Comment()
	' Description: 		回复留言
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-27 8:45:26
	'--------------------------------------------------------------
	Public Function Comment()
		'接收参数
		vRecomment	= Request.Form("fRecomment")
		vReUser = Request.Form("fReUser")
		vReTime = Now()
		If Len(vRecomment)<1 Or Len(vRecomment)>250 Then mLastError = "回复内容的长度请控制在 1 至 250 位" : Comment = False : Exit Function
		'回复留言
		Dim Rs
		Set Rs = DB("Select * From [GuestBook] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要更新的记录 " & vID & " 不存在!" : Comment = False : Exit Function
		Rs("Recomment")	= vRecomment
		Rs("ReUser") = vReUser
		Rs("ReTime") = vReTime
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Comment = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Delete()
	' Description: 		删除记录
	' Params: 			none
	' Return:			True
	' Create on: 		2009-9-27 8:47:13
	'--------------------------------------------------------------
	Public Function Delete()
		DB "Delete From [GuestBook] Where [ID] In (" & vID  & ")" ,0
		Delete = True
	End Function

End Class
%>