<%
'=========================================================
' Class Name��	ClassGuestBook
' Purpose��		������
' Auhtor: 		Foolin
' Email: 		Foolin@126.com
' Createed on: 	2009-9-27 8:30:47
' Modify log:	
' Updated on: 	
' CopyRight (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Class ClassGuestBook

	'vǰ׺��Value�����ݿ��ֶε�ֵ�����Ա��
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
	'mǰ׺��Menber�����Ա
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
		Call ChkLogin()		'����¼
		Call Initialize()	'��ʼ��
	End Sub

	Private Sub Class_Terminate()
		Call Initialize()
	End Sub

	Public Function Initialize()
		vID = -1
		vUser	= "����"
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
	' Function name��	SetValue
	' Description: 		�ӱ���ȡ���ݲ���ֵ
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
		If Len(vTitle) < 1 Or Len(vTitle) > 50 Then mLastError = "����ĳ���������� 1 �� 50 λ" : SetValue = False : Exit Function
		If Len(User) = 0 Then User = "����"
		If Len(vContent)<1 Or Len(Content)>250 Then mLastError = "�������ݵĳ���������� 1 �� 250 λ" : SetValue = False : Exit Function
		If Len(vRecomment)<1 Or Len(vRecomment)>250 Then mLastError = "�ظ����ݵĳ���������� 1 �� 250 λ" : SetValue = False : Exit Function
		If Len(vState) = 0 Then vState = 0
		SetValue = True
	End Function

	'--------------------------------------------------------------
	' Function name��	LetValue
	' Description: 		�����ݿ��ȡ���ݲ���ֵ
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:36:16
	'--------------------------------------------------------------
	Public Function LetValue()
		Dim Rs
		Set Rs = DB("Select * From [GuestBook] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "������Ҫ��ѯ�ļ�¼ " & vID & " ������!" : LetValue = False : Exit Function
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
	' Function name��	Create()
	' Description: 		������¼
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
	' Function name��	Comment()
	' Description: 		�ظ�����
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-27 8:45:26
	'--------------------------------------------------------------
	Public Function Comment()
		'���ղ���
		vRecomment	= Request.Form("fRecomment")
		vReUser = Request.Form("fReUser")
		vReTime = Now()
		If Len(vRecomment)<1 Or Len(vRecomment)>250 Then mLastError = "�ظ����ݵĳ���������� 1 �� 250 λ" : Comment = False : Exit Function
		'�ظ�����
		Dim Rs
		Set Rs = DB("Select * From [GuestBook] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "������Ҫ���µļ�¼ " & vID & " ������!" : Comment = False : Exit Function
		Rs("Recomment")	= vRecomment
		Rs("ReUser") = vReUser
		Rs("ReTime") = vReTime
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Comment = True
	End Function

	'--------------------------------------------------------------
	' Function name��	Delete()
	' Description: 		ɾ����¼
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