<%
'=========================================================
' Class Name：	ClassDiyPage
' Purpose：		自定义页面
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-16 16:08:17
' Modify log:	
' Updated on: 	
' CopyRight (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Class ClassDiyPage

	'v前缀：Value，数据库字段的值（类成员）
	Private vID
	Private vTitle
	Private vPageName
	Private vKeywords
	Private vTemplate
	Private vCode
	Private vState
	Private vIsSystem
	'm前缀：Menber，类成员
	Dim mLastError
	
	'ID
	Public Property Let ID(ByVal pID): vID = pID: End Property
	Public Property Get ID: ID = vID: End Property
	'Title
	Public Property Let Title(ByVal pTitle): vTitle = pTitle: End Property
	Public Property Get Title: Title = vTitle: End Property
	'PageName
	Public Property Let PageName(ByVal pPageName): vPageName = pPageName: End Property
	Public Property Get PageName: PageName = vPageName: End Property
	'Keywords
	Public Property Let Keywords(ByVal pKeywords): vKeywords = pKeywords: End Property
	Public Property Get Keywords: Keywords = vKeywords: End Property
	'Template
	Public Property Let Template(ByVal pTemplate): vTemplate = pTemplate: End Property
	Public Property Get Template: Template = vTemplate: End Property
	'Code
	Public Property Let Code(ByVal pCode): vCode = pCode: End Property
	Public Property Get Code: Code = vCode: End Property
	'State
	Public Property Let State(ByVal pState): vState = pState: End Property
	Public Property Get State: State = vState: End Property
	'IsSystem
	Public Property Let IsSystem(ByVal pIsSystem): vIsSystem = pIsSystem: End Property
	Public Property Get IsSystem: IsSystem = vIsSystem: End Property
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
		vTitle = ""
		vPageName = ""
		vKeywords = ""
		vTemplate = ""
		vCode = ""
		vState = 0
		vIsSystem = 0
		mLastError = ""
	End Function
	
	'--------------------------------------------------------------
	' Function name：	SetValue
	' Description: 		从表单获取数据并赋值
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:36:16
	'--------------------------------------------------------------
	Public Function SetValue()
		vTitle = Request.Form("fTitle")
		vKeywords = Request.Form("fKeywords")
		vPageName = Request.Form("fPageName")
		vTemplate = Request.Form("fTemplate")
		vCode = Request.Form("fCode")
		vIsSystem = Request.Form("fIsSystem")
		vState = Request.Form("fState")
		If Len(vTitle) < 3 Or Len(vTitle) > 50 Then mLastError = "标题的长度请控制在 3 至 50 位" : SetValue = False : Exit Function
		If Len(vCode) = 0 Then mLastError = "页面内容不能为空" : SetValue = False : Exit Function
		If Len(vKeywords) = 0 Then vKeywords = ""
		If Len(vPageName) = 0 Then vPageName = ""
		If Len(vTemplate) = 0 Then vTemplate = ""
		If Len(vState) = 0 Then vState = 0
		If Len(vIsSystem) = 0 Then vIsSystem = 0
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
		Set Rs = DB("Select * From [DiyPage] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要查询的记录 " & vID & " 不存在!" : LetValue = False : Exit Function
		vTitle = Rs("Title")
		vKeywords = Rs("Keywords")
		vPageName = Rs("PageName")
		vTemplate = Rs("Template")
		vCode = Rs("Code")
		vState = Rs("State")
		vIsSystem = Rs("IsSystem")
		Rs.Close
		Set Rs = Nothing
		LetValue = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Create()
	' Description: 		创建记录
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:40:46
	'--------------------------------------------------------------
	Public Function Create()
		'If SetValue = False Then Create = False: Exit Function
		Dim Rs
		Set Rs = DB("Select * From [DiyPage]",3)
		Rs.AddNew
		Rs("Title") = vTitle
		Rs("Keywords") = vKeywords
		Rs("PageName") = vPageName
		Rs("Template") = vTemplate
		Rs("Code") = vCode
		Rs("Keywords") = vKeyWords
		Rs("State") = vState
		Rs("IsSystem") = vIsSystem
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Create = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Modify()
	' Description: 		修改记录
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function Modify()
		Dim Rs
		Set Rs = DB("Select * From [DiyPage] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要更新的记录 " & vID & " 不存在!" : Modify = False : Exit Function
		Rs("Title") = vTitle
		Rs("Keywords") = vKeywords
		Rs("PageName") = vPageName
		Rs("Template") = vTemplate
		Rs("Code") = vCode
		Rs("Keywords") = vKeyWords
		Rs("IsSystem") = vIsSystem
		Rs("State") = vState
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
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function Delete()
		Dim Rs
		Set Rs = DB("Select ID,IsSystem From [DiyPage] Where IsSystem = 1 And [ID] IN(" & vID &")",3)
		If Not Rs.Eof Then mLastError = "你删除记录[" & Rs("ID") & "]是系统定义页面!": Rs.Close: Set Rs = Nothing :  Delete = False : Exit Function
		DB "Delete From [DiyPage] Where [ID] IN(" & Rs("ID") &")" ,0
		Delete = True
	End Function

End Class
%>