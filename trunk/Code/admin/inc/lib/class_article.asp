<%
'=========================================================
' Class Name��	ClassArticle
' Purpose��		������
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-1 16:13:34
' Modify log:	
' Updated on: 	
' CopyRight (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Class ClassArticle

	'vǰ׺��Value�����ݿ��ֶε�ֵ�����Ա��
	Private vID
	Private vColID
	Private vTitle
	Private vAuthor
	Private vSource
	Private vJumpUrl
	Private vHits
	Private vFocusPic
	Private vContent
	Private vKeyWords
	Private vIsTop
	Private vIsFocusPic
	Private vState
	Private vCreateTime
	Private vModifyTime
	'mǰ׺��Menber�����Ա
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
	'JumpUrl
	Public Property Let JumpUrl(ByVal pJumpUrl): vJumpUrl = pJumpUrl: End Property
	Public Property Get JumpUrl: JumpUrl = vJumpUrl: End Property
	'Hits
	Public Property Let Hits(ByVal pHits): vHits = pHits: End Property
	Public Property Get Hits: Hits = vHits: End Property
	'FocusPic
	Public Property Let FocusPic(ByVal pFocusPic): vFocusPic = pFocusPic: End Property
	Public Property Get FocusPic: FocusPic = vFocusPic: End Property
	'Content
	Public Property Let Content(ByVal pContent): vContent = pContent: End Property
	Public Property Get Content: Content = vContent: End Property
	'KeyWords
	Public Property Let KeyWords(ByVal pKeyWords): vKeyWords = pKeyWords: End Property
	Public Property Get KeyWords: KeyWords = vKeyWords: End Property
	'IsTop
	Public Property Let IsTop(ByVal pIsTop): vIsTop = pIsTop: End Property
	Public Property Get IsTop: IsTop = vIsTop: End Property
	'IsFocusPic
	Public Property Let IsFocusPic(ByVal pIsFocusPic): vIsFocusPic = pIsFocusPic: End Property
	Public Property Get IsFocusPic: IsFocusPic = vIsFocusPic: End Property
	'State
	Public Property Let State(ByVal pState): vState = pState: End Property
	Public Property Get State: State = vState: End Property
	'CreateTime
	Public Property Let CreateTime(ByVal pCreateTime): vCreateTime = pCreateTime: End Property
	Public Property Get CreateTime: CreateTime = vCreateTime: End Property
	'ModifyTime
	Public Property Let ModifyTime(ByVal pModifyTime): vModifyTime = pModifyTime: End Property
	Public Property Get ModifyTime: ModifyTime = vModifyTime: End Property
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
		vColID	= -1
		vTitle = ""
		vAuthor = "����"
		vSource	= "��վ"
		vJumpUrl = ""
		vHits	= 0
		vFocusPic = ""
		vContent = ""
		vKeyWords = ""
		vIsTop = 0
		vIsFocusPic = 0
		vState = 0
		vCreateTime	= Now()
		vModifyTime = Now()
		mLastError = ""
	End Function
	
	'--------------------------------------------------------------
	' Function name��	SetValue
	' Description: 		�ӱ���ȡ���ݲ���ֵ
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:36:16
	'--------------------------------------------------------------
	Public Function SetValue()
		vColID = Request.Form("ColID")
		vTitle = Request.Form("Title")
		vAuthor = Request.Form("Author")
		vSource	= Request.Form("Source")
		vJumpUrl = Request.Form("JumpUrl")
		vHits	= Request.Form("Hits")
		vFocusPic = Request.Form("FocusPic")
		vContent = Request.Form("Content")
		vKeyWords = Request.Form("Keywords")
		vIsTop = Request.Form("IsTop")
		vIsFocusPic = Request.Form("IsFocusPic")
		vState = Request.Form("State")
		vCreateTime	= Now()
		vModifyTime = Now()
		If Len(vTitle) < 3 Or Len(vTitle) > 50 Then mLastError = "����ĳ���������� 3 �� 250 λ" : SetValue = False : Exit Function
		If Not IsNumeric(ColID) Then mLastError = "��ĿID����Ϊ����" : SetValue = False : Exit Function
		If Cint(ColID) = 0 Then mLastError = "��ѡ��������Ŀ" : SetValue = False : Exit Function
		If Len(vAuthor) = 0 Then vAuthor = "����"
		If Len(vHits) = 0 Or Not IsNumeric(vHits) Then vHits = 0
		If Len(vJumpUrl) = 0 And Len(vContent) = 0 Then mLastError = "��ת��ַ�����ݣ�����д��һ��" : SetValue = False : Exit Function
		If Len(vJumpUrl) = 0 And Len(vContent) < 5 Then mLastError = "���ݱ����5���ַ�" : SetValue = False : Exit Function
		If Len(vIsTop) = 0 Then vIsTop = 0
		If Len(vState) = 0 Then vState = 0
		If Len(vIsFocusPic) = 0 Then vIsFocusPic = 0
		If vIsFocusPic = 1 And Len(FocusPic) = 0 Then mLastError = "����Ϊ����ͼƬ������ͼƬ��URL����Ϊ��": SetValue = False : Exit Function
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
		Set Rs = DB("Select * From [Article] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "������Ҫ��ѯ�ļ�¼ " & vID & " ������!" : LetValue = False : Exit Function
		vColID = Rs("ColID")
		vTitle = Rs("Title")
		vAuthor = Rs("Author")
		vSource	= Rs("Source")
		vJumpUrl = Rs("JumpUrl")
		vHits	= Rs("Hits")
		vFocusPic = Rs("FocusPic")
		vContent = Rs("Content")
		vKeyWords = Rs("Keywords")
		vIsTop = Rs("IsTop")
		vState = Rs("State")
		vIsFocusPic = Rs("IsFocusPic")
		vCreateTime	= Rs("CreateTime")
		vModifyTime = Rs("ModifyTime")
		Rs.Close
		Set Rs = Nothing
		LetValue = True
	End Function

	'--------------------------------------------------------------
	' Function name��	Create()
	' Description: 		������¼
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:40:46
	'--------------------------------------------------------------
	Public Function Create()
		If SetValue = False Then Create = False: Exit Function
		Dim Rs
		Set Rs = DB("Select * From [Article]",3)
		Rs.AddNew
		Rs("ColID") = vColID
		Rs("Title") = vTitle
		Rs("Author") = vAuthor
		Rs("Source") = vSource
		Rs("JumpUrl") = vJumpUrl
		Rs("Hits")	= vHits
		Rs("FocusPic") = vFocusPic
		Rs("Content") = vContent
		Rs("Keywords") = vKeyWords
		Rs("IsTop") = vIsTop
		Rs("IsFocusPic") = vIsFocusPic
		Rs("State") = vState
		Rs("CreateTime") = vCreateTime
		Rs("ModifyTime") = vModifyTime
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Create = True
	End Function

	'--------------------------------------------------------------
	' Function name��	Modify()
	' Description: 		�޸ļ�¼
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function Modify()
		Dim Rs
		Set Rs = DB("Select * From [Article] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "������Ҫ���µļ�¼ " & vID & " ������!" : Modify = False : Exit Function
		Rs("ColID") = vColID
		Rs("Title") = vTitle
		Rs("Author") = vAuthor
		Rs("Source") = vSource
		Rs("JumpUrl") = vJumpUrl
		Rs("Hits")	= vHits
		Rs("FocusPic") = vFocusPic
		Rs("Content") = vContent
		Rs("Keywords") = vKeyWords
		Rs("IsTop") = vIsTop
		Rs("IsFocusPic") = vIsFocusPic
		Rs("State") = vState
		Rs("CreateTime") = vCreateTime
		Rs("ModifyTime") = vModifyTime
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Modify = True
	End Function

	'--------------------------------------------------------------
	' Function name��	Delete()
	' Description: 		ɾ����¼
	' Params: 			none
	' Return:			True
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function Delete()
		DB "Delete From [Article] Where [ID] In" & vID ,0
		Delete = True
	End Function

End Class
%>