<%
'=========================================================
' Class Name��	ClassArtColumn
' Purpose��		������Ŀ��
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-7 16:26:07
' Modify log:	
' Updated on: 	
' CopyRight (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Class ClassArtColumn

	'vǰ׺��Value�����ݿ��ֶε�ֵ�����Ա��
	Private vID
	Private vName
	Private vInfo
	Private vParentID
	Private vTemplate
	Private vSort
	'mǰ׺��Menber�����Ա
	Dim mLastError
	
	'ID
	Public Property Let ID(ByVal pID): vID = pID: End Property
	Public Property Get ID: ID = vID: End Property
	'Name
	Public Property Let Name(ByVal pName): vName = pName: End Property
	Public Property Get Name: Name = vName: End Property
	'Info
	Public Property Let Info(ByVal pInfo): vInfo = pInfo: End Property
	Public Property Get Info: Info = vInfo: End Property
	'ParentID
	Public Property Let ParentID(ByVal pParentID): vParentID = pParentID: End Property
	Public Property Get ParentID: ParentID = vParentID: End Property
	'Template
	Public Property Let Template(ByVal pTemplate): vTemplate = pTemplate: End Property
	Public Property Get Template: Template = vTemplate: End Property
	'Sort
	Public Property Let Sort(ByVal pSort): vSort = pSort: End Property
	Public Property Get Sort: Sort = vSort: End Property
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
		vName = ""
		vInfo = ""
		vParentID = 0
		vTemplate = ""
		vSort = 0
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
		vName = Request.Form("fName")
		vInfo = Request.Form("fInfo")
		vParentID = Request.Form("fParentID")
		vTemplate = Request.Form("fTemplate")
		vSort = Request.Form("fSort")

		If Len(vName)<1 Or Len(vName) > 50 Then mLastError = "����ĳ���������� 1 �� 50 λ" : SetValue = False : Exit Function
		If Len(vInfo)>250 Then mLastError = "��Ϣ�ĳ����������250 λ" : SetValue = False : Exit Function
		If Not IsNumeric(ParentID) Then mLastError = "����ĿID����Ϊ����" : SetValue = False : Exit Function
		If Len(vTemplate) > 20 Then mLastError = "ģ��ĳ����������20 λ" : SetValue = False : Exit Function
		If Len(vSort) = 0 Then vSort = 0
		If Not IsNumeric(vSort) Then mLastError = "�������Ϊ����" : SetValue = False : Exit Function
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
		Set Rs = DB("Select * From [ArtColumn] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "������Ҫ��ѯ�ļ�¼ " & vID & " ������!" : LetValue = False : Exit Function
		vName = Rs("Name")
		vInfo = Rs("Info")
		vParentID = Rs("ParentID")
		vTemplate = Rs("Template")
		vSort = Rs("Sort")
		Rs.Close
		Set Rs = Nothing
		LetValue = True
	End Function

	'--------------------------------------------------------------
	' Function name��	Create()
	' Description: 		����һ����¼
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-7 16:26:26
	'--------------------------------------------------------------
	Public Function Create()
		Dim Rs
		Set Rs = DB("Select * From [ArtColumn]",3)
		Rs.AddNew
		Rs("Name") = vName
		Rs("Info") = vInfo
		Rs("ParentID") = vParentID
		Rs("Template") = vTemplate
		Rs("Sort") = vSort
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
	' Create on: 		2009-9-7 16:26:35
	'--------------------------------------------------------------
	Public Function Modify()
		'��ֹ����Ŀ���Լ������Լ�������Ŀ�������ѭ��
		If InStr(","&GetColIds(vID, "ARTICLE")&",", ","&vParentID&",") > 0 Then mLastError = "����ĿID�������Լ������Լ�����Ŀ��" : Modify = False : Exit Function
		Dim Rs
		Set Rs = DB("Select * From [ArtColumn] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "������Ҫ���µļ�¼ " & vID & " ������!" : Modify = False : Exit Function
		Rs("Name") = vName
		Rs("Info") = vInfo
		Rs("ParentID") = vParentID
		Rs("Template") = vTemplate
		Rs("Sort") = vSort
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
	' Create on: 		2009-9-7 16:26:45
	'--------------------------------------------------------------
	Public Function Delete()
		DB "UPDATE ArtColumn SET ParentID = 0 WHERE ParentID =" & vID ,0
		DB "Delete From [Article] Where [ColID] = " & vID ,0
		DB "Delete From [ArtColumn] Where [ID] = " & vID ,0
		Delete = True
	End Function

End Class
%>