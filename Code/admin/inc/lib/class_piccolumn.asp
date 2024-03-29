<%
'=========================================================
' Class Name：	ClassPicColumn
' Purpose：		图片栏目类
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-9 10:25:37
' Modify log:	
' Updated on: 	
' CopyRight (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Class ClassPicColumn

	'v前缀：Value，数据库字段的值（类成员）
	Private vID
	Private vName
	Private vInfo
	Private vParentID
	Private vTemplate
	Private vSort
	'm前缀：Menber，类成员
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
	Public Property Get ModifyTime: ModifyTime = vModifyTime: End Property
	'Sort
	Public Property Let Sort(ByVal pSort): vSort = pSort: End Property
	Public Property Get Sort: Sort = vSort: End Property
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
		vName = ""
		vInfo = ""
		vParentID = 0
		vTemplate = ""
		vSort = 0
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
		vName = Request.Form("fName")
		vInfo = Request.Form("fInfo")
		vParentID = Request.Form("fParentID")
		vTemplate = Request.Form("fTemplate")
		vSort = Request.Form("fSort")

		If Len(vName) < 1 Or Len(vName) > 50 Then mLastError = "标题的长度请控制在 1 至 50 位" : SetValue = False : Exit Function
		If Len(vInfo) > 20 Then mLastError = "信息的长度请控制在250 位" : SetValue = False : Exit Function
		If Not IsNumeric(ParentID) Then mLastError = "父栏目ID必须为数字" : SetValue = False : Exit Function
		If Len(vSort) = 0 Then vSort = 0
		If Not IsNumeric(vSort) Then mLastError = "排序必须为数字" : SetValue = False : Exit Function
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
		Set Rs = DB("Select * From [PicColumn] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要查询的记录 " & vID & " 不存在!" : LetValue = False : Exit Function
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
	' Function name：	Create()
	' Description: 		创建一个记录
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-7 16:26:26
	'--------------------------------------------------------------
	Public Function Create()
		Dim Rs
		Set Rs = DB("Select * From [PicColumn]",3)
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
	' Function name：	Modify()
	' Description: 		修改记录
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-7 16:26:35
	'--------------------------------------------------------------
	Public Function Modify()
		'防止父栏目是自己或者自己的子栏目，造成死循环
		If InStr(","&GetColIds(vID, "PICTURE")&",", ","&vParentID&",") > 0 Then mLastError = "父栏目ID不能是自己或者自己子栏目！" : Modify = False : Exit Function
		'更新修改
		Dim Rs
		Set Rs = DB("Select * From [PicColumn] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要更新的记录 " & vID & " 不存在!" : Modify = False : Exit Function
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
	' Function name：	Delete()
	' Description: 		删除记录
	' Params: 			none
	' Return:			True
	' Create on: 		2009-9-7 16:26:45
	'--------------------------------------------------------------
	Public Function Delete()
		DB "UPDATE PicColumn SET ParentID = 0 WHERE ParentID =" & vID ,0
		DB "Delete From [Picture] Where [ColID] = " & vID ,0
		DB "Delete From [PicColumn] Where [ID] = " & vID ,0
		Delete = True
	End Function

End Class
%>