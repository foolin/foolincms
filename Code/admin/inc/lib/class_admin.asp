<%
'=========================================================
' Class Name：	ClassAdmin
' Purpose：		管理员类
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-8-28 16:31:33
' Modify log:	成员改为私有属性
' Updated on: 	2009-9-1 15:30:45
' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================
Class ClassAdmin

	'v前缀：Value，数据库字段的值（类成员）
	Private vID
	Private vUsername
	Private vNickname
	Private vPassword
	Private vLevel
	Private vLoginTime
	Private vLoginCount
	Private vLoginIP
	'm前缀：Menber，类成员
	Dim mLastError
	
	'ID
	Public Property Let ID(ByVal pID): vID = pID: End Property
	Public Property Get ID: ID = vID: End Property
	'Username
	Public Property Let Username(ByVal pUsername): vUsername = pUsername: End Property
	Public Property Get Username: Username = vUsername: End Property
	'Nickname
	Public Property Let Nickname(ByVal pNickname): vNickname = pNickname: End Property
	Public Property Get Nickname: Nickname = vNickname: End Property
	'Password
	Public Property Let Password(ByVal pPassword): vPassword = pPassword: End Property
	Public Property Get Password: Password = vPassword: End Property
	'Level
	Public Property Let Level(ByVal pLevel): vLevel = pLevel: End Property
	Public Property Get Level: Level = vLevel: End Property
	'LoginTime
	Public Property Let LoginTime(ByVal pLoginTime): vLoginTime = pLoginTime: End Property
	Public Property Get LoginTime: LoginTime = vLoginTime: End Property
	'LoginCount
	Public Property Let LoginCount(ByVal pLoginCount): vLoginCount = pLoginCount: End Property
	Public Property Get LoginCount: LoginCount = vLoginCount: End Property
	'LoginIP
	Public Property Let LoginIP(ByVal pLoginIP): vLoginIP = pLoginIP: End Property
	Public Property Get LoginIP: LoginIP = vLoginIP: End Property
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
		vUserName = ""
		vNickName = ""
		vPassword = ""
		vLoginTime = Now()
		vLoginCount = 0
		vLevel = 0
		vLoginIP = GetIP()
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
		vUsername = Request.Form("fUsername")
		vNickname = Request.Form("fNickname")
		If Len(Request.Form("fPassword"))<>0 And Len(Request.Form("fPassword"))<6 Then mLastError = "密码至少要6位" : SetValue = False : Exit Function
		If Len(Request.Form("fPassword"))<>0 And Request.Form("fPassword")<>Request.Form("fRePassword") Then mLastError = "密码前后不一致": SetValue = False : Exit Function
		vPassword = MD5(Request.Form("fPassword"))
		vLevel = Request.Form("fLevel")
		If Len(vUsername) < 3 Or Len(vUsername) > 20 Then mLastError = "管理员帐号的长度请控制在 3 至 20 位" : SetValue = False : Exit Function
		If Len(vNickname) < 1 Or Len(vNickname) > 20 Then mLastError = "管理员帐号的长度请控制在 1 至 20 位" : SetValue = False : Exit Function
		If Not IsNumeric(vLevel) Then mLastError = "管理员等级必须为数字" : SetValue = False : Exit Function
		If Len(vNickname) = 0 Then vNickname = vUsername
		vLoginTime = Now()
		vLoginCount = 0
		vLoginIP = GetIP()
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
		Set Rs = DB("Select * From [Admin] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要查询的记录 " & vID & " 不存在!" : LetValue = False : Exit Function
		vUsername = Rs("Username")
		vNickname = Rs("Nickname")
		vPassword = Rs("Password")
		vLevel = Rs("Level")
		vLoginTime = Rs("LoginTime")
		vLoginCount = Rs("LoginCount")
		vLoginIP = Rs("LoginIP")
		Rs.Close
		Set Rs = Nothing
		LetValue = True
	End Function
	
	Public Function Exist(Byval tUsername)
		Dim Rs, isExist
		Set Rs = DB("Select * From [Admin] Where [Username]=" & tUsername,1)
		If Rs.Eof Then
			isExist = False
		Else
			isExist = True
		End If
		Rs.Close : Set Rs = Nothing
		Exist = isExist
	End Function

	'--------------------------------------------------------------
	' Function name：	Create()
	' Description: 		创建一个管理员
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:40:46
	'--------------------------------------------------------------
	Public Function Create()
		Dim Rs
		Set Rs = DB("Select [ID] From [Admin] Where [Username]='" & vUsername & "'",1)
		If Not Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "管理员帐号 的值 " & vUsername & " 已存在!" : Create = False : Exit Function
		Set Rs = DB("Select * From [Admin]",3)
		Rs.AddNew
		Rs("Username") = vUsername
		Rs("Nickname") = vNickname
		Rs("Password") = vPassword
		Rs("Level") = vLevel
		Rs("LoginTime") = vLoginTime
		Rs("LoginCount") = vLoginCount
		Rs("LoginIP") = vLoginIP
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Create = True
	End Function

	'--------------------------------------------------------------
	' Function name：	Modify()
	' Description: 		修改帐号全部信息
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function Modify()
		Dim Rs
		Set Rs = DB("Select * From [Admin] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要更新的记录 " & vID & " 不存在!" : Modify = False : Exit Function
		Rs("Username") = vUsername
		Rs("Nickname") = vNickname
		Rs("Password") = vPassword
		Rs("Level") = vLevel
		Rs("LoginTime") = vLoginTime
		Rs("LoginCount") = vLoginCount
		Rs("LoginIP") = vLoginIP
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		Modify = True
	End Function
	
	'--------------------------------------------------------------
	' Function name：	ModifyPsw()
	' Description: 		修改帐号密码
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function ModifyPsw()
		Dim Rs
		Set Rs = DB("Select * From [Admin] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要更新的记录 " & vID & " 不存在!" : ModifyPsw = False : Exit Function
		Rs("Username") = vUsername
		Rs("Nickname") = vNickname
		Rs("Password") = vPassword
		Rs("Level") = vLevel
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		ModifyPsw = True
	End Function
	
	'--------------------------------------------------------------
	' Function name：	ModifyInfo()
	' Description: 		修改帐号信息，不修改密码
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-20 10:05:55
	'--------------------------------------------------------------
	Public Function ModifyInfo()
		Dim Rs
		Set Rs = DB("Select * From [Admin] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "你所需要更新的记录 " & vID & " 不存在!" : ModifyInfo = False : Exit Function
		Rs("Username") = vUsername
		Rs("Nickname") = vNickname
		Rs("Level") = vLevel
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		ModifyInfo = True
	End Function
	

	'--------------------------------------------------------------
	' Function name：	Delete()
	' Description: 		删除管理员
	' Params: 			none
	' Return:			True
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function Delete()
		DB "Delete From [Admin] Where [ID] IN ("&  vID &")" ,0
		Delete = True
	End Function
	
	'冻结用户
	Public Function Freeze()
		Call DB("UPDATE [Admin] SET [Level]=-1 WHERE [ID] IN ("&  vID &")" ,0)
		Freeze = True
	End Function
	
	'解冻用户
	Public Function Unfreeze()
		Call DB("UPDATE [Admin] SET [Level]=0 WHERE [ID] IN ("&  vID &")" ,0)
		Unfreeze = True
	End Function
End Class
%>