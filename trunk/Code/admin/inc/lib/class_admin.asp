<%
'=========================================================
' Class Name��	ClassAdmin
' Purpose��		����Ա��
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-8-28 16:31:33
' Modify log:	��Ա��Ϊ˽������
' Updated on: 	2009-9-1 15:30:45
' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================
Class ClassAdmin

	'vǰ׺��Value�����ݿ��ֶε�ֵ�����Ա��
	Private vID
	Private vUsername
	Private vNickname
	Private vPassword
	Private vLevel
	Private vLoginTime
	Private vLoginCount
	Private vLoginIP
	'mǰ׺��Menber�����Ա
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
		Call ChkLogin()		'����¼
		Call Initialize()	'��ʼ��
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
	' Function name��	SetValue
	' Description: 		�ӱ���ȡ���ݲ���ֵ
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:36:16
	'--------------------------------------------------------------
	Public Function SetValue()
		vUsername = Request.Form("fUsername")
		vNickname = Request.Form("fNickname")
		If Len(Request.Form("fPassword"))<>0 And Len(Request.Form("fPassword"))<6 Then mLastError = "��������Ҫ6λ" : SetValue = False : Exit Function
		If Len(Request.Form("fPassword"))<>0 And Request.Form("fPassword")<>Request.Form("fRePassword") Then mLastError = "����ǰ��һ��": SetValue = False : Exit Function
		vPassword = MD5(Request.Form("fPassword"))
		vLevel = Request.Form("fLevel")
		If Len(vUsername) < 3 Or Len(vUsername) > 20 Then mLastError = "����Ա�ʺŵĳ���������� 3 �� 20 λ" : SetValue = False : Exit Function
		If Len(vNickname) < 1 Or Len(vNickname) > 20 Then mLastError = "����Ա�ʺŵĳ���������� 1 �� 20 λ" : SetValue = False : Exit Function
		If Not IsNumeric(vLevel) Then mLastError = "����Ա�ȼ�����Ϊ����" : SetValue = False : Exit Function
		If Len(vNickname) = 0 Then vNickname = vUsername
		vLoginTime = Now()
		vLoginCount = 0
		vLoginIP = GetIP()
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
		Set Rs = DB("Select * From [Admin] Where [ID]=" & vID,1)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "������Ҫ��ѯ�ļ�¼ " & vID & " ������!" : LetValue = False : Exit Function
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
	' Function name��	Create()
	' Description: 		����һ������Ա
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:40:46
	'--------------------------------------------------------------
	Public Function Create()
		Dim Rs
		Set Rs = DB("Select [ID] From [Admin] Where [Username]='" & vUsername & "'",1)
		If Not Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "����Ա�ʺ� ��ֵ " & vUsername & " �Ѵ���!" : Create = False : Exit Function
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
	' Function name��	Modify()
	' Description: 		�޸��ʺ�ȫ����Ϣ
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function Modify()
		Dim Rs
		Set Rs = DB("Select * From [Admin] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "������Ҫ���µļ�¼ " & vID & " ������!" : Modify = False : Exit Function
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
	' Function name��	ModifyPsw()
	' Description: 		�޸��ʺ�����
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function ModifyPsw()
		Dim Rs
		Set Rs = DB("Select * From [Admin] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "������Ҫ���µļ�¼ " & vID & " ������!" : ModifyPsw = False : Exit Function
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
	' Function name��	ModifyInfo()
	' Description: 		�޸��ʺ���Ϣ�����޸�����
	' Params: 			none
	' Return:			True|Flase
	' Create on: 		2009-9-20 10:05:55
	'--------------------------------------------------------------
	Public Function ModifyInfo()
		Dim Rs
		Set Rs = DB("Select * From [Admin] Where [ID]=" & vID,3)
		If Rs.Eof Then Rs.Close : Set Rs = Nothing : mLastError = "������Ҫ���µļ�¼ " & vID & " ������!" : ModifyInfo = False : Exit Function
		Rs("Username") = vUsername
		Rs("Nickname") = vNickname
		Rs("Level") = vLevel
		Rs.Update
		Rs.Close
		Set Rs = Nothing
		ModifyInfo = True
	End Function
	

	'--------------------------------------------------------------
	' Function name��	Delete()
	' Description: 		ɾ������Ա
	' Params: 			none
	' Return:			True
	' Create on: 		2009-8-28 16:58:31
	'--------------------------------------------------------------
	Public Function Delete()
		DB "Delete From [Admin] Where [ID] IN ("&  vID &")" ,0
		Delete = True
	End Function
	
	'�����û�
	Public Function Freeze()
		Call DB("UPDATE [Admin] SET [Level]=-1 WHERE [ID] IN ("&  vID &")" ,0)
		Freeze = True
	End Function
	
	'�ⶳ�û�
	Public Function Unfreeze()
		Call DB("UPDATE [Admin] SET [Level]=0 WHERE [ID] IN ("&  vID &")" ,0)
		Unfreeze = True
	End Function
End Class
%>