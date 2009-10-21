<%
Function ChkLogin()
	Dim blnFlag: blnFlag = False
	Dim strAdminName, strAdminPassword
	strAdminName = GetLogin("AdminName")
	strAdminPassword = GetLogin("AdminPassword")
	If strAdminName="" And strAdminPassword="" Then
		Response.write "<script type='text/javascript'>alert('你尚未登录');window.close();this.top.location.href='login.asp';history.go(-1);</script>"
		Response.End()
	End If
	'验证登陆
	Dim Rs
	Set Rs = DB("SELECT Password,LoginTime FROM Admin WHERE Username = '"& strAdminName &"'", 1)
	If Rs.Eof Then
		Call MsgBox("尚未登录或者登录超时","logout.asp")
		Response.End()
	End If
	If Md5(Rs("Password")&Rs("LoginTime"))<> strAdminPassword Then
		Call MsgBox("非法登录","logout.asp")
		Response.End()
	End If
	blnFlag = True
	ChkLogin = blnFlag
End Function

'检查权限函数，chkType检查类型，chkAct-检查的操作
'chkType = article,picture,guestbook,mytag,diypage,template,config,weblog,admin
'chkAct = view,create,modify,delete,all
Function ChkPower(Byval chkType, Byval chkAct)
	ChkLogin()
	Dim bFlag: bFlag = False
	Dim UserLevel: UserLevel = Cint(GetLogin("AdminLevel"))
	Dim LowPower, NormalPower, HightPower, SuperPower
	LowPower = "|article|picture|guestbook|"	'初级管理员
	NormalPower = "|article|picture|guestbook|artcolumn|piccolumn|"	'普通管理员
	HightPower = "|article|picture|guestbook|artcolumn|piccolumn|mytag|diypage|config|weblog|"	'高级管理员
	SuperPower = "allpower"	'高级管理员 template|admin_user
	Select Case UserLevel
		Case 3	'超级管理员
			bFlag = True
		Case 2	'高级管理员
			If InStr(HightPower, "|" & LCase(chkType) & "|") > 0 Then
				bFlag = True
				
			Else
				bFlag = False
			End If
		Case 1	'中级管理员
			If InStr(NormalPower, "|" & LCase(chkType) & "|") > 0 Then
				bFlag = True
			Else
				bFlag = False
			End If
		Case 0	'普通管理员
			If InStr(LowPower, "|" & LCase(chkType) & "|") > 0 Then
				bFlag = True
			Else
				bFlag = False
			End If
		Case -1	'冻结用户
			bFlag = False
			Call MsgBox("您帐户已经被冻结，请联系管理员！", "Logout.asp")
		Case Else
			bFlag = False
	End Select
	If bFlag = False Then
		Call MsgBox("对不起，您没有权限！", "BACK")
	End If
	ChkPower = bFlag
End Function
%>
