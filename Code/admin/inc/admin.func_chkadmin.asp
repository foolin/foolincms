<%
Function ChkLogin()
	Dim blnFlag: blnFlag = False
	Dim strAdminName, strAdminPassword
	strAdminName = GetLogin("AdminName")
	strAdminPassword = GetLogin("AdminPassword")
	If strAdminName="" And strAdminPassword="" Then
		Response.write "<script type='text/javascript'>alert('����δ��¼');window.close();this.top.location.href='login.asp';history.go(-1);</script>"
		Response.End()
	End If
	'��֤��½
	Dim Rs
	Set Rs = DB("SELECT Password,LoginTime FROM Admin WHERE Username = '"& strAdminName &"'", 1)
	If Rs.Eof Then
		Call MsgBox("��δ��¼���ߵ�¼��ʱ","logout.asp")
		Response.End()
	End If
	If Md5(Rs("Password")&Rs("LoginTime"))<> strAdminPassword Then
		Call MsgBox("�Ƿ���¼","logout.asp")
		Response.End()
	End If
	blnFlag = True
	ChkLogin = blnFlag
End Function

'���Ȩ�޺�����chkType������ͣ�chkAct-���Ĳ���
'chkType = article,picture,guestbook,mytag,diypage,template,config,weblog,admin
'chkAct = view,create,modify,delete,all
Function ChkPower(Byval chkType, Byval chkAct)
	ChkLogin()
	Dim bFlag: bFlag = False
	Dim UserLevel: UserLevel = Cint(GetLogin("AdminLevel"))
	Dim LowPower, NormalPower, HightPower, SuperPower
	LowPower = "|article|picture|guestbook|"	'��������Ա
	NormalPower = "|article|picture|guestbook|artcolumn|piccolumn|"	'��ͨ����Ա
	HightPower = "|article|picture|guestbook|artcolumn|piccolumn|mytag|diypage|config|weblog|"	'�߼�����Ա
	SuperPower = "allpower"	'�߼�����Ա template|admin_user
	Select Case UserLevel
		Case 3	'��������Ա
			bFlag = True
		Case 2	'�߼�����Ա
			If InStr(HightPower, "|" & LCase(chkType) & "|") > 0 Then
				bFlag = True
				
			Else
				bFlag = False
			End If
		Case 1	'�м�����Ա
			If InStr(NormalPower, "|" & LCase(chkType) & "|") > 0 Then
				bFlag = True
			Else
				bFlag = False
			End If
		Case 0	'��ͨ����Ա
			If InStr(LowPower, "|" & LCase(chkType) & "|") > 0 Then
				bFlag = True
			Else
				bFlag = False
			End If
		Case -1	'�����û�
			bFlag = False
			Call MsgBox("���ʻ��Ѿ������ᣬ����ϵ����Ա��", "Logout.asp")
		Case Else
			bFlag = False
	End Select
	If bFlag = False Then
		Call MsgBox("�Բ�����û��Ȩ�ޣ�", "BACK")
	End If
	ChkPower = bFlag
End Function
%>
