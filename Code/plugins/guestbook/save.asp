<!--#include file="../../inc/config.asp"-->
<!--#include file="../../inc/func_common.asp"-->
<!--#include file="../inc/plugin.conn.asp"-->
<%	'�ж��Ƿ񿪷�����
	If ISOPENGBOOK = 0 Then
		Call MsgBox("�Բ��𣬱�վ���Թ����ݲ����ţ�","BACK")
	End If
	'�ж�������Դ�Ƿ�Ƿ�
	Dim serverUrl1, serverUrl2
	serverUrl1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
	serverUrl2 = Cstr(Request.ServerVariables("SERVER_NAME"))
	If Mid(serverUrl1, 8, Len(serverUrl2)) <>  serverUrl2 Then
		Call MsgBox("�Բ��𣬱�վ��ֹ�ⲿ�ύ���ݣ�","REFRESH")
	End If
	'����Ƿ�Ƶ������
	If DateDIff("s", CDate(Session("PostTime")), Now()) < GBOOKTIME Then
		Call MsgBox("��ֹƵ���ύ���ԣ���"& GBOOKTIME &"��֮��������...","BACK")
	End If
	'�����֤��
	If Len(Request("fChkCode"))=0 Then Call MsgBox("��֤�벻��Ϊ�գ�", "BACK")
	If Request("fChkCode") <> Session("ChkCode") Then
		Call MsgBox("��֤�벻��ȷ��", "BACK")
	End If
	'��������
	Dim strTitle, strContent, strUser, strEmail, strHomepage, strIP, strSql
		strTitle = Req("fTitle")
		strContent = Req("fContent")
		strUser = Left(Req("fUser"),20)
		strEmail = Left(Req("fEmail"),20)
		strHomePage = Left(Req("fHomePage"),50)
		strIP = GetIP()
		If Len(strTitle)<1 Or Len(strContent)>50 Then Call MsgBox("����ĳ���������� 1 �� 50 λ","BACK")
		If Len(strContent)<1 Or Len(strContent)>250 Then Call MsgBox("���ݵĳ���������� 1 �� 250 λ","BACK")
		If Len(strUser) = 0 Then strUser = "����"
		If Len(strEmail) = 0 Then strEmail = "@"
		If Len(strHomePage) = 0 Then strHomePage = "#"
		'���������ַ�
		strTitle = FilterDirtyStr(strTitle)
		strContent = FilterDirtyStr(strContent)
		'�ж��Ƿ���Ҫ���
		If ISAUDITGBOOK = 0 Then
			strSql = "INSERT INTO GuestBook([Title],[Content],[User],[Email],[HomePage],[IP],[CreateTime], [State])"
			strSql = strSql & " VALUES('"& strTitle &"','"& strContent &"','"& strUser &"','"& strEmail &"','"& strHomePage &"','"& strIP &"','"& Now() &"', 1)"
		Else
			strSql = "INSERT INTO GuestBook([Title],[Content],[User],[Email],[HomePage],[IP],[CreateTime], [State])"
			strSql = strSql & " VALUES('"& strTitle &"','"& strContent &"','"& strUser &"','"& strEmail &"','"& strHomePage &"','"& strIP &"','"& Now() &"', 0)"
		End If
	Call DB(strSql,0)
	Session("PostTime") = Now()	'��¼�ύʱ��
	Call MsgBox("��������Ѿ��ύ����л�������ԣ�", "REFRESH")
%>
