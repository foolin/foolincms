<%
'=========================================================
' File Name��	func_file.asp
' Purpose��		�ļ����ò�������
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-13 14:59:09
' Version:		v1.0.0 Build 20090913
' CopyRight (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved
'=========================================================

'�Ƿ�����ļ�
Function ExistFile(ByVal strFilePath)
	Dim objFso, isExist
	Set objFso =  CreateObject("Scripting.FileSystemObject")
	If objFso.FileExists(Server.MapPath(strFilePath)) Then
		isExist = True
	Else
		isExist = False
	End If   
	Set objFso=nothing 
	If Err Then Err.clear
	ExistFile = isExist
End Function

'��ȡ�ļ�
Function GetFile(ByVal strFilePath)
	Dim objFso, txtFile, strContent
	Set objFso =  CreateObject("Scripting.FileSystemObject")
	If objFso.FileExists(Server.MapPath(strFilePath)) Then
		Set txtFile = objFso.OpenTextFile(Server.MapPath(strFilePath))
		While Not txtFile.AtEndOfStream
			 strContent = strContent & txtFile.ReadLine & vbCrLf 
		Wend
	Else
		strContent = "<font color='red'>�������ļ�[" & strFilePath & "]�����飡</font>"
	End If   
	Set objFso=nothing 
	If Err Then Err.clear
	GetFile = strContent
End Function

' �����ļ�
Function CreateFile(Byval content,Byval fileDir)
	On Error Resume Next	
	Dim objFso, txtFile
	fileDir = replace(fileDir, "\", "/") : fileDir = replace(fileDir, "//", "/")
	If Right(fileDir, 1) = "/" Then CreateFile = False: Exit Function
	Call CreateFolder(Left(fileDir, InStrRev(fileDir,"/")))	'�Զ������ļ���
	Set objFso =  Server.CreateObject("Scripting.FileSystemObject")
	Set txtFile= objFso.CreateTextFile(Server.MapPath(fileDir),True)
	txtFile.WriteLine content
	txtFile.Close
	Set objFso = Nothing
	If Err Then Err.clear: CreateFile = False else CreateFile = True
End Function


' ɾ���ļ�
Function DeleteFile(byval fileDir)
	On Error Resume Next
	Dim objFso
	If Len(fileDir) = 0 or IsNull(fileDir) Then Exit Function
	fileDir = replace(fileDir, "\", "/") : fileDir = replace(fileDir, "//", "/")
	If right(fileDir, 1) = "/" Then
		DeleteFile = DeleteFolder(fileDir)
	Else
		Set objFso = Server.CreateObject("Scripting.FileSystemObject")
		objFso.DeleteFile Server.Mappath(fileDir)
		Set objFso = Nothing
		If Err Then Err.Clear: DeleteFile = False Else DeleteFile = True
	End If
End Function



'�����ļ���
Function CreateFolder(Byval dirPath)
        On Error Resume Next 
        Dim astrPath, ulngPath, i, strTmpPath , strPath
        Dim objFSO
		strPath = Server.MapPath(dirPath)
        If InStr(strPath, "\") <=0 or InStr(strPath, ":") <= 0 Then 
                CreateFolder = False 
                Exit Function 
        End If
        Set objFSO = Server.CreateObject("Scripting.FileSystemObject") 
        If objFSO.FolderExists(strPath) Then 
                CreateFolder = True 
                Exit Function 
        End If 
        astrPath = Split(strPath, "\") 
        ulngPath = UBound(astrPath) 
        strTmpPath = "" 
        For i = 0 To ulngPath 
                strTmpPath = strTmpPath & astrPath(i) & "\" 
                If Not objFSO.FolderExists(strTmpPath) Then 
                        '���� 
                        objFSO.CreateFolder(strTmpPath) 
                End If 
        Next 
        Set objFSO = Nothing 
        If Err = 0 Then 
                CreateFolder = True 
        Else 
				Err.Clear
                CreateFolder = False 
        End If 
End Function  


' ɾ���ļ���
Function DeleteFolder(byval dirPath)
	On Error Resume Next
	Dim objFso
	dirPath = replace(dirPath, "\", "/") : dirPath = replace(dirPath, "//", "/")
	dirPath = Server.Mappath(dirpath)
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")
	If objFso.FolderExists(dirpath) Then 
		objFso.DeleteFolder(dirpath)
	Else
		Response.Write("�ļ�������")
	End If
	Set objFso = Nothing 
	If Err Then Response.Write(Err.Description): Err.Clear: DeleteFolder = False Else DeleteFolder = True
End Function

'�Ƿ�����ļ���
Function ExistFolder(ByVal strFilePath)
	Dim objFso, isExist
	Set objFso =  CreateObject("Scripting.FileSystemObject")
	If objFso.FolderExists(Server.MapPath(strFilePath)) Then
		isExist = True
	Else
		isExist = False
	End If   
	Set objFso=nothing 
	If Err Then Err.clear
	ExistFolder = isExist
End Function

%>