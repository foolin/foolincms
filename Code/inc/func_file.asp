<%
'=========================================================
' File Name：	func_file.asp
' Purpose：		文件常用操作函数
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-9-13 14:59:09
' Version:		v1.0.0 Build 20090913
' CopyRight (c) 2009 E酷工作室（Foolin）All Rights Reserved
'=========================================================

'是否存在文件
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

'获取文件
Function GetFile(ByVal strFilePath)
	Dim objFso, txtFile, strContent
	Set objFso =  CreateObject("Scripting.FileSystemObject")
	If objFso.FileExists(Server.MapPath(strFilePath)) Then
		Set txtFile = objFso.OpenTextFile(Server.MapPath(strFilePath))
		While Not txtFile.AtEndOfStream
			 strContent = strContent & txtFile.ReadLine & vbCrLf 
		Wend
	Else
		strContent = "<font color='red'>不存在文件[" & strFilePath & "]，请检查！</font>"
	End If   
	Set objFso=nothing 
	If Err Then Err.clear
	GetFile = strContent
End Function

' 创建文件
Function CreateFile(Byval content,Byval fileDir)
	On Error Resume Next	
	Dim objFso, txtFile
	fileDir = replace(fileDir, "\", "/") : fileDir = replace(fileDir, "//", "/")
	If Right(fileDir, 1) = "/" Then CreateFile = False: Exit Function
	Call CreateFolder(Left(fileDir, InStrRev(fileDir,"/")))	'自动创建文件夹
	Set objFso =  Server.CreateObject("Scripting.FileSystemObject")
	Set txtFile= objFso.CreateTextFile(Server.MapPath(fileDir),True)
	txtFile.WriteLine content
	txtFile.Close
	Set objFso = Nothing
	If Err Then Err.clear: CreateFile = False else CreateFile = True
End Function


' 删除文件
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



'创建文件夹
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
                        '创建 
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


' 删除文件夹
Function DeleteFolder(byval dirPath)
	On Error Resume Next
	Dim objFso
	dirPath = replace(dirPath, "\", "/") : dirPath = replace(dirPath, "//", "/")
	dirPath = Server.Mappath(dirpath)
	Set objFso = Server.CreateObject("Scripting.FileSystemObject")
	If objFso.FolderExists(dirpath) Then 
		objFso.DeleteFolder(dirpath)
	Else
		Response.Write("文件不存在")
	End If
	Set objFso = Nothing 
	If Err Then Response.Write(Err.Description): Err.Clear: DeleteFolder = False Else DeleteFolder = True
End Function

'是否存在文件夹
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