<%
' 读取文件文件
Function ReadFile(ByVal strFilePath)
	Dim objFile, strTempConent
	On Error Resume Next
	Set objFile = Server.CreateObject("adodb.stream")
	With objFile
		.Type = 2: .Mode = 3: .Open: .Charset = "GB2312" : .Position = objFile.Size: .Loadfromfile Server.Mappath(strFilePath): strTempConent = .ReadText: .Close
	End With
	Set objFile = Nothing
	If Err Then  Response.Write Err.Description & Warn("无法加载文件[" & strFilePath & "]"): Response.End
	ReadFile = strTempConent
End Function


'删除文件函数 '2009-4-3
Function DelFile(path)
    set objfso=server.CreateObject("Scripting.FileSystemObject")
    If objfso.fileExists(Server.MapPath(path)) Then
        objfso.Deletefile(Server.MapPath(path))
    End If
    set objfso=Nothing
End Function


'完全过滤Html代码
Function RemoveHTML(strHTML) 
	Dim objRegExp, Match, Matches 
	Set objRegExp = New Regexp
	
	objRegExp.IgnoreCase = True 
	objRegExp.Global = True 
	''取闭合的<> 
	objRegExp.Pattern = "<.+?>" 
	''进行匹配 
	Set Matches = objRegExp.Execute(strHTML)
	
	'' 遍历匹配集合，并替换掉匹配的项目 
	For Each Match in Matches 
	strHtml=Replace(strHTML,Match.Value,"") 
	Next 
	RemoveHTML=strHTML 
	Set objRegExp = Nothing 
End Function

' 过滤字符
Function FilterStr(Byval str)
	FilterStr = LCase(str)
	FilterStr = Replace(FilterStr, " ", "")
	FilterStr = replace(FilterStr, "'", "")
	FilterStr = replace(FilterStr, """", "")
	FilterStr = replace(FilterStr, "=", "")
	FilterStr = replace(FilterStr, "*", "")
End Function


' 创建文件
Function CreateFile(Byval content,Byval fileDir)
	fileDir = replace(fileDir, "\", "/") : fileDir = replace(fileDir, "//", "/")
	If Right(fileDir, 1) = "/" Then fileDir = fileDir & "index." & Defaultext
	call CreateFolder(fileDir)
	On Error Resume Next
	Dim obj : Set obj = server.createobject("adodb.Stream")
	obj.type = 2
	obj.open
	obj.charset = response.charset
	obj.position = obj.Size
	obj.writeText = content
	obj.savetofile server.mappath(fileDir), 2
	obj.close
	If err Then err.clear: createfile = false else createfile = true
	set obj = nothing
end function

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

' 删除文件
function deletefile(byval fileDir)
	If len(fileDir) = 0 or isnull(fileDir) Then exit function
	fileDir = replace(fileDir, "\", "/") : fileDir = replace(fileDir, "//", "/")
	If right(fileDir, 1) = "/" Then
		deletefile = deletefolder(fileDir)
	else
		on error resume next
		fso.deletefile server.mappath(fileDir)
		If err Then err.clear: deletefile = false else deletefile = true
	end If
end function

' 删除文件夹
function deletefolder(byval dirpath)
	on error resume next
	fso.deletefolder server.mappath(dirpath)
	If err Then err.clear: deletefolder = false else deletefolder = true
end function

%>