<!--#include file="include/include.asp"-->
<%	Dim serverUrl1, serverUrl2, strSkinPath, strTempCss
	serverUrl1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
	serverUrl2 = Cstr(Request.ServerVariables("SERVER_NAME"))
	If Mid(serverUrl1, 8, Len(serverUrl2)) <>  serverUrl2 Then
		Response.Write "你想干什么？<a href='http://www.eekku.com/'>E酷Cms</a>"
	Else
		strSkinPath = Replace(TemplatePath & "/" & Trim(Request("path")), "//", "/")
		If Right(LCase(strSkinPath), 4) = ".css" Then	'显示Css中背景图片
			If IsCache = 1 And ChkCache("Css_" & strSkinPath) Then
				strTempCss = GetCache("Css_" & strSkinPath)
			Else
				strTempCss = ReadFile(strSkinPath)
				strTempCss = Replace(strTempCss, "../../../", "")
				strTempCss = Replace(strTempCss, "../../", "template/")
				strTempCss = Replace(strTempCss, "../", "skin.asp?path=")
				If IsCache = 1 Then
					Call SetCache("Css_" & strSkinPath, strTempCss)
				End If
			End If
			Response.Write(strTempCss)
		Else
			Response.Redirect strSkinPath
		End If
	End If
%>
