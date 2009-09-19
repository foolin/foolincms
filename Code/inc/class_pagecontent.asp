<%
'======================================
' File Name：	Class_PageContent.asp
' Purpose：		内容分页类
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-7-21 11:08:14
' Updated on: 	2009-7-23 11:39:51(修正GetPageUrl网址带参数出现丢失参数bug)
'======================================
Class ClassPageContent
	Dim vContent	'分页内容
    Dim vCurrPage 	'当前页
    Dim vPageCount 	'分页总数
	Dim vPageSize	'分页内容多大（多少数字）
	Dim vPageUrl	'分页地址
	
	'初始化类
	Private Sub Class_Initialize()
		vContent = ""
		vCurrPage = 0
		vPageCount = 0
		vPageSize = 1000
		vPageUrl = GetURL()
	End Sub
	
	'释放类
	Private Sub Class_Terminate()
	End Sub
    
	'内容
    Public Function OutContent()
        
    End Function
	
	' 自动分页
	function AutosplitPages(byval strNewscontent,byval Page_split_page,byval AutoPagesNum)
		dim i, IsCount, OneChar, strCount, Foundstr, Pages_i_str, Pages_i_Arr
		AutoPagesNum = clng(AutoPagesNum)
		Page_split_page = cstr(Page_split_page)
		if len(strNewscontent) < int(AutoPagesNum + round(AutoPagesNum / 5)) then AutosplitPages = strNewscontent: exit function
		if strNewscontent <> "" and AutoPagesNum <> 0 and instr(1, strNewscontent, Page_split_page) = 0 then
		IsCount = true
		Pages_i_str = ""
		for i = 1 to len(strNewscontent)
			OneChar = Mid(strNewscontent, i, 1)
			if OneChar = "<" then
			IsCount = false
			elseif OneChar = ">" then
			IsCount = true
			else
			if IsCount = true then
				if Abs(Asc(OneChar)) > 255 then
				strCount = strCount + 2
				else
				strCount = strCount + 1
				end if
				if strCount >= AutoPagesNum and i < len(strNewscontent) then
				Foundstr = left(strNewscontent, i)
				if AllowsplitPages(Foundstr, "table|a|b>|i>|strong|div|span") = true then
					Pages_i_str = Pages_i_str & trim(cstr(i)) & ","
					strCount = 0
				end if
				end if
			end if
			end if
		next
		if len(Pages_i_str) > 1 then Pages_i_str = left(Pages_i_str, len(Pages_i_str) - 1)
		Pages_i_Arr = split(Pages_i_str, ",")
		for i = ubound(Pages_i_Arr) to lbound(Pages_i_Arr) Step -1
			strNewscontent = left(strNewscontent, Pages_i_Arr(i)) & Page_split_page & Mid(strNewscontent, Pages_i_Arr(i) + 1)
		next
		end if
		AutosplitPages = strNewscontent
	end function
	
	' 作用：判断是否允许字符串加入分页标记
	function AllowsplitPages(byval Tempstr,byval Findstr)
		dim inti, Beginstr, Endstr, BeginstrNum, EndstrNum, ArrstrFind, i
		Tempstr = lcase(Tempstr)
		Findstr = lcase(Findstr)
		if Tempstr <> "" and Findstr <> "" then
		ArrstrFind = split(Findstr, "|")
		for i = 0 to ubound(ArrstrFind)
			Beginstr = "<" & ArrstrFind(i)
			Endstr = "</" & ArrstrFind(i)
			Inti = 0
			Do While instr(Inti + 1, Tempstr, Beginstr) <> 0
			Inti = instr(Inti + 1, Tempstr, Beginstr)
			BeginstrNum = BeginstrNum + 1
			Loop
			Inti = 0
			Do While instr(Inti + 1, Tempstr, Endstr) <> 0
			Inti = instr(Inti + 1, Tempstr, Endstr)
			EndstrNum = EndstrNum + 1
			Loop
			if EndstrNum = BeginstrNum then
			AllowsplitPages = true
			else
			AllowsplitPages = false
			exit function
			end if
		next
		else
		AllowsplitPages = false
		end if
	end function
	
	'分页代码
	Public Function Page(pages, page, PageUrl)
		Dim pages, page, pageUrl
		Dim PageStr
		pages = CLng(vPageCount)
		page = CLng(vCurrPage)
		pageUrl = GetURL()
		If page>1 Then
			PageStr = PageStr & " <a href="""&PageUrl&"page=1"">[首页]</a> "	
			PageStr = PageStr & " <a href="""&PageUrl&"page="&page-1&""">[上一页]</a> "
			PageStr = PageStr & " <a href="""&PageUrl&"page=1"">[1]</a> "		
		Else
			PageStr = PageStr & " [<strong>1</strong>] "
		End If
		Dim p, pp, num, n, i
			num = 5 '显示几个页数
			i = 0	'判断是否超数字循环变量
			p = page - Int(num/2): If p < 2 Then p = 2 '中间起始数字 
			pp = pages - 1	'最后页数
		If p > 2 Then PageStr = PageStr & "..."      
		For n = p To pp
			i = i + 1            
			If n=page Then
				PageStr = PageStr & " [<strong>"&n&"</strong>] "
			Else
				PageStr = PageStr & " <a href="""&PageUrl&"page="&n&""">["&n&"]</a> " 
			End If       
			If i >= num Then Exit For           
		Next
		If n < pp Then PageStr = PageStr & "..." 
		If page<pages Then
			PageStr = PageStr & "<a href="""&PageUrl&"page="&pages&""">["&pages&"]</a> "	
			PageStr = PageStr & " <a href="""&PageUrl&"page="&page+1&""">[下一页]</a>"	
			PageStr = PageStr & " <a href="""&PageUrl&"page="&pages&""">[尾页]</a>"	
		ElseIf pages > 1 Then		
			PageStr = PageStr & "[<strong>"&pages&"</strong>]"
		End If
		PageBar = PageStr
	End Function

	'获取当前的网址
	Public Function GetURL() 
		Dim strHostName, strScriptName, strSubUrl, strRequestItem 
		Dim curUrl, pageUrl, Reg
		strHostName = CStr(Request.ServerVariables("LOCAL_ADDR"))
		strScriptName = CStr(Request.ServerVariables("SCRIPT_NAME"))
		strSubUrl = ""
		If Request.QueryString <> "" Then
		   strScriptName  =strScriptName & "?"
		   For Each strRequestItem In Request.QueryString
				If InStr(strScriptName,strRequestItem) = 0 Then
					 If strSubUrl = "" Then
						strSubUrl = strSubUrl & strRequestItem & "=" & Server.URLEncode(Request.QueryString(""&strRequestItem&""))
					 Else
					  strSubUrl = strSubUrl & "&" & strRequestItem & "=" & Server.URLEncode(Request.QueryString(""&strRequestItem&""))
					 End If
				End If
		   Next
		End If

		'处理参数
		curUrl = "http://"&strHostName&strScriptName&strSubUrl
		Set Reg = New RegExp
		Reg.Ignorecase = True
		Reg.Global = True
		Reg.Pattern = "page=\d*"
		curUrl = Reg.Replace(curUrl,  "")
		If Instr(curUrl, "?") = 0 Then		'如果没有后缀，则添加？。即时test.asp变为test.asp?
			curUrl = curUrl & "?"
		ElseIf Instr(curUrl, "=") <> 0 Then
			curUrl = Replace(curUrl, "?&", "?")		'去除test.asp?&aa=1中多余的&
			If Right(curUrl, 1) <> "&" Then			'如果test.asp?aa=1则变为test.asp?aa=1&
				curUrl = curUrl & "&"
			End If
		End If
		GetURL = curUrl
	End Function
End Class

%>
