<%
'======================================
' File Name：	ClassPagelist.asp
' Purpose：		分页类
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-7-21 11:08:14
' Updated on: 	2009-9-14 19:26:26(修正GetUrl网址带参数出现bug)
'======================================
Class ClassPageList
    Dim ID ' 主键,默认为 ID
    Dim Field ' 字段,默认为 *
    Dim Table ' 数据表,不可为空
    Dim Where ' 条件,不用带
    Dim Order ' 排序,默认按主键排序
	Dim Sql		'SQL语句
    Dim Result ' 返回类型,0则Getrows,1则采指针,默认 0
    Dim PageSize ' 每页记录数,默认15
    Dim AbsolutePage ' 当前页,默认 1 [第一页]
    
    Dim Data ' 返回记录集
    Dim Eof ' 是否有记录
    Dim RecordCount ' 总记录数
    Dim PageCount ' 总页数
	Dim PageUrl		'地址
    
    Public Function List()
        Dim Rs
        If Not IsNumeric(AbsolutePage) Or Len(AbsolutePage) = 0 Then AbsolutePage = 1 Else AbsolutePage = Int(AbsolutePage)
        If AbsolutePage < 1 Then AbsolutePage = 1
        If Len(ID) = 0 Then ID = "[ID]"
        If Len(Field) = 0 Then Field = "*"
        If Len(Where) > 0 Then Where = "Where " & Where
        If Len(Order) > 0 Then Order = "Order By " & Order Else Order = "Order By " & ID & " Desc"
        If Not IsNumeric(PageSize) Or Len(PageSize) = 0 Or PageSize < 1 Then PageSize = 15
        If Not IsNumeric(Result) Or Len(Result) = 0 Or Result = 0 Then Result = 0 Else Result = 1
        PageSize = Int(PageSize)
		If Len(Sql) = 0 Then
			Set Rs = DB("Select " & Field & " From " & Table & " " & Where & " " & Order, 2)
		Else
			Set Rs = DB(Sql, 2)
		End If
		Rs.PageSize = PageSize
		If Not Rs.Eof Then Rs.AbsolutePosition = (AbsolutePage - 1) * PageSize + 1
		RecordCount = Rs.RecordCount: PageCount = Rs.PageCount
        If Result = 0 Then
            If Rs.Eof Then
                Eof = True
            Else
                Eof = False
                If RecordCount < PageSize Then Data = Rs.GetRows(RecordCount) Else Data = Rs.GetRows(PageSize)
                Rs.Close: Set Rs = Nothing
            End If
        Else
            If Rs.Eof Then Eof = True Else Eof = False
            Set Data = Rs
        End If
    End Function
	
	'输出分页代码
	Public Function Page()
		Dim tempPages, tempPage, tempPageUrl
		tempPages = PageCount
		tempPage = AbsolutePage
		tempPageUrl = GetURL()
		Page = PageBar(tempPages, tempPage, tempPageUrl)
	End Function
	
	'分页函数
	Public Function PageBar(pages, page, PageUrl)
		Dim PageStr
		pages = CLng(pages)
		page = CLng(page)
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
	Private Function GetURL() 
		Dim rqUrl, curUrl, Reg
		Dim strScriptName, strSubUrl, strRequestItem
		strScriptName=CStr(Request.ServerVariables("SCRIPT_NAME"))
		strSubUrl=""
		If Request.QueryString<>"" Then
		   strScriptName=strScriptName&"?"
		   For Each strRequestItem In Request.QueryString
			If InStr(strScriptName,strRequestItem)=0 Then
			 If strSubUrl="" Then
			  strSubUrl=strSubUrl&strRequestItem&"="&Server.URLEncode(Request.QueryString(""&strRequestItem&""))
			 Else
			  strSubUrl=strSubUrl&"&"&strRequestItem&"="&Server.URLEncode(Request.QueryString(""&strRequestItem&""))
			 End If
			End If
		   Next
		End If
		rqUrl = strScriptName & strSubUrl
		curUrl = Right(rqUrl, Len(rqUrl) - InstrRev(rqUrl,"/"))
		'正则表达式获取参数
		Set Reg = New RegExp
		Reg.Ignorecase = True
		Reg.Global = True
		'Reg.Pattern = "page=.*&{0,1}"
		Reg.Pattern = "page=\d*"
		'Reg.Pattern = "page="
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

