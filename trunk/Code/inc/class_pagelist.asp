<%
'======================================
' File Name��	ClassPagelist.asp
' Purpose��		��ҳ��
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-7-21 11:08:14
' Updated on: 	2009-9-14 19:26:26(����GetUrl��ַ����������bug)
'======================================
Class ClassPageList
    Dim ID ' ����,Ĭ��Ϊ ID
    Dim Field ' �ֶ�,Ĭ��Ϊ *
    Dim Table ' ���ݱ�,����Ϊ��
    Dim Where ' ����,���ô�
    Dim Order ' ����,Ĭ�ϰ���������
	Dim Sql		'SQL���
    Dim Result ' ��������,0��Getrows,1���ָ��,Ĭ�� 0
    Dim PageSize ' ÿҳ��¼��,Ĭ��15
    Dim AbsolutePage ' ��ǰҳ,Ĭ�� 1 [��һҳ]
    
    Dim Data ' ���ؼ�¼��
    Dim Eof ' �Ƿ��м�¼
    Dim RecordCount ' �ܼ�¼��
    Dim PageCount ' ��ҳ��
	Dim PageUrl		'��ַ
    
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
	
	'�����ҳ����
	Public Function Page()
		Dim tempPages, tempPage, tempPageUrl
		tempPages = PageCount
		tempPage = AbsolutePage
		tempPageUrl = GetURL()
		Page = PageBar(tempPages, tempPage, tempPageUrl)
	End Function
	
	'��ҳ����
	Public Function PageBar(pages, page, PageUrl)
		Dim PageStr
		pages = CLng(pages)
		page = CLng(page)
		If page>1 Then
			PageStr = PageStr & " <a href="""&PageUrl&"page=1"">[��ҳ]</a> "	
			PageStr = PageStr & " <a href="""&PageUrl&"page="&page-1&""">[��һҳ]</a> "
			PageStr = PageStr & " <a href="""&PageUrl&"page=1"">[1]</a> "		
		Else
			PageStr = PageStr & " [<strong>1</strong>] "
		End If
		Dim p, pp, num, n, i
			num = 5 '��ʾ����ҳ��
			i = 0	'�ж��Ƿ�����ѭ������
			p = page - Int(num/2): If p < 2 Then p = 2 '�м���ʼ���� 
			pp = pages - 1	'���ҳ��
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
			PageStr = PageStr & " <a href="""&PageUrl&"page="&page+1&""">[��һҳ]</a>"	
			PageStr = PageStr & " <a href="""&PageUrl&"page="&pages&""">[βҳ]</a>"	
		ElseIf pages > 1 Then		
			PageStr = PageStr & "[<strong>"&pages&"</strong>]"
		End If
		PageBar = PageStr
	End Function
	
	'��ȡ��ǰ����ַ
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
		'������ʽ��ȡ����
		Set Reg = New RegExp
		Reg.Ignorecase = True
		Reg.Global = True
		'Reg.Pattern = "page=.*&{0,1}"
		Reg.Pattern = "page=\d*"
		'Reg.Pattern = "page="
		curUrl = Reg.Replace(curUrl,  "")
		If Instr(curUrl, "?") = 0 Then		'���û�к�׺������ӣ�����ʱtest.asp��Ϊtest.asp?
			curUrl = curUrl & "?"
		ElseIf Instr(curUrl, "=") <> 0 Then
			curUrl = Replace(curUrl, "?&", "?")		'ȥ��test.asp?&aa=1�ж����&
			If Right(curUrl, 1) <> "&" Then			'���test.asp?aa=1���Ϊtest.asp?aa=1&
				curUrl = curUrl & "&"
			End If
		End If
		GetURL = curUrl
	End Function
End Class

%>

