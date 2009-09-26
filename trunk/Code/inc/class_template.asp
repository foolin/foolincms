<%
'=========================================================
' Class Name：	ClassTemplate
' Purpose：		模板类，处理标签解析执行
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-7-17 11:08:14
' Updated on: 	2009-9-25 15:50:56
' Modify log:	  自定义页面，增加URL参数入口，可以通过URL或者Id进行浏览自定义页面。(2009-9-25 15:50:56)		
'=========================================================

Class ClassTemplate
	
	'类成员
	Private mReg		'正则表达式对象
	Private mContent	'内容
	Private mTemplate	'模板
	Private mCurrPage	'当前页数
	
	'设置当前页
	Public Property Let Page(ByVal cPage) 
		mCurrPage = CInt(cPage)
	End Property
	'取当前页
	Public Property Get Page
		Page = mCurrPage
	End Property
	
	'取内容
	Public Property Get Content
		Content = mContent
	End Property
	
	
	'初始化类
	Private Sub Class_Initialize()
		Set mReg = New RegExp
		mReg.Ignorecase = True
		mReg.Global = True
		mContent = ""
		mTemplate = "" ' 模板路径
		mCurrPage = 1
	End Sub
	
	'释放类
	Private Sub Class_Terminate()
		Set mReg = Nothing
	End Sub
	
	
	'--------------------------------------------------------------
	' Function name：	Load()
	' Description: 		载入模板方式一（要写路径）
	' Params: 			templtateFile - 载入的模板路径
	' Create on: 		2009-7-17 18:23:45
	' Notice:			
	'--------------------------------------------------------------
	Public Function Load(ByVal tplFile)
		mTemplate = tplFile
		If IsCache = 1 Then
			If ChkCache("Template_" & Server.Mappath(mTemplate)) Then
				mContent = GetCache("Template_" & Server.Mappath(mTemplate))
			Else
				Call LoadTemplate()
				Call SetCache("Template_" & Server.Mappath(mTemplate), mContent)
			End If
		Else
			Call LoadTemplate()
		End If
	End Function
	
	'--------------------------------------------------------------
	' Function name：	LoadTpl()
	' Description: 		载入模板方式二，直接写模板名称即可
	' Params: 			templtateFile - 载入的模板文件名
	' Create on: 		2009-7-17 18:23:45
	' Notice:			
	'--------------------------------------------------------------
	Public Function LoadTpl(ByVal tplFile)
		mTemplate = TemplatePath & "/" & tplFile
		If IsCache = 1 Then
			If ChkCache("Template_" & Server.Mappath(mTemplate)) Then
				mContent = GetCache("Template_" & Server.Mappath(mTemplate))
			Else
				Call LoadTemplate()
				Call SetCache("Template_" & Server.Mappath(mTemplate), mContent)
			End If
		Else
			Call LoadTemplate()
		End If
	End Function
	
	
	'--------------------------------------------------------------
	' Function name：	Compile_Index()	
	' Purpose: 			首页调用标签执行，注意执行顺序
	' Author:			Foolin
	' Create on: 		2009-7-31 19:07:19
	' Params:	
	' Return:							
	'--------------------------------------------------------------
	Public Function Compile_Index()
		Call Parser_Include(3)	'执行包含文件分析处理
		Call Parser_MyTag()		'执行自定义标签分析处理
		Call Parser_Sys()		'执行系统标签分析处理
		Call Parser_List(0)		'执行List列表标签分析
		Call Parser_IF()		'调用If标签分析
	End Function
	
	'--------------------------------------------------------------
	' Function name：	Compile_List()	
	' Purpose: 			图片（文章）列表调用标签执行，注意执行顺序
	' Author:			Foolin
	' Create on: 		2009-7-31 19:12:02
	' Params:			ColId - 栏目名称
	' Return:							
	'--------------------------------------------------------------
	Public Function Compile_List(ByVal ColId)
		Call Parser_Include(3)	'执行包含文件分析处理
		Call Parser_MyTag()		'执行自定义标签分析处理
		Call Parser_Sys()		'执行系统标签分析处理
		Call Parser_List(ColId)		'执行List列表标签分析
		Call Parser_IF()	
	End Function
	
	'--------------------------------------------------------------
	' Function name：	Compile_Field()	
	' Purpose: 			文章（图片）调用标签执行，注意执行顺序
	' Author:			Foolin
	' Create on: 		2009-7-18 21:23:58
	' Params:			id - 图片或者文章id, blnIsPic --- 布尔值：True -- 图片，False -- 文章
	' Return:			
	' Modify log:					
	'--------------------------------------------------------------
	Public Function Compile_Field(ByVal id, ByVal blnIsPic)
		Call Parser_Include(3)	'执行包含文件分析处理
		Call Parser_MyTag()		'执行自定义标签分析处理
		Call Parser_Sys()		'执行系统标签分析处理
		Call Parser_List(-1)		'执行List列表标签分析
		Call Parser_Field(id, blnIsPic)	'执行Field标签标签分析
		Call Parser_IF()	
	End Function
	
	'--------------------------------------------------------------
	' Function name：	Compile_Field()	
	' Purpose: 			文章（图片）调用标签执行，注意执行顺序
	' Author:			Foolin
	' Create on: 		2009-7-31 19:12:08
	' Params:			param - 页面参数
	' Return:			
	' Modify log:					
	'--------------------------------------------------------------
	Public Function Compile_DiyPage(ByVal param)
		Call Parser_DiyPage(param)		'执行DiyPage标签分析
		Call Parser_Include(3)	'执行包含文件分析处理
		Call Parser_MyTag()		'执行自定义标签分析处理
		Call Parser_Sys()		'执行系统标签分析处理
		Call Parser_List(-1)		'执行List列表标签分析
		Call Parser_IF()	
	End Function
	
	'--------------------------------------------------------------
	' Function name：	Compile_Plugin()	
	' Purpose: 			插件调用标签执行，注意执行顺序
	' Author:			Foolin
	' Create on: 		2009-8-7
	' Params:	
	' Return:							
	'--------------------------------------------------------------
	Public Function Compile_Plugin()
		Call Parser_Include(3)	'执行包含文件分析处理
		Call Parser_MyTag()		'执行自定义标签分析处理
		Call Parser_Sys()		'执行系统标签分析处理
		Call Parser_List(0)		'执行List列表标签分析
		Call Parser_IF()		'调用If标签分析
	End Function
	

	'--------------------------------------------------------------
	' Function name：	Parser_Include()	
	' Purpose: 			分析包含标签{include file="文件url"}
	' Author:			Foolin
	' Create on: 		2009-7-18 21:23:58
	' Params:			nLayer - 嵌套层数；防止死循环嵌套，最大不能超过三层
	' Return:			none
	' Modify log:		
	' Notice:						
	'--------------------------------------------------------------
	Public Function Parser_Include(ByVal nLayer)
		On Error Resume Next
		If Cint(nLayer) <= 0 Then
			Exit Function
		ElseIf Cint(nLayer) > 3 Then
			Response.Write Warn("{include}指令最多只能嵌套3层！"): Response.End
		End If
		'执行正则表达式进行匹配
		Dim Matches, Match
		Dim incFilePath, incContent, tempIncPath, strAttrs
		mReg.Pattern = "\{include(.+?)\}"
		Set Matches = mReg.Execute(mContent)
		For Each Match In Matches
			If Len(Replace(Match.SubMatches(0), " ", "")) > 0 Then
				strAttrs = Match.SubMatches(0)
				tempIncPath = TemplatePath & "/"  & GetAttrValue(strAttrs, "file", False) '加上模板目录路径
				tempIncPath = Replace(tempIncPath, "//", "/")	'如果有双斜杠//，则替换成单斜杠
				'判断缓存，提高效率
				If IsCache = 1 Then
					If ChkCache("Template_" & Server.Mappath(tempIncPath)) Then
						incContent = GetCache("Template_" & Server.Mappath(tempIncPath))
					Else
						incContent = LoadFile(tempIncPath)
						Call SetCache("Template_" & Server.Mappath(tempIncPath), incContent)	'缓存
					End If
				Else
					incContent = LoadFile(tempIncPath)
				End If
			Else 
				incContent = ""
			End If
			mContent = Replace(mContent, Match.Value, incContent) ' 替换
			If Err Then Err.Clear: Response.Write Err.Description & Warn("{include}格式不合法，请检查！"): Response.End
		Next
		If RegExists("\{include(.+?)\}", mContent) Then Call Parser_Include(nLayer - 1)	'递归调用
	End Function
	

	'--------------------------------------------------------------
	' Function name：	Parser_DiyPage()	
	' Purpose: 			自定义页面（DiyPage）标签:{diypage:字段名 属性=值}
	' Author:			Foolin
	' Create on: 		2009-7-24 18:20:26
	' Params:			param - 页面参数
	' Return:			none
	' Update on:		2009-9-25 16:01:19
	' Modify log:		增加文件名参数类型（即是可以通过id或者文件名作为参数）
	' Notice:						
	'--------------------------------------------------------------
	Public Function Parser_DiyPage(ByVal param)
		On Error Resume Next
		Dim Matches, Match
		Dim objRs, strSql
		Dim tagName, strAttrs, tagLen, tagLenExt
		If Len(param) = 0 Then Response.Write(Warn("参数错误！")): Response.End()
		If IsNumeric(param) Then
			strSql = "SELECT * FROM DiyPage WHERE State = 1 AND ID =  " & param
		Else
			strSql = "SELECT * FROM DiyPage WHERE State = 1 AND PageName = '" & param &"'"
		End If
		Set objRs = DB(strSql, 1)
		If objRs.Eof Then Response.Write(Warn("不存在[" & param & "]页面。请在后台[Diy页面]管理进行添加该页面！")): Response.End()
		If Len(Trim(objRs("Template"))) > 0 Then
			mTemplate = TemplatePath & "/" & objRs("Template")
		Else
			mTemplate = TemplatePath & "/diypage.html"
		End If
		mTemplate = Replace(Replace(mTemplate, "///", "/"), "//", "/")	'过滤三斜杠///和双斜杠//
		If IsCache = 1 Then
			If ChkCache("Template_" & Server.Mappath(mTemplate)) Then
				mContent = GetCache("Template_" & Server.Mappath(mTemplate))
			Else
				Call LoadTemplate()
				Call SetCache("Template_" & Server.Mappath(mTemplate), mContent)
			End If
		Else
			Call LoadTemplate()
		End If
		mReg.Pattern = "\{diypage\s*:\s*([^\s\}/]*)(.+?)?\}"
		Set Matches = mReg.Execute(mContent)
		For Each Match In Matches
			tagName = Trim(Match.SubMatches(0))
			strAttrs = Match.SubMatches(1)
			tagLen = Int(GetAttrValue(strAttrs, "len", True))
			tagLenExt = GetAttrValue(strAttrs, "lenext", False)
			If tagLen > 0 And Len(objRs(tagName)) > tagLen Then
				mContent = Replace(mContent, Match.Value, Left(objRs(tagName), tagLen) & tagLenExt) ' 替换
			Else
				mContent = Replace(mContent, Match.Value, objRs(tagName)) ' 替换
			End If
			If Err Then Err.Clear: mContent = Replace(mContent, Match.Value, Warn(Match.Value))
			tagLen = 0: tagLenExt = ""	'清除
		Next
	End Function


	'--------------------------------------------------------------
	' Function name：	Parser_MyTag()	
	' Purpose: 			分析自定义标签:{my:字段名 /}
	' Author:			Foolin
	' Create on: 		2009-7-20 20:20:26
	' Params:			none
	' Return:			none
	' Modify log:		
	' Notice:						
	'--------------------------------------------------------------
	Public Function Parser_MyTag()
		On Error Resume Next
		Dim Matches, Match, tagName, tagValue
		mReg.Pattern = "\{my\s*:\s*([\s\S]*?)\s*(?:/{0,1})\}"
		Set Matches = mReg.Execute(mContent)
		For Each Match In Matches
			tagName = Replace(Match.SubMatches(0), " ", "")
			If Len(tagName) > 0 Then
				tagValue = GetMyTag(tagName)
			Else
				tagValue = ""
			End If
			mContent = Replace(mContent, Match.Value, tagValue) ' 替换
			If Err Then Response.Write Warn("{my}格式不合法，请检查！" & Err.Number & Err.Source  & Err.Description & Err ): Err.Clear: Response.End
		Next
	End Function
	

	'--------------------------------------------------------------
	' Function name：	Parser_Sys()
	' Purpose: 			分析系统标签:{sys:tagname /}
	' Author:			Foolin
	' Create on: 		2009-7-20 20:20:26
	' Params:			none
	' Return:			none
	' Modify log:		
	' Notice:						
	'--------------------------------------------------------------
	Public Function Parser_Sys()
		On Error Resume Next
		Dim Matches, Match
		Dim strName, strValue

		mReg.Pattern = "\{sys\s*:\s*([^\s\}/]*)(.+?)?\}"
		Set Matches = mReg.Execute(mContent)
		For Each Match In Matches
			strName = Replace(Match.SubMatches(0), " ", "")
			If Len(strName) > 0 Then
				Execute ("strValue = " & Replace(Match.SubMatches(0), " ", ""))
			Else
				strValue = ""
			End If
			mContent = Replace(mContent, Match.Value, strValue) ' 替换
			If Err Then Err.Clear: Response.Write Warn("{sys}格式不合法，请检查！"): Response.End
		Next
		'替换css、html里面的图片
		mReg.pattern = "<(.*?)(src=|href=|value=)""(images/|css/|js/|scripts/)(.*?)""(.*?)>"
		If IsHideTempPath = 1 Then
			mContent = mReg.replace(mcontent, "<$1$2""skin.asp?path=$3$4""$5>")
		Else
			mContent = mReg.replace(mcontent, "<$1$2""" & TemplatePath & "/$3$4""$5>")
		End If
		mReg.pattern = "url\((.*?)\)"	'替换网页中css中背景图片
		If IsHideTempPath = 1 Then
			mContent = mReg.replace(mcontent, "url(skin.asp?path=$1)")
		Else
			mContent = mReg.replace(mcontent, "url(" & TemplatePath & "/$1)")
		End If
		
	End Function
		
	
	'--------------------------------------------------------------
	' Function name：	Parser_List()	
	' Purpose: 			分析列表标签：{list:my名称 属性=值}{/list:my名称}
	' Author:			Foolin
	' Create on: 		2009-7-21 18:26:31
	' Params:			none
	' Return:			none
	' Modify log:		增加ColId参数，增加列表Column="auto"功能。
	' Notice:						
	'--------------------------------------------------------------
	Public Function Parser_List(ByVal ColId)
		On Error Resume Next
		Dim Matches, Match, strListName, strAtrr, strInnerText
		Dim tagMode, tagCache, tagRow, tagCol, tagClass, tagWidth, tagIsPage	'公共属性
		Dim tagTable, tagField, tagWhere, tagOrder, tagSQL	'组合SQL模式属性
		Dim tagSrc, tagColumn	'缺省模式属性
		Dim objRs, strTempValue
		Dim i, j
		'srcType作用:替换标签时，处理特殊标签，例如titleurl,colname,colurl只对article和picture两个表有效。
		Dim srcType: srcType = "other"
		
		mReg.Pattern = "\{list:([\S^\}]*)(.+?)\}([\s\S]*?)\{/list:\1\}"
		Set Matches = mReg.Execute(mContent)
		For Each Match In Matches
			strListName = Match.SubMatches(0)		'标签名称
			strAtrr = Match.SubMatches(1)		'标签属性
			strInnerText  = Match.SubMatches(2) '内层标签
			'公共属性
			tagMode = GetAttrValue(strAtrr, "mode", False)		'模式，default|table|sql。默认为default
			tagCache = GetAttrValue(strAtrr, "cache", True)		'缓存时间,默认为0
			tagRow = GetAttrValue(strAtrr, "row", True)			'行数，默认为10行
			tagCol = GetAttrValue(strAtrr, "col", True)			'列数，默认为1列
			tagClass = GetAttrValue(strAtrr, "class", False)	'表格样式，class样式
			tagWidth = GetAttrValue(strAtrr, "width", True)		'表格宽度，默认为100%
			tagIsPage = GetAttrValue(strAtrr, "ispage", false)	'是否分页，默认为false
			If Len(tagMode) = 0 Then tagMode = "default"
			If LCase(Trim(tagMode))<>"table" And LCase(Trim(tagMode))<>"sql" Then tagMode = "default"
			If Len(tagCache) = 0 Or Not IsNumeric(tagCache) Then tagCache = -1 ' 标签不用缓存
			If Len(tagRow) = 0 Or Not IsNumeric(tagRow) Then tagRow = 10
			If Int(tagRow) < 1 Then tagRow = 1
			If Len(tagCol) = 0 Or Not IsNumeric(tagCol) Then tagCol = 1
			If Int(tagCol) < 1 Then tagCol = 1
			If Len(tagWidth) = 0 Then tagWidth = "100%"
			If Len(tagClass) > 0 Then tagClass = " class=""" & tagClass & """ "
			If Len(tagIsPage) > 0 And (Trim(tagIsPage) = "true" Or Trim(tagIsPage) = "1") Then tagIsPage = true Else tagIsPage = false
			tagCache = Int(tagCache): tagRow = Int(tagRow): tagCol = Int(tagCol)

			'mode="table"组合SQL语句属性
			tagTable = GetAttrValue(strAtrr, "table", False)			'数据库表
			tagField = GetAttrValue(strAtrr, "field", False)		'字段
			tagWhere = GetAttrValue(strAtrr, "where", False)		'符合条件
			tagOrder = GetAttrValue(strAtrr, "order", False)		'排序条件
			If Len(tagTable) = 0 Then tagTable = "Article"			'默认设为文章列表
			If Len(tagField) = 0 Then tagField = "*"				'默认全部字段
			
			'纯SQL语句属性
			tagSQL = GetAttrValue(strAtrr, "sql", False)
			
			'当mode="table"模式时
			If tagMode = "table" Then
				tagSQL = "SELECT " & tagField & " FROM " & tagTable
				If Len(tagWhere) > 0 Then tagSQL = tagSQL & " WHERE " & tagWhere
				If Len(tagOrder) > 0 Then tagSQL = tagSQL & " ORDER BY " & tagOrder 
			ElseIf tagMode <> "sql" Then	'mode缺省模式属性
				tagSrc = GetAttrValue(strAtrr, "src", False)
				tagColumn = GetAttrValue(strAtrr, "column", False)
				If Len(tagSrc)= 0 Then tagSrc = "article"			'默认设为文章列表
				If Len(tagWhere) > 0 Then 
					tagWhere = tagWhere & " AND State = 1 "
				Else
					tagWhere = " State = 1 "
				End If
				If Len(tagColumn) > 0 Then
					If LCase(tagColumn) <> "auto" Then
						tagWhere = tagWhere & " AND ColID IN (" & tagColumn & ") "
					ElseIf  LCase(tagColumn) = "auto" And ColId > 0 Then
					 	tagWhere = tagWhere & " AND ColID IN (" & ColId & ") "
					End If
				End If
				Select Case LCase(tagSrc)
					Case "image", "pic", "picture"
						tagTable = "Picture"
					Case "imgart", "picart"
						tagTable = "Article"
						tagWhere = tagWhere & " AND IsFocusPic = 1 AND FocusPic<>'' "
					Case Else
						tagTable = "Article"
				End Select
				Select Case LCase(tagOrder)
					Case "hot"
						tagOrder = " Hits DESC"
					Case "asc"
						tagOrder = " ID"
					Case "last","desc"
						tagOrder = " ID DESC"
					Case Else
						tagOrder = " ID DESC"
				End Select
				tagSQL = "SELECT " & tagField & " FROM " & tagTable & " WHERE " & tagWhere & " ORDER BY " & tagOrder
			End If
			
			'如果缺省tagSQL值，则为默认值。
			If Len(tagSQL) = 0 Then tagSQL = "SELECT * FORM Article Where State = 1 ORDER BY ID DESC"
			
			'判断是否是article|picture表，以便处理特殊标签
			If InStr(LCase(tagSQL), "colid") > 0 Or InStr(LCase(tagSQL), "*") > 0 Then
				If InStr(LCase(tagSQL), "article") > 0 Then
					srcType = "article"
				ElseIf InStr(LCase(tagSQL), "picture") > 0 Then
					srcType = "picture"
				End If
			End If
			
			If ChkCache(mTemplate & tagSQL) Then
				strTempValue = GetCache(mTemplate & tagSQL)
			Else
				strTempValue = ""
				'是否为分页
				If tagIsPage = True Then
					Set objRs = New ClassPageList
					objRs.Result = 1
					objRs.Sql = tagSQL
					objRs.PageSize = tagRow * tagCol 
					objRs.AbsolutePage = mCurrPage
					objRs.List()
				Else
					Set objRs = DB(tagSQL, 2)
				End If
				If Err Then Response.Write Warn("模板中{list}标签属性SQL出错[" & tagSQL & "] <br /><br />错误描述 : " & Err.Description): Response.End
				'如果tagCol＞1则表格形式输出
				If tagCol > 1 Then strTempValue = strTempValue & "<table width=""" & tagWidth & """ " & tagClass & ">" & vbCrLf

				Session(CacheFlag & "List_i")  = 0
				If tagIsPage = True Then	'判断是否分页
					Session(CacheFlag & "List_num") = objRs.Data.RecordCount 
				Else
					Session(CacheFlag & "List_num") = objRs.RecordCount 
				End If
				If Session(CacheFlag & "List_num") > tagRow * tagCol Then Session(CacheFlag & "List_num") = tagRow * tagCol
				j = 0	'判断col变量
				For i = 1 To tagRow * tagCol	'循环输出记录
					If tagIsPage = True Then	'判断是否分页
						If objRs.Data.Eof Then Exit For	
					Else
						If objRs.Eof Then Exit For	'没有记录，则退出
					End If
					j = j + 1
					Session(CacheFlag & "List_i")  = Session(CacheFlag & "List_i") + 1
					If tagCol > 1 Then ' 表
						If j = 1 Then strTempValue = strTempValue & "  <tr>" & vbCrLf
						strTempValue = strTempValue & "	<td valign=""top"" width=""" & Round(100 / tagCol) & "%"">"
					End If
					'替换标签
					If tagIsPage = True Then	'判断是否分页
						strTempValue = strTempValue & ReplaceListTags(strListName, strInnerText, objRs.Data, srcType)
					Else
						strTempValue = strTempValue & ReplaceListTags(strListName, strInnerText, objRs, srcType)
					End If

						
					If tagCol > 1 Then ' 表
						strTempValue = strTempValue & "	</td>" & vbCrLf
						If j = tagCol Then strTempValue = strTempValue & "  </tr>" & vbCrLf: j = 0
					End If
					If tagIsPage = True Then	'判断是否分页
						objRs.Data.MoveNext
					Else
						objRs.MoveNext
					End If
				Next
				
				If tagCol > 1 Then
					If j < tagCol And j > 0 Then
						For i = 1 To tagCol - j
							strTempValue = strTempValue & "	<td width=""" & Round(100 / tagCol) & "%""></td>" & vbCrLf
						Next
						strTempValue = strTempValue & "  </tr>" & vbCrLf
					End If
					strTempValue = strTempValue & "</table>" & vbCrLf
				End If
				
				'替换分页{tag:page /}
				If tagIsPage = True Then
					mContent = RegReplace(mContent, "\{tag:page\s*/\}", objRs.Page)
				End If

				If tagIsPage = True Then	'判断是否分页
					objRs.Data.Close
					Set objRs = Nothing
				Else
					objRs.Close: Set objRs = Nothing
				End If

			End If
			mContent = Replace(mContent, Match.Value, strTempValue) ' 替换
			If Err Then Response.Write Err.Description & "<br />" : Err.Clear: Response.Write  Warn("{list}格式不合法，请检查！"): Response.End
		Next
		' 多次调用，列表嵌套
		If RegExists("\{list:([\S^\}]*)(.+?)\}([\s\S]*?)\{/list:\1\}", mContent) Then Call Parser_List(ColId)
	End Function
	

	'--------------------------------------------------------------
	' Function name：	Parser_Field()	
	' Purpose: 			分析单篇文章或者图片{field:字段名 /}{art:字段名/}{pic:picture/}
	' Author:			Foolin
	' Create on: 		2009-7-23 20:26:31
	' Params: 		 	id -- 文章(或者图片)记录id
	'				 	blnIsPic --- 布尔值：True -- 图片，False -- 文章
	' Return:			none
	' Modify log:		
	' Notice:						
	'--------------------------------------------------------------
	Public Function Parser_Field(ByVal id, ByVal blnIsPic)
		On Error Resume Next
		Dim objRs, strSql
		If Len(id) = 0 or Not IsNumeric(id) Then Response.Write  Warn("id错误，请检查！"): Response.End
		If blnIsPic Then
			strSql = "SELECT * FROM Picture WHERE State = 1 AND ID = " & id
		Else
			strSql = "SELECT * FROM Article WHERE State = 1 AND ID = " & id
		End If
		Set objRs = DB(strSql, 1)
		If objRs.Eof Then Response.Write  Warn("不存在id[" & id & "]记录！"): Response.End
		'更新点击率
		If blnIsPic Then
			Call DB("UPDATE Picture SET Hits = Hits + 1 WHERE ID = " & objRs("ID"), 1)
		Else
			Call DB("UPDATE Article SET Hits = Hits + 1 WHERE ID = " & objRs("ID"), 1)
		End If
		'判断是否存在{field}
		Call ReplaceFieldTags(objRs, "field", blnIsPic)
		Call ReplaceFieldTags(objRs, "tag", blnIsPic)
		If blnIsPic Then
			'替换字段
			Call ReplacePreNextTags(objRs, "tag", blnIsPic)
			Call ReplaceFieldTags(objRs, "pic", blnIsPic)
			Call ReplaceFieldTags(objRs, "picture", blnIsPic)
			Call ReplaceFieldTags(objRs, "img", blnIsPic)
			Call ReplaceFieldTags(objRs, "image", blnIsPic)
		Else
			Call ReplaceFieldTags(objRs, "art", blnIsPic)
			Call ReplaceFieldTags(objRs, "article", blnIsPic)
		End If
		
		objRs.Close: Set objRs = Nothing
	End Function
	
	'--------------------------------------------------------------
	' Function name：	Parser_IF()
	' Purpose: 			替换List:Tag标签签，Parser_List()调用此函数
	' Author:			Foolin
	' Create on: 		2009-7-21 20:26:31
	' Params:			none
	' Return:			none				
	'--------------------------------------------------------------
	Public Function Parser_IF()
		On Error Resume Next
		Dim Matches, Match
		Dim TestIF
		mReg.Pattern = "{If:(.+?)}([\s\S]*?){Else}([\s\S]*?){/If}"
		Set Matches = mReg.Execute(mContent)
		For Each Match In Matches
			Execute ("If " & Match.SubMatches(0) & " Then TestIf = True Else TestIf = False")
			If TestIF Then mContent = Replace(mContent, Match.Value, Match.SubMatches(1)) Else mContent = Replace(mContent, Match.Value, Match.SubMatches(2)) ' 替换
			If Err Then Response.Write Warn("{IF}标签出错[" & Match.SubMatches(0) & "]" & Err.Description): Err.Clear: Response.End
		Next
		mReg.Pattern = "{If:(.+?)}([\s\S]*?){/If}"
		Set Matches = mReg.Execute(Content)
		For Each Match In Matches
			Execute ("If " & Match.SubMatches(0) & " Then TestIf = True Else TestIf = False")
			If TestIF Then mContent = Replace(mContent, Match.Value, Match.SubMatches(1)) Else mContent = Replace(mContent, Match.Value, "") ' 替换
			If Err Then Response.Write Warn("{IF}标签出错[" & Match.SubMatches(0) & "]" & Err.Description): Err.Clear: Response.End
		Next
	End Function


	'--------------------------------------------------------------
	' Function name：	ReplaceListTags()
	' Purpose: 			替换List:Tag标签签，Parser_List()调用此函数
	' Author:			Foolin
	' Create on: 		2009-7-21 20:26:31
	' Params:			strListName -- 列表名称
	'				 	strTemp - 待处理文本
	'				 	objRs - 数据集
	'					srcType - 类型（只对Article和Picture有效）
	' Return:			处理替换过后的标签值
	' Modify log:		
	' Notice:						
	'--------------------------------------------------------------
	Private Function ReplaceListTags(ByVal strListName, ByVal strTemp, ByVal objRs, ByVal srcType)
		On Error Resume Next
		Dim Matches, Match, tagName, tagAttrs
		Dim attrLen, attrLenExt, attrFormat, attrClearHtml
		Dim clearHtmlValue
		'mReg.Pattern = "\[\s*list:(\S+)(\s[^\]]*)?\]"	'匹配[list:field]或者[list:field len="" lenext=""]
		mReg.Pattern = "\[\s*" & strListName & "\s*:(.+?)\]"
		Set Matches = mReg.Execute(strTemp)
		For Each Match In Matches
			tagName = Trim(Replace(Match.SubMatches(0), "	", " ")): tagName = Split(tagName, " ")(0)
			'If Len(tagName) = 0 Then Exit For	'标签名称
			tagAttrs = Trim(Match.SubMatches(0)) 	'标签属性
			Select Case LCase(tagName)
			Case "url"
				If srcType = "article" Then
					strTemp = Replace(strTemp, Match.Value, "article.asp?id=" & objRs("ID"))
				ElseIf srcType = "picture" Then
					strTemp = Replace(strTemp, Match.Value, "picture.asp?id=" & objRs("ID"))
				Else
					strTemp = Replace(strTemp, Match.Value, Warn(Match.Value))
				End If
			Case "colname"
				If srcType = "article" Then
					strTemp = Replace(strTemp, Match.Value, GetColName(objRs("ColID"), "article"))
				ElseIf srcType = "picture" Then
					strTemp = Replace(strTemp, Match.Value, GetColName(objRs("ColID"), "picture"))
				Else
					strTemp = Replace(strTemp, Match.Value, Warn(Match.Value))
				End If
			Case "colurl"
				If srcType = "article" Then
					strTemp = Replace(strTemp, Match.Value, "artlist.asp?id=" & objRs("ColID"))
				ElseIf srcType = "picture" Then
					strTemp = Replace(strTemp, Match.Value, "piclist.asp?id=" & objRs("ColID"))
				Else
					strTemp = Replace(strTemp, Match.Value, Warn(Match.Value))
				End If
			Case "i"	'i为输出序号
				strTemp = Replace(strTemp, Match.Value, Session(CacheFlag & "List_i"))
			Case "num"	'记录总数
				strTemp = Replace(strTemp, Match.Value, Session(CacheFlag & "List_num"))
			Case "field"	'字段名：格式[listname:field name=""]
				If Len(tagAttrs) > 0 Then tagName = GetAttrValue(tagAttrs, "name", False)
				If Len(tagName) > 0 Then
					'检查是否带有格式化时间属性
					attrFormat = GetAttrValue(tagAttrs, "format", False)
					If Len(attrFormat) > 0 And IsDate(objRs(tagName)) Then
						strTemp = Replace(strTemp, Match.Value, FormatTime(objRs(tagName), attrFormat))
					End If
					'是否去除HTML代码
					attrClearHtml = Trim(GetAttrValue(tagAttrs, "clearhtml", False))
					If LCase(attrClearHtml) = "true" Then
						clearHtmlValue = ClearHtml(objRs(tagName))
					Else
						clearHtmlValue = objRs(tagName)
					End If
					If Err Then
						Err.Clear : clearHtmlValue = Warn(Match.Value)
					Else
						'检查是否存在截取字符属性
						attrLen = Int(GetAttrValue(tagAttrs, "len", True))
						If Len(attrLen) > 0 And attrLen > 0 And Len(clearHtmlValue) > attrLen Then
							attrLenExt = GetAttrValue(tagAttrs, "lenext", False)
							clearHtmlValue = Left(clearHtmlValue, attrLen) & attrLenExt
						End If
					End If
					strTemp = Replace(strTemp, Match.Value, clearHtmlValue)
				End If
			Case Else	'字段名
				'检查是否带有格式化时间属性
				attrFormat = GetAttrValue(tagAttrs, "format", False)
				If Len(attrFormat) > 0 And IsDate(objRs(tagName)) Then
					strTemp = Replace(strTemp, Match.Value, FormatTime(objRs(tagName), attrFormat))
				End If

				'是否去除HTML代码
				attrClearHtml = Trim(GetAttrValue(tagAttrs, "clearhtml", False))
				If Len(attrClearHtml) > 0 And LCase(attrClearHtml) = "true" Then
					clearHtmlValue = ClearHtml(objRs(tagName))
				Else
					clearHtmlValue = objRs(tagName)
				End If
				If Err Then
					Err.Clear : clearHtmlValue = Warn(Match.Value)
				Else
					'检查是否存在截取字符属性
					If Len(tagAttrs) > 0 Then attrLen = Int(GetAttrValue(tagAttrs, "len", True))
					If Len(attrLen) > 0 And attrLen > 0 And Len(clearHtmlValue) > attrLen Then
						attrLenExt = GetAttrValue(tagAttrs, "lenext", False)
						clearHtmlValue = Left(clearHtmlValue, attrLen) & attrLenExt
					End If
				End If
				strTemp = Replace(strTemp, Match.Value, clearHtmlValue)
			End Select
			
			'字段值为空，则将标签替换为空
			If Len(CStr(objRs(tagName))) = 0 Then
				strTemp = Replace(strTemp, Match.Value, Replace(VarType(objRs(tagName)), "1", ""))
			End If

			'不存在字段，则警告
			If Err Then  Err.Clear : strTemp = Replace(strTemp, Match.Value, Warn(Match.Value))
			'清除attrLen和attrLenExt
			tagAttrs = "": attrLen = 0: attrLenExt = "": attrClearHtml = "": clearHtmlValue = ""
		Next
		ReplaceListTags = strTemp
	End Function
	

	'--------------------------------------------------------------
	' Function name：	ReplaceFieldTags()	
	' Purpose: 			替换{field:Tag}标签, Paser_Field()函数会调用此函数
	' Author:			Foolin
	' Create on: 		2009-7-23 20:26:31
	' Params: 		 	objRs-数据集
	'				 	fieldName - 标签名称
	'				 	blnIsPic - 是否为图片。True - 图片， False - 文章
	' Return:			none
	' Modify log:		
	' Notice:						
	'--------------------------------------------------------------
	Private Function ReplaceFieldTags(ByVal objRs, ByVal fieldName, ByVal blnIsPic)
		On Error Resume Next
		Dim Matches, Match, pattern
		Dim intFlagType, tagName, attrTypeValue, strTemp
		pattern = "\{" & fieldName & "\s*:(.+?)/?\}"
		If Not RegExists(pattern, mContent) Then Exit Function
		'作为GetPreLink或者GetNextLink函数入口参数
		If blnIsPic Then
			intFlagType = 1
		Else
			intFlagType = 0
		End If
		mReg.Pattern = pattern
		'If Err Then Response.Write "Stop:" & Err.Description: Response.End()
		Set Matches = mReg.Execute(mContent)
		For Each Match In Matches
			'取标签名称（pre|next）
			tagName = Trim(Replace(Match.SubMatches(0), "	", " ")): tagName = Split(tagName, " ")(0)
			'If Len(tagName) = 0 Then Exit For	'标签名称不存在则退出
			attrTypeValue = GetAttrValue(Trim(Match.SubMatches(0)), "type", False) 	'type属性值
			
			If LCase(tagName) = "pre" Or  LCase(tagName) = "previous" Then
				
				Select Case LCase(Trim(attrTypeValue))
				Case "id"
					strTemp = GetPreLink(objRs("id"), intFlagType, 0)
				Case "title"
					strTemp = GetPreLink(objRs("id"), intFlagType, 1)
				Case "url"
					strTemp = GetPreLink(objRs("id"), intFlagType, 2)
				Case "link"
					strTemp = GetPreLink(objRs("id"), intFlagType, 3)
				Case Else
					strTemp = GetPreLink(objRs("id"), intFlagType, 3)
				End Select
				
			ElseIf LCase(tagName) = "next" Then
			
				Select Case LCase(Trim(attrTypeValue))
				Case "id"
					strTemp = GetNextLink(objRs("id"), intFlagType, 0)
				Case "title"
					strTemp = GetNextLink(objRs("id"), intFlagType, 1)
				Case "url"
					strTemp = GetNextLink(objRs("id"), intFlagType, 2)
				Case "link"
					strTemp = GetNextLink(objRs("id"), intFlagType, 3)
				Case Else
					strTemp = GetNextLink(objRs("id"), intFlagType, 3)
				End Select
			ElseIf LCase(fieldName) <> "tag" And Len(fieldName) > 0 Then
				strTemp = objRs(tagName)
				If Err Then Err.Clear : strTemp = Warn(Match.Value)	'如果不存在记录
			End If

			mContent = Replace(mContent, Match.Value, strTemp) ' 替换
		Next
	End Function
	

	'--------------------------------------------------------------
	' Function name：	GetAttrValue()	
	' Purpose: 			获取标签属性的值
	' Author:			Foolin
	' Create on: 		2009-7-23 20:26:31
	' Params: 		 	strTags - 全部标签属性
	'					strAttrName - 标签属性名称
	'					blnIsNum - 返回属性值是否为数字
	' Return:			返回属性值
	' Modify log:		
	' Notice:						
	'--------------------------------------------------------------
	Private Function GetAttrValue(ByVal strTags, ByVal strAttrName, ByVal blnIsNum)
		Dim tagValue: tagValue = ""
		If Len(strTags) <= 3 Then GetAttrValue = "": Exit Function
		Dim Matches, Match
		mReg.Pattern = strAttrName & "\s*=\s*""([\s\S]*?)"""
		Set Matches = mReg.Execute(strTags)
		For Each Match In Matches
			tagValue = Trim(Match.SubMatches(0))
			'If Err Then Err.Clear: Response.Write Err.Description & Warn("{dblist1}格式不合法，请检查！"): Response.End
		Next
		If blnIsNum Then
			tagValue = Replace(tagValue, " ", "")
			If Len(tagValue) > 0 And IsNumeric(tagValue) And InStr(tagValue, ",") = 0 Then tagValue = Int(tagValue)
		End If
		GetAttrValue = tagValue
	End Function

	' 载入模板
	Private Function LoadTemplate()
		Dim Obj
		On Error Resume Next
		Set Obj = Server.CreateObject("adodb.stream")
		With Obj
			.Type = 2: .Mode = 3: .Open: .Charset = "GB2312" : .Position = Obj.Size: .Loadfromfile Server.Mappath(mTemplate): mContent = .ReadText: .Close
		End With
		Set Obj = Nothing
		If Err Then Response.Write Err.Description & Warn("无法加载模板[" & mTemplate & "]"):Response.End
	End Function
	
	' 载入文件
	Private Function LoadFile(ByVal strFilePath)
		Dim objFile, strTempConent
		On Error Resume Next
		Set objFile = Server.CreateObject("adodb.stream")
		With objFile
			.Type = 2: .Mode = 3: .Open: .Charset = "GB2312" : .Position = objFile.Size: .Loadfromfile Server.Mappath(strFilePath): strTempConent = .ReadText: .Close
		End With
		Set objFile = Nothing
		If Err Then  Response.Write Err.Description & Warn("无法加载文件[" & strFilePath & "]"): Response.End
		LoadFile = strTempConent
	End Function

	
	' 是否存在此类标签
	Private Function RegExists(ByVal pattern, ByVal strContent)
		mReg.Pattern = pattern
		RegExists = mReg.Test(strContent)
	End Function
	
	' 正表达式替换
	Private Function RegReplace(ByVal repContent, ByVal pattern, ByVal repValue)
		mReg.Pattern = pattern
		RegReplace = mReg.Replace(repContent, repValue)
	End Function
	
	
	'替换函数
	Private Function Rep(strSource, strDestn)
		mContent = Replace(mContent, strSource, strDestn)
	End Function

End Class
%>