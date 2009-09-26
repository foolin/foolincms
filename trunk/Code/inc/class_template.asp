<%
'=========================================================
' Class Name��	ClassTemplate
' Purpose��		ģ���࣬�����ǩ����ִ��
' Auhtor: 		Foolin
' E-mail: 		Foolin@126.com
' Createed on: 	2009-7-17 11:08:14
' Updated on: 	2009-9-25 15:50:56
' Modify log:	  �Զ���ҳ�棬����URL������ڣ�����ͨ��URL����Id��������Զ���ҳ�档(2009-9-25 15:50:56)		
'=========================================================

Class ClassTemplate
	
	'���Ա
	Private mReg		'������ʽ����
	Private mContent	'����
	Private mTemplate	'ģ��
	Private mCurrPage	'��ǰҳ��
	
	'���õ�ǰҳ
	Public Property Let Page(ByVal cPage) 
		mCurrPage = CInt(cPage)
	End Property
	'ȡ��ǰҳ
	Public Property Get Page
		Page = mCurrPage
	End Property
	
	'ȡ����
	Public Property Get Content
		Content = mContent
	End Property
	
	
	'��ʼ����
	Private Sub Class_Initialize()
		Set mReg = New RegExp
		mReg.Ignorecase = True
		mReg.Global = True
		mContent = ""
		mTemplate = "" ' ģ��·��
		mCurrPage = 1
	End Sub
	
	'�ͷ���
	Private Sub Class_Terminate()
		Set mReg = Nothing
	End Sub
	
	
	'--------------------------------------------------------------
	' Function name��	Load()
	' Description: 		����ģ�巽ʽһ��Ҫд·����
	' Params: 			templtateFile - �����ģ��·��
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
	' Function name��	LoadTpl()
	' Description: 		����ģ�巽ʽ����ֱ��дģ�����Ƽ���
	' Params: 			templtateFile - �����ģ���ļ���
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
	' Function name��	Compile_Index()	
	' Purpose: 			��ҳ���ñ�ǩִ�У�ע��ִ��˳��
	' Author:			Foolin
	' Create on: 		2009-7-31 19:07:19
	' Params:	
	' Return:							
	'--------------------------------------------------------------
	Public Function Compile_Index()
		Call Parser_Include(3)	'ִ�а����ļ���������
		Call Parser_MyTag()		'ִ���Զ����ǩ��������
		Call Parser_Sys()		'ִ��ϵͳ��ǩ��������
		Call Parser_List(0)		'ִ��List�б��ǩ����
		Call Parser_IF()		'����If��ǩ����
	End Function
	
	'--------------------------------------------------------------
	' Function name��	Compile_List()	
	' Purpose: 			ͼƬ�����£��б���ñ�ǩִ�У�ע��ִ��˳��
	' Author:			Foolin
	' Create on: 		2009-7-31 19:12:02
	' Params:			ColId - ��Ŀ����
	' Return:							
	'--------------------------------------------------------------
	Public Function Compile_List(ByVal ColId)
		Call Parser_Include(3)	'ִ�а����ļ���������
		Call Parser_MyTag()		'ִ���Զ����ǩ��������
		Call Parser_Sys()		'ִ��ϵͳ��ǩ��������
		Call Parser_List(ColId)		'ִ��List�б��ǩ����
		Call Parser_IF()	
	End Function
	
	'--------------------------------------------------------------
	' Function name��	Compile_Field()	
	' Purpose: 			���£�ͼƬ�����ñ�ǩִ�У�ע��ִ��˳��
	' Author:			Foolin
	' Create on: 		2009-7-18 21:23:58
	' Params:			id - ͼƬ��������id, blnIsPic --- ����ֵ��True -- ͼƬ��False -- ����
	' Return:			
	' Modify log:					
	'--------------------------------------------------------------
	Public Function Compile_Field(ByVal id, ByVal blnIsPic)
		Call Parser_Include(3)	'ִ�а����ļ���������
		Call Parser_MyTag()		'ִ���Զ����ǩ��������
		Call Parser_Sys()		'ִ��ϵͳ��ǩ��������
		Call Parser_List(-1)		'ִ��List�б��ǩ����
		Call Parser_Field(id, blnIsPic)	'ִ��Field��ǩ��ǩ����
		Call Parser_IF()	
	End Function
	
	'--------------------------------------------------------------
	' Function name��	Compile_Field()	
	' Purpose: 			���£�ͼƬ�����ñ�ǩִ�У�ע��ִ��˳��
	' Author:			Foolin
	' Create on: 		2009-7-31 19:12:08
	' Params:			param - ҳ�����
	' Return:			
	' Modify log:					
	'--------------------------------------------------------------
	Public Function Compile_DiyPage(ByVal param)
		Call Parser_DiyPage(param)		'ִ��DiyPage��ǩ����
		Call Parser_Include(3)	'ִ�а����ļ���������
		Call Parser_MyTag()		'ִ���Զ����ǩ��������
		Call Parser_Sys()		'ִ��ϵͳ��ǩ��������
		Call Parser_List(-1)		'ִ��List�б��ǩ����
		Call Parser_IF()	
	End Function
	
	'--------------------------------------------------------------
	' Function name��	Compile_Plugin()	
	' Purpose: 			������ñ�ǩִ�У�ע��ִ��˳��
	' Author:			Foolin
	' Create on: 		2009-8-7
	' Params:	
	' Return:							
	'--------------------------------------------------------------
	Public Function Compile_Plugin()
		Call Parser_Include(3)	'ִ�а����ļ���������
		Call Parser_MyTag()		'ִ���Զ����ǩ��������
		Call Parser_Sys()		'ִ��ϵͳ��ǩ��������
		Call Parser_List(0)		'ִ��List�б��ǩ����
		Call Parser_IF()		'����If��ǩ����
	End Function
	

	'--------------------------------------------------------------
	' Function name��	Parser_Include()	
	' Purpose: 			����������ǩ{include file="�ļ�url"}
	' Author:			Foolin
	' Create on: 		2009-7-18 21:23:58
	' Params:			nLayer - Ƕ�ײ�������ֹ��ѭ��Ƕ�ף�����ܳ�������
	' Return:			none
	' Modify log:		
	' Notice:						
	'--------------------------------------------------------------
	Public Function Parser_Include(ByVal nLayer)
		On Error Resume Next
		If Cint(nLayer) <= 0 Then
			Exit Function
		ElseIf Cint(nLayer) > 3 Then
			Response.Write Warn("{include}ָ�����ֻ��Ƕ��3�㣡"): Response.End
		End If
		'ִ��������ʽ����ƥ��
		Dim Matches, Match
		Dim incFilePath, incContent, tempIncPath, strAttrs
		mReg.Pattern = "\{include(.+?)\}"
		Set Matches = mReg.Execute(mContent)
		For Each Match In Matches
			If Len(Replace(Match.SubMatches(0), " ", "")) > 0 Then
				strAttrs = Match.SubMatches(0)
				tempIncPath = TemplatePath & "/"  & GetAttrValue(strAttrs, "file", False) '����ģ��Ŀ¼·��
				tempIncPath = Replace(tempIncPath, "//", "/")	'�����˫б��//�����滻�ɵ�б��
				'�жϻ��棬���Ч��
				If IsCache = 1 Then
					If ChkCache("Template_" & Server.Mappath(tempIncPath)) Then
						incContent = GetCache("Template_" & Server.Mappath(tempIncPath))
					Else
						incContent = LoadFile(tempIncPath)
						Call SetCache("Template_" & Server.Mappath(tempIncPath), incContent)	'����
					End If
				Else
					incContent = LoadFile(tempIncPath)
				End If
			Else 
				incContent = ""
			End If
			mContent = Replace(mContent, Match.Value, incContent) ' �滻
			If Err Then Err.Clear: Response.Write Err.Description & Warn("{include}��ʽ���Ϸ������飡"): Response.End
		Next
		If RegExists("\{include(.+?)\}", mContent) Then Call Parser_Include(nLayer - 1)	'�ݹ����
	End Function
	

	'--------------------------------------------------------------
	' Function name��	Parser_DiyPage()	
	' Purpose: 			�Զ���ҳ�棨DiyPage����ǩ:{diypage:�ֶ��� ����=ֵ}
	' Author:			Foolin
	' Create on: 		2009-7-24 18:20:26
	' Params:			param - ҳ�����
	' Return:			none
	' Update on:		2009-9-25 16:01:19
	' Modify log:		�����ļ����������ͣ����ǿ���ͨ��id�����ļ�����Ϊ������
	' Notice:						
	'--------------------------------------------------------------
	Public Function Parser_DiyPage(ByVal param)
		On Error Resume Next
		Dim Matches, Match
		Dim objRs, strSql
		Dim tagName, strAttrs, tagLen, tagLenExt
		If Len(param) = 0 Then Response.Write(Warn("��������")): Response.End()
		If IsNumeric(param) Then
			strSql = "SELECT * FROM DiyPage WHERE State = 1 AND ID =  " & param
		Else
			strSql = "SELECT * FROM DiyPage WHERE State = 1 AND PageName = '" & param &"'"
		End If
		Set objRs = DB(strSql, 1)
		If objRs.Eof Then Response.Write(Warn("������[" & param & "]ҳ�档���ں�̨[Diyҳ��]���������Ӹ�ҳ�棡")): Response.End()
		If Len(Trim(objRs("Template"))) > 0 Then
			mTemplate = TemplatePath & "/" & objRs("Template")
		Else
			mTemplate = TemplatePath & "/diypage.html"
		End If
		mTemplate = Replace(Replace(mTemplate, "///", "/"), "//", "/")	'������б��///��˫б��//
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
				mContent = Replace(mContent, Match.Value, Left(objRs(tagName), tagLen) & tagLenExt) ' �滻
			Else
				mContent = Replace(mContent, Match.Value, objRs(tagName)) ' �滻
			End If
			If Err Then Err.Clear: mContent = Replace(mContent, Match.Value, Warn(Match.Value))
			tagLen = 0: tagLenExt = ""	'���
		Next
	End Function


	'--------------------------------------------------------------
	' Function name��	Parser_MyTag()	
	' Purpose: 			�����Զ����ǩ:{my:�ֶ��� /}
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
			mContent = Replace(mContent, Match.Value, tagValue) ' �滻
			If Err Then Response.Write Warn("{my}��ʽ���Ϸ������飡" & Err.Number & Err.Source  & Err.Description & Err ): Err.Clear: Response.End
		Next
	End Function
	

	'--------------------------------------------------------------
	' Function name��	Parser_Sys()
	' Purpose: 			����ϵͳ��ǩ:{sys:tagname /}
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
			mContent = Replace(mContent, Match.Value, strValue) ' �滻
			If Err Then Err.Clear: Response.Write Warn("{sys}��ʽ���Ϸ������飡"): Response.End
		Next
		'�滻css��html�����ͼƬ
		mReg.pattern = "<(.*?)(src=|href=|value=)""(images/|css/|js/|scripts/)(.*?)""(.*?)>"
		If IsHideTempPath = 1 Then
			mContent = mReg.replace(mcontent, "<$1$2""skin.asp?path=$3$4""$5>")
		Else
			mContent = mReg.replace(mcontent, "<$1$2""" & TemplatePath & "/$3$4""$5>")
		End If
		mReg.pattern = "url\((.*?)\)"	'�滻��ҳ��css�б���ͼƬ
		If IsHideTempPath = 1 Then
			mContent = mReg.replace(mcontent, "url(skin.asp?path=$1)")
		Else
			mContent = mReg.replace(mcontent, "url(" & TemplatePath & "/$1)")
		End If
		
	End Function
		
	
	'--------------------------------------------------------------
	' Function name��	Parser_List()	
	' Purpose: 			�����б��ǩ��{list:my���� ����=ֵ}{/list:my����}
	' Author:			Foolin
	' Create on: 		2009-7-21 18:26:31
	' Params:			none
	' Return:			none
	' Modify log:		����ColId�����������б�Column="auto"���ܡ�
	' Notice:						
	'--------------------------------------------------------------
	Public Function Parser_List(ByVal ColId)
		On Error Resume Next
		Dim Matches, Match, strListName, strAtrr, strInnerText
		Dim tagMode, tagCache, tagRow, tagCol, tagClass, tagWidth, tagIsPage	'��������
		Dim tagTable, tagField, tagWhere, tagOrder, tagSQL	'���SQLģʽ����
		Dim tagSrc, tagColumn	'ȱʡģʽ����
		Dim objRs, strTempValue
		Dim i, j
		'srcType����:�滻��ǩʱ�����������ǩ������titleurl,colname,colurlֻ��article��picture��������Ч��
		Dim srcType: srcType = "other"
		
		mReg.Pattern = "\{list:([\S^\}]*)(.+?)\}([\s\S]*?)\{/list:\1\}"
		Set Matches = mReg.Execute(mContent)
		For Each Match In Matches
			strListName = Match.SubMatches(0)		'��ǩ����
			strAtrr = Match.SubMatches(1)		'��ǩ����
			strInnerText  = Match.SubMatches(2) '�ڲ��ǩ
			'��������
			tagMode = GetAttrValue(strAtrr, "mode", False)		'ģʽ��default|table|sql��Ĭ��Ϊdefault
			tagCache = GetAttrValue(strAtrr, "cache", True)		'����ʱ��,Ĭ��Ϊ0
			tagRow = GetAttrValue(strAtrr, "row", True)			'������Ĭ��Ϊ10��
			tagCol = GetAttrValue(strAtrr, "col", True)			'������Ĭ��Ϊ1��
			tagClass = GetAttrValue(strAtrr, "class", False)	'�����ʽ��class��ʽ
			tagWidth = GetAttrValue(strAtrr, "width", True)		'����ȣ�Ĭ��Ϊ100%
			tagIsPage = GetAttrValue(strAtrr, "ispage", false)	'�Ƿ��ҳ��Ĭ��Ϊfalse
			If Len(tagMode) = 0 Then tagMode = "default"
			If LCase(Trim(tagMode))<>"table" And LCase(Trim(tagMode))<>"sql" Then tagMode = "default"
			If Len(tagCache) = 0 Or Not IsNumeric(tagCache) Then tagCache = -1 ' ��ǩ���û���
			If Len(tagRow) = 0 Or Not IsNumeric(tagRow) Then tagRow = 10
			If Int(tagRow) < 1 Then tagRow = 1
			If Len(tagCol) = 0 Or Not IsNumeric(tagCol) Then tagCol = 1
			If Int(tagCol) < 1 Then tagCol = 1
			If Len(tagWidth) = 0 Then tagWidth = "100%"
			If Len(tagClass) > 0 Then tagClass = " class=""" & tagClass & """ "
			If Len(tagIsPage) > 0 And (Trim(tagIsPage) = "true" Or Trim(tagIsPage) = "1") Then tagIsPage = true Else tagIsPage = false
			tagCache = Int(tagCache): tagRow = Int(tagRow): tagCol = Int(tagCol)

			'mode="table"���SQL�������
			tagTable = GetAttrValue(strAtrr, "table", False)			'���ݿ��
			tagField = GetAttrValue(strAtrr, "field", False)		'�ֶ�
			tagWhere = GetAttrValue(strAtrr, "where", False)		'��������
			tagOrder = GetAttrValue(strAtrr, "order", False)		'��������
			If Len(tagTable) = 0 Then tagTable = "Article"			'Ĭ����Ϊ�����б�
			If Len(tagField) = 0 Then tagField = "*"				'Ĭ��ȫ���ֶ�
			
			'��SQL�������
			tagSQL = GetAttrValue(strAtrr, "sql", False)
			
			'��mode="table"ģʽʱ
			If tagMode = "table" Then
				tagSQL = "SELECT " & tagField & " FROM " & tagTable
				If Len(tagWhere) > 0 Then tagSQL = tagSQL & " WHERE " & tagWhere
				If Len(tagOrder) > 0 Then tagSQL = tagSQL & " ORDER BY " & tagOrder 
			ElseIf tagMode <> "sql" Then	'modeȱʡģʽ����
				tagSrc = GetAttrValue(strAtrr, "src", False)
				tagColumn = GetAttrValue(strAtrr, "column", False)
				If Len(tagSrc)= 0 Then tagSrc = "article"			'Ĭ����Ϊ�����б�
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
			
			'���ȱʡtagSQLֵ����ΪĬ��ֵ��
			If Len(tagSQL) = 0 Then tagSQL = "SELECT * FORM Article Where State = 1 ORDER BY ID DESC"
			
			'�ж��Ƿ���article|picture���Ա㴦�������ǩ
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
				'�Ƿ�Ϊ��ҳ
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
				If Err Then Response.Write Warn("ģ����{list}��ǩ����SQL����[" & tagSQL & "] <br /><br />�������� : " & Err.Description): Response.End
				'���tagCol��1������ʽ���
				If tagCol > 1 Then strTempValue = strTempValue & "<table width=""" & tagWidth & """ " & tagClass & ">" & vbCrLf

				Session(CacheFlag & "List_i")  = 0
				If tagIsPage = True Then	'�ж��Ƿ��ҳ
					Session(CacheFlag & "List_num") = objRs.Data.RecordCount 
				Else
					Session(CacheFlag & "List_num") = objRs.RecordCount 
				End If
				If Session(CacheFlag & "List_num") > tagRow * tagCol Then Session(CacheFlag & "List_num") = tagRow * tagCol
				j = 0	'�ж�col����
				For i = 1 To tagRow * tagCol	'ѭ�������¼
					If tagIsPage = True Then	'�ж��Ƿ��ҳ
						If objRs.Data.Eof Then Exit For	
					Else
						If objRs.Eof Then Exit For	'û�м�¼�����˳�
					End If
					j = j + 1
					Session(CacheFlag & "List_i")  = Session(CacheFlag & "List_i") + 1
					If tagCol > 1 Then ' ��
						If j = 1 Then strTempValue = strTempValue & "  <tr>" & vbCrLf
						strTempValue = strTempValue & "	<td valign=""top"" width=""" & Round(100 / tagCol) & "%"">"
					End If
					'�滻��ǩ
					If tagIsPage = True Then	'�ж��Ƿ��ҳ
						strTempValue = strTempValue & ReplaceListTags(strListName, strInnerText, objRs.Data, srcType)
					Else
						strTempValue = strTempValue & ReplaceListTags(strListName, strInnerText, objRs, srcType)
					End If

						
					If tagCol > 1 Then ' ��
						strTempValue = strTempValue & "	</td>" & vbCrLf
						If j = tagCol Then strTempValue = strTempValue & "  </tr>" & vbCrLf: j = 0
					End If
					If tagIsPage = True Then	'�ж��Ƿ��ҳ
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
				
				'�滻��ҳ{tag:page /}
				If tagIsPage = True Then
					mContent = RegReplace(mContent, "\{tag:page\s*/\}", objRs.Page)
				End If

				If tagIsPage = True Then	'�ж��Ƿ��ҳ
					objRs.Data.Close
					Set objRs = Nothing
				Else
					objRs.Close: Set objRs = Nothing
				End If

			End If
			mContent = Replace(mContent, Match.Value, strTempValue) ' �滻
			If Err Then Response.Write Err.Description & "<br />" : Err.Clear: Response.Write  Warn("{list}��ʽ���Ϸ������飡"): Response.End
		Next
		' ��ε��ã��б�Ƕ��
		If RegExists("\{list:([\S^\}]*)(.+?)\}([\s\S]*?)\{/list:\1\}", mContent) Then Call Parser_List(ColId)
	End Function
	

	'--------------------------------------------------------------
	' Function name��	Parser_Field()	
	' Purpose: 			������ƪ���»���ͼƬ{field:�ֶ��� /}{art:�ֶ���/}{pic:picture/}
	' Author:			Foolin
	' Create on: 		2009-7-23 20:26:31
	' Params: 		 	id -- ����(����ͼƬ)��¼id
	'				 	blnIsPic --- ����ֵ��True -- ͼƬ��False -- ����
	' Return:			none
	' Modify log:		
	' Notice:						
	'--------------------------------------------------------------
	Public Function Parser_Field(ByVal id, ByVal blnIsPic)
		On Error Resume Next
		Dim objRs, strSql
		If Len(id) = 0 or Not IsNumeric(id) Then Response.Write  Warn("id�������飡"): Response.End
		If blnIsPic Then
			strSql = "SELECT * FROM Picture WHERE State = 1 AND ID = " & id
		Else
			strSql = "SELECT * FROM Article WHERE State = 1 AND ID = " & id
		End If
		Set objRs = DB(strSql, 1)
		If objRs.Eof Then Response.Write  Warn("������id[" & id & "]��¼��"): Response.End
		'���µ����
		If blnIsPic Then
			Call DB("UPDATE Picture SET Hits = Hits + 1 WHERE ID = " & objRs("ID"), 1)
		Else
			Call DB("UPDATE Article SET Hits = Hits + 1 WHERE ID = " & objRs("ID"), 1)
		End If
		'�ж��Ƿ����{field}
		Call ReplaceFieldTags(objRs, "field", blnIsPic)
		Call ReplaceFieldTags(objRs, "tag", blnIsPic)
		If blnIsPic Then
			'�滻�ֶ�
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
	' Function name��	Parser_IF()
	' Purpose: 			�滻List:Tag��ǩǩ��Parser_List()���ô˺���
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
			If TestIF Then mContent = Replace(mContent, Match.Value, Match.SubMatches(1)) Else mContent = Replace(mContent, Match.Value, Match.SubMatches(2)) ' �滻
			If Err Then Response.Write Warn("{IF}��ǩ����[" & Match.SubMatches(0) & "]" & Err.Description): Err.Clear: Response.End
		Next
		mReg.Pattern = "{If:(.+?)}([\s\S]*?){/If}"
		Set Matches = mReg.Execute(Content)
		For Each Match In Matches
			Execute ("If " & Match.SubMatches(0) & " Then TestIf = True Else TestIf = False")
			If TestIF Then mContent = Replace(mContent, Match.Value, Match.SubMatches(1)) Else mContent = Replace(mContent, Match.Value, "") ' �滻
			If Err Then Response.Write Warn("{IF}��ǩ����[" & Match.SubMatches(0) & "]" & Err.Description): Err.Clear: Response.End
		Next
	End Function


	'--------------------------------------------------------------
	' Function name��	ReplaceListTags()
	' Purpose: 			�滻List:Tag��ǩǩ��Parser_List()���ô˺���
	' Author:			Foolin
	' Create on: 		2009-7-21 20:26:31
	' Params:			strListName -- �б�����
	'				 	strTemp - �������ı�
	'				 	objRs - ���ݼ�
	'					srcType - ���ͣ�ֻ��Article��Picture��Ч��
	' Return:			�����滻����ı�ǩֵ
	' Modify log:		
	' Notice:						
	'--------------------------------------------------------------
	Private Function ReplaceListTags(ByVal strListName, ByVal strTemp, ByVal objRs, ByVal srcType)
		On Error Resume Next
		Dim Matches, Match, tagName, tagAttrs
		Dim attrLen, attrLenExt, attrFormat, attrClearHtml
		Dim clearHtmlValue
		'mReg.Pattern = "\[\s*list:(\S+)(\s[^\]]*)?\]"	'ƥ��[list:field]����[list:field len="" lenext=""]
		mReg.Pattern = "\[\s*" & strListName & "\s*:(.+?)\]"
		Set Matches = mReg.Execute(strTemp)
		For Each Match In Matches
			tagName = Trim(Replace(Match.SubMatches(0), "	", " ")): tagName = Split(tagName, " ")(0)
			'If Len(tagName) = 0 Then Exit For	'��ǩ����
			tagAttrs = Trim(Match.SubMatches(0)) 	'��ǩ����
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
			Case "i"	'iΪ������
				strTemp = Replace(strTemp, Match.Value, Session(CacheFlag & "List_i"))
			Case "num"	'��¼����
				strTemp = Replace(strTemp, Match.Value, Session(CacheFlag & "List_num"))
			Case "field"	'�ֶ�������ʽ[listname:field name=""]
				If Len(tagAttrs) > 0 Then tagName = GetAttrValue(tagAttrs, "name", False)
				If Len(tagName) > 0 Then
					'����Ƿ���и�ʽ��ʱ������
					attrFormat = GetAttrValue(tagAttrs, "format", False)
					If Len(attrFormat) > 0 And IsDate(objRs(tagName)) Then
						strTemp = Replace(strTemp, Match.Value, FormatTime(objRs(tagName), attrFormat))
					End If
					'�Ƿ�ȥ��HTML����
					attrClearHtml = Trim(GetAttrValue(tagAttrs, "clearhtml", False))
					If LCase(attrClearHtml) = "true" Then
						clearHtmlValue = ClearHtml(objRs(tagName))
					Else
						clearHtmlValue = objRs(tagName)
					End If
					If Err Then
						Err.Clear : clearHtmlValue = Warn(Match.Value)
					Else
						'����Ƿ���ڽ�ȡ�ַ�����
						attrLen = Int(GetAttrValue(tagAttrs, "len", True))
						If Len(attrLen) > 0 And attrLen > 0 And Len(clearHtmlValue) > attrLen Then
							attrLenExt = GetAttrValue(tagAttrs, "lenext", False)
							clearHtmlValue = Left(clearHtmlValue, attrLen) & attrLenExt
						End If
					End If
					strTemp = Replace(strTemp, Match.Value, clearHtmlValue)
				End If
			Case Else	'�ֶ���
				'����Ƿ���и�ʽ��ʱ������
				attrFormat = GetAttrValue(tagAttrs, "format", False)
				If Len(attrFormat) > 0 And IsDate(objRs(tagName)) Then
					strTemp = Replace(strTemp, Match.Value, FormatTime(objRs(tagName), attrFormat))
				End If

				'�Ƿ�ȥ��HTML����
				attrClearHtml = Trim(GetAttrValue(tagAttrs, "clearhtml", False))
				If Len(attrClearHtml) > 0 And LCase(attrClearHtml) = "true" Then
					clearHtmlValue = ClearHtml(objRs(tagName))
				Else
					clearHtmlValue = objRs(tagName)
				End If
				If Err Then
					Err.Clear : clearHtmlValue = Warn(Match.Value)
				Else
					'����Ƿ���ڽ�ȡ�ַ�����
					If Len(tagAttrs) > 0 Then attrLen = Int(GetAttrValue(tagAttrs, "len", True))
					If Len(attrLen) > 0 And attrLen > 0 And Len(clearHtmlValue) > attrLen Then
						attrLenExt = GetAttrValue(tagAttrs, "lenext", False)
						clearHtmlValue = Left(clearHtmlValue, attrLen) & attrLenExt
					End If
				End If
				strTemp = Replace(strTemp, Match.Value, clearHtmlValue)
			End Select
			
			'�ֶ�ֵΪ�գ��򽫱�ǩ�滻Ϊ��
			If Len(CStr(objRs(tagName))) = 0 Then
				strTemp = Replace(strTemp, Match.Value, Replace(VarType(objRs(tagName)), "1", ""))
			End If

			'�������ֶΣ��򾯸�
			If Err Then  Err.Clear : strTemp = Replace(strTemp, Match.Value, Warn(Match.Value))
			'���attrLen��attrLenExt
			tagAttrs = "": attrLen = 0: attrLenExt = "": attrClearHtml = "": clearHtmlValue = ""
		Next
		ReplaceListTags = strTemp
	End Function
	

	'--------------------------------------------------------------
	' Function name��	ReplaceFieldTags()	
	' Purpose: 			�滻{field:Tag}��ǩ, Paser_Field()��������ô˺���
	' Author:			Foolin
	' Create on: 		2009-7-23 20:26:31
	' Params: 		 	objRs-���ݼ�
	'				 	fieldName - ��ǩ����
	'				 	blnIsPic - �Ƿ�ΪͼƬ��True - ͼƬ�� False - ����
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
		'��ΪGetPreLink����GetNextLink������ڲ���
		If blnIsPic Then
			intFlagType = 1
		Else
			intFlagType = 0
		End If
		mReg.Pattern = pattern
		'If Err Then Response.Write "Stop:" & Err.Description: Response.End()
		Set Matches = mReg.Execute(mContent)
		For Each Match In Matches
			'ȡ��ǩ���ƣ�pre|next��
			tagName = Trim(Replace(Match.SubMatches(0), "	", " ")): tagName = Split(tagName, " ")(0)
			'If Len(tagName) = 0 Then Exit For	'��ǩ���Ʋ��������˳�
			attrTypeValue = GetAttrValue(Trim(Match.SubMatches(0)), "type", False) 	'type����ֵ
			
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
				If Err Then Err.Clear : strTemp = Warn(Match.Value)	'��������ڼ�¼
			End If

			mContent = Replace(mContent, Match.Value, strTemp) ' �滻
		Next
	End Function
	

	'--------------------------------------------------------------
	' Function name��	GetAttrValue()	
	' Purpose: 			��ȡ��ǩ���Ե�ֵ
	' Author:			Foolin
	' Create on: 		2009-7-23 20:26:31
	' Params: 		 	strTags - ȫ����ǩ����
	'					strAttrName - ��ǩ��������
	'					blnIsNum - ��������ֵ�Ƿ�Ϊ����
	' Return:			��������ֵ
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
			'If Err Then Err.Clear: Response.Write Err.Description & Warn("{dblist1}��ʽ���Ϸ������飡"): Response.End
		Next
		If blnIsNum Then
			tagValue = Replace(tagValue, " ", "")
			If Len(tagValue) > 0 And IsNumeric(tagValue) And InStr(tagValue, ",") = 0 Then tagValue = Int(tagValue)
		End If
		GetAttrValue = tagValue
	End Function

	' ����ģ��
	Private Function LoadTemplate()
		Dim Obj
		On Error Resume Next
		Set Obj = Server.CreateObject("adodb.stream")
		With Obj
			.Type = 2: .Mode = 3: .Open: .Charset = "GB2312" : .Position = Obj.Size: .Loadfromfile Server.Mappath(mTemplate): mContent = .ReadText: .Close
		End With
		Set Obj = Nothing
		If Err Then Response.Write Err.Description & Warn("�޷�����ģ��[" & mTemplate & "]"):Response.End
	End Function
	
	' �����ļ�
	Private Function LoadFile(ByVal strFilePath)
		Dim objFile, strTempConent
		On Error Resume Next
		Set objFile = Server.CreateObject("adodb.stream")
		With objFile
			.Type = 2: .Mode = 3: .Open: .Charset = "GB2312" : .Position = objFile.Size: .Loadfromfile Server.Mappath(strFilePath): strTempConent = .ReadText: .Close
		End With
		Set objFile = Nothing
		If Err Then  Response.Write Err.Description & Warn("�޷������ļ�[" & strFilePath & "]"): Response.End
		LoadFile = strTempConent
	End Function

	
	' �Ƿ���ڴ����ǩ
	Private Function RegExists(ByVal pattern, ByVal strContent)
		mReg.Pattern = pattern
		RegExists = mReg.Test(strContent)
	End Function
	
	' �����ʽ�滻
	Private Function RegReplace(ByVal repContent, ByVal pattern, ByVal repValue)
		mReg.Pattern = pattern
		RegReplace = mReg.Replace(repContent, repValue)
	End Function
	
	
	'�滻����
	Private Function Rep(strSource, strDestn)
		mContent = Replace(mContent, strSource, strDestn)
	End Function

End Class
%>