<%
Class ClassPlugin
	Public mErrorCode
	Public mPlusName
	Public mXMLPath
	Public mXMLDocument
	public mTPL ' 模板
	
	' 初始化
	Private Sub Class_Initialize()
		mErrorCode = -1
	End Sub
	
	' 注销对象
	Private Sub Class_Terminate()
		If IsObject(mXMLDocument) Then Set mXMLDocument = Nothing
	End Sub
	
	' 打开插件配置文件
	' mErrorCode = 0 表示打开配置文件
	Public Function Open(PlusNameCode)
		mPlusName = PlusNameCode
		Dim XmlDom
		XmlDom = GetXMLDOM()
		If XmlDom = False Then ErrorCode = -18239123: Exit Function
		mXMLPath = Server.Mappath(InstallDir & "plugin/" & mPlusName & "/plugin.Xml")
		Set mXMLDocument = Server.CreateObject(XmlDom)
		mXMLDocument.Async = False
		mXMLDocument.Load (mXMLPath) ' 插件路径
		mErrorCode = mXMLDocument.parseerror.ErrorCode
	End Function
	
	' 检测状态,前面
	Public Function ChkState()
		If Config("state") = 0 Then response.write "插件被禁用" : Response.End
	End Function


	' 载入模板[支持父模板]
	public function NewTpl(tplFile)
		dim reg,tplcomm,tplplus
		tplcomm = ReadFile(InstallDir & templatedir & "/common.html")
		if len(tplcomm) = 0 then tplcomm = "{tag:inside}"
		tplplus = ReadFile(tplfile)
		set reg = new regexp
		reg.ignorecase = true
		reg.global = true
		reg.pattern = "{tag:inside}"
		tplcomm = reg.replace(tplcomm, tplplus) ' 载入内置模板
		reg.pattern = "{tag:sitepath}"
		tplcomm = reg.replace(tplcomm, getsitepathbytitle(config("title"))) ' 站内路径
		reg.pattern = "{field:title}"
		tplcomm = reg.replace(tplcomm, config("title")) ' 替换标题
		reg.pattern = "{field:keywords}"
		tplcomm = reg.replace(tplcomm, config("keywords")) ' 替换关键字
		reg.pattern = "{field:description}"
		tplcomm = reg.replace(tplcomm, config("description")) ' 替换描述
		reg.pattern = "{(main|plus|config|value|var|val|lang|skin|tpl)\.(.+?)}" ' 替换基本参数
		dim match,matchs
		set matchs = reg.execute(tplcomm)
		for each match in matchs
			select case lcase(match.submatches(0))
			case "main","plus" ' M Plus
				tplcomm = replace(tplcomm,match.value,main(match.submatches(1)))
			case "config","value","var","val" ' M Config
				tplcomm = replace(tplcomm,match.value,config(match.submatches(1)))
			case "skin","tpl" ' M Skin
				tplcomm = replace(tplcomm,match.value,skin(match.submatches(1)))
			case else
				tplcomm = replace(tplcomm,match.value,"")
			end select
		next
		set mTpl = New ClassTemplate 	' 创建模板对像
		mTpl.Content = tplcomm 			' 设置模板代码
		mTpl.Compile_Plugin		 		' 加载模板后解析所有标签
	end function
	
	' 替换标签
	Function SetTpl(tag,val)
		If isobject(mTpl) Then mTpl.content = replace(mTpl.Content, tag, val)
	End Function
	
	' 得到处理后的模板内容
	function GetTpl()
		if isobject(mTpl) then GetTpl = rewriterule(mTpl.Content)
	end function
	
	' 获取基本信息
	Public Function Main(Attr)
		Attr = LCase("plugin/main/" & Attr)
		Main = SelectXmlNodeText(Attr)
	End Function
	
	' 获取风格
	Public Function Skin(Attr)
		Attr = LCase("plugin/skin/" & Attr)
		Skin = SelectXmlNodeText(Attr)
	End Function
	
	' 获取配置信息
	Public Function Config(Attr)
		Dim Attr1, Attr2, i, j, x
		Dim XmlItem, SubAttr, ElementName
		Attr = LCase(Replace(Attr, "\", "/")): Attr1 = Attr: Attr2 = "value"
		If InStr(Attr, "/") > 0 Then Attr1 = Split(Attr, "/")(0): Attr2 = Split(Attr, "/")(1)
		Set XmlItem = mXMLDocument.getElementsByTagName(LCase("plugin/config/key"))
		For j = 0 To XmlItem.Length - 1
			Set SubAttr = mXMLDocument.getElementsByTagName(LCase("plugin/config/key")).Item(j).Attributes
			For i = 0 To SubAttr.Length - 1
				If LCase(SubAttr(i).Name) = "name" Then
					If LCase(SubAttr(i).Value) = LCase(Attr1) Then
						For x = 0 To SubAttr.Length - 1
							If LCase(SubAttr(x).Name) = LCase(Attr2) Then
								Config = SubAttr(x).Value
								If Len(Config) > 0 And IsNumeric(Config) And Instr(Config,",")=0 Then Config = Int(Config)
								Exit Function
							End If
						Next
					End If
				End If
			Next
		Next
	End Function
	
	' 获取配置参数个数
	Public Function ConfigLength()
		Dim XmlItem
		Set XmlItem = mXMLDocument.getElementsByTagName(LCase("plugin/config/key"))
		ConfigLength = XmlItem.Length
	End Function

	' 获取配置信息
	Public Function ConfigItem(x, Attr)
		Dim i, j
		Dim XmlItem, SubAttr, ElementName
		Set XmlItem = mXMLDocument.getElementsByTagName(LCase("plugin/config/key"))
		For j = 0 To XmlItem.Length - 1
			If x = j Then
				Set SubAttr = mXMLDocument.getElementsByTagName(LCase("plus/config/key")).Item(j).Attributes
				For i = 0 To SubAttr.Length - 1
					If LCase(SubAttr(i).Name) = LCase(Attr) Then ConfigItem = SubAttr(i).Value: Exit Function
				Next
			End If
		Next
	End Function
	
	' 保存配置
	Public Function ConfigSave(x, val)
		Dim SubAttr, i
		Set SubAttr = mXMLDocument.getElementsByTagName(LCase("plugin/config/key")).Item(x).Attributes
		For i = 0 To SubAttr.Length - 1
			If LCase(SubAttr(i).Name) = LCase("value") Then SubAttr(i).Value = val: Exit For
		Next
		mXMLDocument.Save mXMLPath
	End Function

	' 获得当个 ElementName 元素
	Public Function SelectXmlNodeText(ElementName)
		Dim XmlItem
		Set XmlItem = mXMLDocument.getElementsByTagName(ElementName)
		If XmlItem.Length <> 0 Then SelectXmlNodeText = XmlItem.Item(0).Text Else SelectXmlNodeText = ""
		If Len(SelectXmlNodeText) > 0 And IsNumeric(SelectXmlNodeText) Then SelectXmlNodeText = Int(SelectXmlNodeText)
	End Function
End Class
%>
