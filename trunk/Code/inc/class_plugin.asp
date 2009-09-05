<%
Class ClassPlugin
	Public mErrorCode
	Public mPlusName
	Public mXMLPath
	Public mXMLDocument
	public mTPL ' ģ��
	
	' ��ʼ��
	Private Sub Class_Initialize()
		mErrorCode = -1
	End Sub
	
	' ע������
	Private Sub Class_Terminate()
		If IsObject(mXMLDocument) Then Set mXMLDocument = Nothing
	End Sub
	
	' �򿪲�������ļ�
	' mErrorCode = 0 ��ʾ�������ļ�
	Public Function Open(PlusNameCode)
		mPlusName = PlusNameCode
		Dim XmlDom
		XmlDom = GetXMLDOM()
		If XmlDom = False Then ErrorCode = -18239123: Exit Function
		mXMLPath = Server.Mappath(InstallDir & "plugin/" & mPlusName & "/plugin.Xml")
		Set mXMLDocument = Server.CreateObject(XmlDom)
		mXMLDocument.Async = False
		mXMLDocument.Load (mXMLPath) ' ���·��
		mErrorCode = mXMLDocument.parseerror.ErrorCode
	End Function
	
	' ���״̬,ǰ��
	Public Function ChkState()
		If Config("state") = 0 Then response.write "���������" : Response.End
	End Function


	' ����ģ��[֧�ָ�ģ��]
	public function NewTpl(tplFile)
		dim reg,tplcomm,tplplus
		tplcomm = ReadFile(InstallDir & templatedir & "/common.html")
		if len(tplcomm) = 0 then tplcomm = "{tag:inside}"
		tplplus = ReadFile(tplfile)
		set reg = new regexp
		reg.ignorecase = true
		reg.global = true
		reg.pattern = "{tag:inside}"
		tplcomm = reg.replace(tplcomm, tplplus) ' ��������ģ��
		reg.pattern = "{tag:sitepath}"
		tplcomm = reg.replace(tplcomm, getsitepathbytitle(config("title"))) ' վ��·��
		reg.pattern = "{field:title}"
		tplcomm = reg.replace(tplcomm, config("title")) ' �滻����
		reg.pattern = "{field:keywords}"
		tplcomm = reg.replace(tplcomm, config("keywords")) ' �滻�ؼ���
		reg.pattern = "{field:description}"
		tplcomm = reg.replace(tplcomm, config("description")) ' �滻����
		reg.pattern = "{(main|plus|config|value|var|val|lang|skin|tpl)\.(.+?)}" ' �滻��������
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
		set mTpl = New ClassTemplate 	' ����ģ�����
		mTpl.Content = tplcomm 			' ����ģ�����
		mTpl.Compile_Plugin		 		' ����ģ���������б�ǩ
	end function
	
	' �滻��ǩ
	Function SetTpl(tag,val)
		If isobject(mTpl) Then mTpl.content = replace(mTpl.Content, tag, val)
	End Function
	
	' �õ�������ģ������
	function GetTpl()
		if isobject(mTpl) then GetTpl = rewriterule(mTpl.Content)
	end function
	
	' ��ȡ������Ϣ
	Public Function Main(Attr)
		Attr = LCase("plugin/main/" & Attr)
		Main = SelectXmlNodeText(Attr)
	End Function
	
	' ��ȡ���
	Public Function Skin(Attr)
		Attr = LCase("plugin/skin/" & Attr)
		Skin = SelectXmlNodeText(Attr)
	End Function
	
	' ��ȡ������Ϣ
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
	
	' ��ȡ���ò�������
	Public Function ConfigLength()
		Dim XmlItem
		Set XmlItem = mXMLDocument.getElementsByTagName(LCase("plugin/config/key"))
		ConfigLength = XmlItem.Length
	End Function

	' ��ȡ������Ϣ
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
	
	' ��������
	Public Function ConfigSave(x, val)
		Dim SubAttr, i
		Set SubAttr = mXMLDocument.getElementsByTagName(LCase("plugin/config/key")).Item(x).Attributes
		For i = 0 To SubAttr.Length - 1
			If LCase(SubAttr(i).Name) = LCase("value") Then SubAttr(i).Value = val: Exit For
		Next
		mXMLDocument.Save mXMLPath
	End Function

	' ��õ��� ElementName Ԫ��
	Public Function SelectXmlNodeText(ElementName)
		Dim XmlItem
		Set XmlItem = mXMLDocument.getElementsByTagName(ElementName)
		If XmlItem.Length <> 0 Then SelectXmlNodeText = XmlItem.Item(0).Text Else SelectXmlNodeText = ""
		If Len(SelectXmlNodeText) > 0 And IsNumeric(SelectXmlNodeText) Then SelectXmlNodeText = Int(SelectXmlNodeText)
	End Function
End Class
%>
