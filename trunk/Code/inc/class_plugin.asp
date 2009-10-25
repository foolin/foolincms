<%
'�㷨�ο�5u
Class ClassPlugin
	Public ErrorCode
	Public PlusName
	Public XMLPath
	Public XMLDocument
	public TPL ' ģ��
	
	' ��ʼ��
	Private Sub Class_Initialize()
		ErrorCode = -1
	End Sub
	
	' ע������
	Private Sub Class_Terminate()
		If IsObject(XMLDocument) Then Set XMLDocument = Nothing
	End Sub
	
	' �򿪲�������ļ�
	' ErrorCode = 0 ��ʾ�������ļ�
	Public Function Open(PlusNameCode)
		PlusName = PlusNameCode
		Dim XmlDom
		XmlDom = GetXMLDOM()
		If XmlDom = False Then ErrorCode = -18239123: Exit Function
		XMLPath = Server.Mappath(Installdir & "plugins/" & PlusName & "/config.Xml")
		Set XMLDocument = Server.CreateObject(XmlDom)
		XMLDocument.Async = False
		XMLDocument.Load (XMLPath) ' ���·��
		ErrorCode = XMLDocument.parseerror.ErrorCode
	End Function
	
	' ���״̬,ǰ��
	Public Function State()
		If Config("state") = 0 Then Response.Write("�Բ��𣬸Ĳ��δ������"): Response.end
	End Function

	' ����ģ��[֧�ָ�ģ��]
	Public Function Newtpl(tplfile)
		dim reg,tplcomm,tplplus
		tplcomm = getfile(installdir & templatedir & "/common.html")
		if len(tplcomm) = 0 then tplcomm = "{tag:inside /}"
		tplplus = GetFile(tplfile)
		set reg = new regexp
		reg.ignorecase = true
		reg.global = true
		reg.pattern = "{tag:inside /}"
		tplcomm = reg.replace(tplcomm, tplplus) ' ��������ģ��
		reg.pattern = "\{sys\s*:\s*sitepath(\s*/)?\}"
		tplcomm = reg.replace(tplcomm, IndexPath & " �� " & Config("title")) ' վ��·��
		reg.pattern = "\{sys\s*:\s*title(\s*/)?\}"
		tplcomm = reg.replace(tplcomm, Config("title")) ' �滻����
		reg.pattern = "\{sys\s*:\s*sitekeywords(\s*/)?\}"
		tplcomm = reg.replace(tplcomm, Config("keywords")) ' �滻�ؼ���
		reg.pattern = "\{sys\s*:\s*sitedesc(\s*/)?\}"
		tplcomm = reg.replace(tplcomm, Config("description")) ' �滻����
		reg.pattern = "{(main|plugin|config|value|var|val|skin|tpl)\.(.+?)}" ' �滻��������
		dim match,matchs
		set matchs = reg.execute(tplcomm)
		for each match in matchs
			select case lcase(match.submatches(0))
			case "main","plugin" ' M Plus
				tplcomm = replace(tplcomm,match.value,main(match.submatches(1)))
			case "config","value","var","val" ' M Config
				tplcomm = replace(tplcomm,match.value,config(match.submatches(1)))
			case "skin","tpl" ' M Skin
				tplcomm = replace(tplcomm,match.value,skin(match.submatches(1)))
			case else
				tplcomm = replace(tplcomm,match.value,"")
			end select
		next
		Set tpl = New ClassTemplate ' ����ģ�����
		tpl.content = tplcomm ' ����ģ�����
		tpl.Compile_Plugin() ' ����ģ���������б�ǩ
	End Function
	
	' �滻��ǩ
	Function SetTpl(tag,val)
		if isobject(tpl) then tpl.content = replace(tpl.content,tag,val)
	End Function
	
	' �õ�������ģ������
	Function GetTpl()
		If Isobject(tpl) then gettpl = rewriterule(tpl.content)
	End Function
	
	' ��ȡ������Ϣ
	Public Function Main(Attr)
		Attr = LCase("plugin/main/" & Attr)
		Main = SelectXmlNodeText(Attr)
	End Function
	
	' ��ȡ���
	Public Function Skin(Attr)
		Attr = LCase("plugin/Skin/" & Attr)
		Skin = SelectXmlNodeText(Attr)
	End Function
	
	' ��ȡ������Ϣ
	Public Function Config(Attr)
		Dim Attr1, Attr2, i, j, x
		Dim XmlItem, SubAttr, ElementName
		Attr = LCase(Replace(Attr, "\", "/")): Attr1 = Attr: Attr2 = "value"
		If InStr(Attr, "/") > 0 Then Attr1 = Split(Attr, "/")(0): Attr2 = Split(Attr, "/")(1)
		Set XmlItem = XMLDocument.getElementsByTagName(LCase("plugin/config/key"))
		For j = 0 To XmlItem.Length - 1
			Set SubAttr = XMLDocument.getElementsByTagName(LCase("plugin/config/key")).Item(j).Attributes
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
		Set XmlItem = XMLDocument.getElementsByTagName(LCase("plugin/config/key"))
		ConfigLength = XmlItem.Length
	End Function

	' ��ȡ������Ϣ
	Public Function ConfigItem(x, Attr)
		Dim i, j
		Dim XmlItem, SubAttr, ElementName
		Set XmlItem = XMLDocument.getElementsByTagName(LCase("plugin/config/key"))
		For j = 0 To XmlItem.Length - 1
			If x = j Then
				Set SubAttr = XMLDocument.getElementsByTagName(LCase("plugin/config/key")).Item(j).Attributes
				For i = 0 To SubAttr.Length - 1
					If LCase(SubAttr(i).Name) = LCase(Attr) Then ConfigItem = SubAttr(i).Value: Exit Function
				Next
			End If
		Next
	End Function
	
	' ��������
	Public Function ConfigSave(x, val)
		Dim SubAttr, i
		Set SubAttr = XMLDocument.getElementsByTagName(LCase("plugin/config/key")).Item(x).Attributes
		For i = 0 To SubAttr.Length - 1
			If LCase(SubAttr(i).Name) = LCase("value") Then SubAttr(i).Value = val: Exit For
		Next
		XMLDocument.Save XMLPath
	End Function

	' ��õ��� ElementName Ԫ��
	Public Function SelectXmlNodeText(ElementName)
		Dim XmlItem
		Set XmlItem = XMLDocument.getElementsByTagName(ElementName)
		If XmlItem.Length <> 0 Then SelectXmlNodeText = XmlItem.Item(0).Text Else SelectXmlNodeText = ""
		If Len(SelectXmlNodeText) > 0 And IsNumeric(SelectXmlNodeText) Then SelectXmlNodeText = Int(SelectXmlNodeText)
	End Function
End Class
%>
