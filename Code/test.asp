<%
	Dim strTest,arrTest
	'strTest = "	�й�	�Ŵ�	��Ů "
	strTest = "	�й�	"
	arrTest = Split(Trim(strTest), "	")
	Response.Write(Split(Trim(strTest), " ")(0))
	Response.Write(IsArray(arrTest))
	Dim i
	For i = 0 To UBound(arrTest)
		Response.Write(i & ":" & arrTest(i) & "<br />")
	Next
%>