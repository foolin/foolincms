<%
	Dim strTest,arrTest
	'strTest = "	中国	古代	美女 "
	strTest = "	中国	"
	arrTest = Split(Trim(strTest), "	")
	Response.Write(Split(Trim(strTest), " ")(0))
	Response.Write(IsArray(arrTest))
	Dim i
	For i = 0 To UBound(arrTest)
		Response.Write(i & ":" & arrTest(i) & "<br />")
	Next
%>