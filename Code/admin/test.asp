<!--#include file="inc/admin.include.asp"-->
<!--#include file="lib/class_admin.asp"-->
<!--#include file="../inc/func_file.asp"-->
<%
 If Request("action") = "admin" Then
 	Dim admin
	Set admin = New ClassAdmin
	admin.ID = 1
	'If admin.SetValue And admin.Modify Then
	If admin.LetValue Then
		Call MsgBox("�ɹ�:" & admin.Password,"BACK")
	End If
	Set admin = Nothing
 End If
 Dim strTest: strTest = "aaa|bbb|ccc"
 Dim arrPicPath, i
	arrPicPath = Split(strTest, "|")
	For i = 0 To UBound(arrPicPath)
		'vPicPath = arrPicPath(i)
		strTest = arrPicPath(i) & "<br />-------------------<br>"
		Response.Write(strTest)
	Next
 Dim fileDir
 	fileDir = "HHH/EEE/KKKK/ZZZZ/DDDD/UUUU/abc.txt"
	'Response.Write Left(fileDir, InStrRev(fileDir,"/"))
 'If CreateFile("aaaaaa" & chr(10) & chr(9) & "bbbbb" ,fileDir) = True Then Response.Write("HHH�ɹ�") else Response.Write("������")
 'If CreateFolder(fileDir) = True Then Response.Write("HHH�ɹ�") else Response.Write("������")
 'If DeleteFolder(fileDir) = True Then Response.Write("�ɹ�") else Response.Write("������")
 'Response.Write GetFile(fileDir)
 If DeleteFile(fileDir) = True Then Response.Write("Delete�ɹ�") else Response.Write("������")
 'If ExistFolder(fileDir) = True Then Response.Write("����") else Response.Write("������")
 'If ExistFile(fileDir) = True Then Response.Write("����") else Response.Write("������")
%>
<script type="text/javascript" charset="gb2312" src="../inc/kindeditor/kindeditor.js"></script>
<script type="text/javascript">
<!--
//��ʼ���༭��
KE.show({
    id : 'Content1',
	filterMode: false
});
//-->
</script>
	<form action="" method="post">
        <table width="100%">
          <tr>
            <td colspan="2" class="title">E��CMSϵͳ�����¼</td>
            <input type="hidden" name="action" value="admin" />
          </tr>
          <tr>
            <td class="txtR">�û�����</td>
            <td class="txtL"><input name="Username" class="input" style="width:150px;" type="text" /></td>
          </tr>
          <tr>
            <td class="txtR">��&nbsp;&nbsp;�룺</td>
            <td class="txtL"><input name="Password" class="input" style="width:150px;"  type="password" /></td>
          </tr>
          <tr>
            <td class="txtR">��ͨ����Ա��</td>
            <td class="txtL"><input name="Level" class="input"  style="width:100px;"  type="text" />
                                <textarea id="Content1" name="Content" style="width:100%;height:400px;visibility:hidden;">
                    	aaaa
                    </textarea>

            
            </td>
          </tr>
          <tr>
            <td colspan="2"><input type="submit" class="btn" value="��¼" />
            <input type="reset" class="btn" value="����" /></td>
          </tr>
        </table>
	</form>