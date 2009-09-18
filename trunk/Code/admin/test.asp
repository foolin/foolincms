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
		Call MsgBox("成功:" & admin.Password,"BACK")
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
 'If CreateFile("aaaaaa" & chr(10) & chr(9) & "bbbbb" ,fileDir) = True Then Response.Write("HHH成功") else Response.Write("×××")
 'If CreateFolder(fileDir) = True Then Response.Write("HHH成功") else Response.Write("×××")
 'If DeleteFolder(fileDir) = True Then Response.Write("成功") else Response.Write("×××")
 'Response.Write GetFile(fileDir)
 If DeleteFile(fileDir) = True Then Response.Write("Delete成功") else Response.Write("×××")
 'If ExistFolder(fileDir) = True Then Response.Write("存在") else Response.Write("×××")
 'If ExistFile(fileDir) = True Then Response.Write("存在") else Response.Write("×××")
%>
<script type="text/javascript" charset="gb2312" src="../inc/kindeditor/kindeditor.js"></script>
<script type="text/javascript">
<!--
//初始化编辑器
KE.show({
    id : 'Content1',
	filterMode: false
});
//-->
</script>
	<form action="" method="post">
        <table width="100%">
          <tr>
            <td colspan="2" class="title">E酷CMS系统管理登录</td>
            <input type="hidden" name="action" value="admin" />
          </tr>
          <tr>
            <td class="txtR">用户名：</td>
            <td class="txtL"><input name="Username" class="input" style="width:150px;" type="text" /></td>
          </tr>
          <tr>
            <td class="txtR">密&nbsp;&nbsp;码：</td>
            <td class="txtL"><input name="Password" class="input" style="width:150px;"  type="password" /></td>
          </tr>
          <tr>
            <td class="txtR">普通管理员：</td>
            <td class="txtL"><input name="Level" class="input"  style="width:100px;"  type="text" />
                                <textarea id="Content1" name="Content" style="width:100%;height:400px;visibility:hidden;">
                    	aaaa
                    </textarea>

            
            </td>
          </tr>
          <tr>
            <td colspan="2"><input type="submit" class="btn" value="登录" />
            <input type="reset" class="btn" value="重置" /></td>
          </tr>
        </table>
	</form>