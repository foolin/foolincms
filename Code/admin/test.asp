<!--#include file="inc/admin.include.asp"-->
<!--#include file="lib/class_admin.asp"-->
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
 
%>
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
            <td class="txtL"><input name="Level" class="input"  style="width:100px;"  type="text" /></td>
          </tr>
          <tr>
            <td colspan="2"><input type="submit" class="btn" value="登录" />
            <input type="reset" class="btn" value="重置" /></td>
          </tr>
        </table>
	</form>
    
<%
Dim strCode, strUrl, iStart, iEnd
    strCode = "<embed src=http://files2.17173.com/dzflash/qe.swf pluginspage=' type='application/x-shockwave-flash' width=300 height=200></embed></OBJECT>"
    iStart = Instr(strCode, "src=") + 4
    iEnd = Instr(strCode, "pluginspage=")
    strUrl = Trim(Mid(strCode, iStart, iEnd - iStart))
    Response.Write strUrl
%>	