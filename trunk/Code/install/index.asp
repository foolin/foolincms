<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../inc/func_common.asp"-->
<!--#include file="../inc/func_file.asp"-->
<!--#include file="../inc/md5.asp"-->
<%

	Dim act: act = Request("action")
	If Len(act) = 0 Then act = "step1"
	If LCase(act) = "step3" Then
		Call Install()
	End If
	If LCase(act) = "step1" Or LCase(act) = "step2" Then
		Call ChkInstall()
	End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; varcharset=gb2312" />
   <title>E�����ݹ���ϵͳ-��װ</title>
<style type="text/css">
<!--
body{
	font-family:Georgia, "Times New Roman", Times, serif;
	font-size:13px;
}
p{margin:5px;}

.wrap{
	text-align:center;
}

.footer{
	line-height:22px;
	text-align:center;
	padding-top:30px;
}
.footer p{margin:5px;}

.title{
	font-size:24px;
	font-weight:bold;
	text-align:center;
	padding-top:20px;
	padding-bottom:20px;
}
.btn{
	text-align:center;
	padding:10px;
}
.btn input{
	padding:6px;
	font-size:14px;
}


#step1{ text-align:left;}
.license {
	margin:0px auto;
	line-height:22px;
	height:450px;
	width:700px;
	padding:10px 20px;
	border:#EBEBEB 8px solid;
	overflow:scroll;
	scrollbar-face-color:#EEE ;
	scrollbar-shadow-color: #ffffff; 
	scrollbar-highlight-color:#ffffff; 
	scrollbar-3dlight-color: #ffffff;  
	scrollbar-darkshadow-color: #ffffff; 
	scrollbar-track-color:#ffffff; 
	scrollbar-arrow-color: ffffff;
	background:#F9F9F9;
}


#step2 {}
.form{margin:0px auto; border:#EBEBEB 5px solid;}
table.form { width:550px;}
table.form tr{background:#F3F3F3;}
table.form td{ padding:5px;}
td.name { text-align:right;}
td.inputtxt { width:75%; text-align:left; color:#666;}
.inputtxt input{ width:250px; height:22px; line-height:22px;}


#step3 {}
.state{ margin:0px auto; border:#EBEBEB 5px solid; width:650px; text-align:left;}
.state dl{ margin:2px; background:#F3F3F3; padding:10px;}
.state dl dt{ text-align:center; font-size:16px; font-weight:bold; background:#FFF; line-height:35px; margin-bottom:10px;}
.state dl dd{ line-height:25px; border-bottom:2px #FFF solid; margin-left:0px; padding-left:10px; }
.success{ color:green;}
.warn{ color:#00F;}
.error{ color:#F00;}


#cancel {}
.cancel{ margin:0px auto;  text-align:left; border:#EBEBEB 5px solid; width:600px; line-height:25px; padding:3px;}
.ourinfo{ background:#F9F9F9; padding:20px 10px;}

#hasInstall {}
.box{ margin:0px auto;  text-align:left; border:#EBEBEB 5px solid;background:#F9F9F9; width:700px; line-height:25px; padding:3px;}

-->
</style>
<script language="javascript">
	function go(act){
		this.location.href='index.asp?action=' + act;
	}
</script>
</head>
<body>
<div class="wrap">

<%
	Select Case LCase(act)
		Case "step1"
			Call Step1()
		Case "step2"
			Call Step2()
		Case "step3"
			Call Step3()
		Case "cancel"
			Call Cancel()
		Case "hasinstall"
			Call HasInstall()
		Case Else
			Call Step1()
	End Select
%>
    
    <div class="footer">
            <p>��Ȩ���� (c)2009-2010��E�Ṥ���� (www.eekku.com) ��������Ȩ���� </p>
            <p>��ϵͳ��Foolin(����)����������Author: Foolin &nbsp;&nbsp; QQ: 970026999 E-mail: Foolin@126.com </p>
    </div>
    
</div>

</body>
</html>
<!-- ����װ����������������   E-mail:Foolin@126.com   2009��7��28��15:14:40 -->

<%Sub Step1()%>

	<!--��һ��:���Э��-->
    <div id="step1">
    	<div class="title">��һ����E�����ݹ���ϵͳ��EekkuCMS����װ���Э��</div>
        <div class="license">
          <p>��л��ѡ��E�����ݹ���ϵͳ�����¼��EekkuCMS����EekkuCMS��һ������ʵ�õ�С��վ����򣬻��� ASP + Access �ļ���������ȫ��Դ�뿪�š� </p>
          <p>EekkuCMS �Ĺٷ���ַ�ǣ� www.eekku.com</p>
          <p>Ϊ��ʹ����ȷ���Ϸ���ʹ�ñ������������ʹ��ǰ����Ķ���������Э�����</p>
          <p>һ������ȨЭ�������ҽ������� EekkuCMS 1.x.x �汾��E�Ṥ���ҶԱ���ȨЭ��ӵ�����ս���Ȩ��</p>
          <p>����Э����ɵ�Ȩ�� <br />
            &nbsp;&nbsp;&nbsp;&nbsp; 1������������ȫ���ر������û���ȨЭ��Ļ����ϣ��������Ӧ���ڷ���ҵ��;��������֧�������Ȩ��Ȩ���á� <br />
            &nbsp;&nbsp;&nbsp;&nbsp; 2����������Э��涨��Լ�������Ʒ�Χ���޸� EekkuCMS Դ��������������Ӧ������վҪ�� <br />
            &nbsp;&nbsp;&nbsp;&nbsp;   3����ӵ��ʹ�ñ������������վȫ����������Ȩ���������е�����Щ���ݵ���ط������� <br />
            &nbsp;&nbsp;&nbsp;&nbsp;   4�������ҵ��Ȩ֮�������Խ������Ӧ������ҵ��;��ͬʱ�������������Ȩ������ȷ���ļ���֧�����ݣ��Թ���ʱ�����ڼ���֧��������ӵ��ͨ��ָ���ķ�ʽ���ָ����Χ�ڵļ���֧�ַ�����ҵ��Ȩ�û����з�ӳ����������Ȩ����������������Ϊ��Ҫ���ǣ���û��һ�������ɵĳ�ŵ��֤�� </p>
<p><strong>����Э��涨��Լ�������� </strong><br />
   &nbsp;&nbsp;&nbsp;&nbsp; 1�����ý���������ڹ��Ҳ����������վ������ɫ�顢���������в������Ĳ�����վ����<br />
   &nbsp;&nbsp;&nbsp;&nbsp; 2��δ���ٷ���ɣ����öԱ��������֮��������ҵ��Ȩ���г��⡢���ۡ���Ѻ�򷢷������֤�� <br />
   &nbsp;&nbsp;&nbsp;&nbsp; 3��δ���ٷ���ɣ���ֹ�ڱ������������κβ��ֻ������Է�չ�κ������汾���޸İ汾��������汾�������·ַ��� <br />
  &nbsp;&nbsp;&nbsp;&nbsp;  4�������δ�����ر�Э������������Ȩ������ֹ��������ɵ�Ȩ�������ջأ����е���Ӧ�������Ρ� </p>
<p><strong>�ġ����޵������������� </strong><br />
   &nbsp;&nbsp;&nbsp;&nbsp; 1������������������ļ�����Ϊ���ṩ�κ���ȷ�Ļ��������⳥�򵣱�����ʽ�ṩ�ġ� <br />
   &nbsp;&nbsp;&nbsp;&nbsp; 2���û�������Ը��ʹ�ñ�������������˽�ʹ�ñ�����ķ��գ�����δ�����Ʒ��������֮ǰ�����ǲ���ŵ������û��ṩ�κ���ʽ�ļ���֧�֡�ʹ�õ�����Ҳ���е��κ���ʹ�ñ���������������������Ρ� <br />
   &nbsp;&nbsp;&nbsp;&nbsp; 3�������ı���ʽ����ȨЭ����ͬ˫������ǩ���Э��һ����������ȫ�ĺ͵�ͬ�ķ���Ч������һ����ʼȷ�ϱ�Э�鲢��װ��ϵͳ��������Ϊ��ȫ��Ⲣ���ܱ�Э��ĸ�������������������������Ȩ����ͬʱ���ܵ���ص�Լ�������ơ�Э����ɷ�Χ�������Ϊ����ֱ��Υ������ȨЭ�鲢������Ȩ��������Ȩ��ʱ��ֹ��Ȩ������ֹͣ�𺦣�������׷��������ε�Ȩ���� <br />
   &nbsp;&nbsp;&nbsp;&nbsp; 4���������������������������APIʾ�����Ӱ�����Щ�ļ���Ȩ�����ڱ�����ٷ���������Щ�ļ���û������Ȩ�����ģ���ο���������ʹ����ɺϷ���ʹ�á�</p>
<p>��Ȩ���� (c)2009-2010��E�Ṥ���� ��������Ȩ���� </p>
<p>Э�鷢��ʱ�䣺  2009��9��21�� By Foolin </p>
<p>�汾���¸��£�2009��9��21�� By Foolin </p>
        </div>
        <div class="btn">
            <input type="button" value="ͬ��"  onclick="go('step2');"/> &nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="��ͬ��" onclick="go('cancel');" />
        </div>
    </div>
<%
 End Sub
 
 
 Sub Step2()
%>
    <div id="step2">
    	<div class="title">�ڶ�����E�����ݹ���ϵͳ��EekkuCMS����װ</div>
        <form action="?action=step3" method="post" name="form1" onsubmit="return chkForm();">
        <table class="form">
            <tr>
                <td class="name">��¼�ʺţ�</td><td class="inputtxt"><input name="Username" type="text" /> ����д����Ա�˺�</td>
            </tr>
            <tr>
                <td class="name">��¼���룺</td><td  class="inputtxt"><input name="Password" type="password" /> ��д����Ա����</td>
            </tr>
            <tr>
                <td class="name">�ظ����룺</td><td  class="inputtxt"><input name="RePassword" type="password" /></td>
            </tr>
            <tr>
                <td class="name">���ݿ����ƣ�</td><td  class="inputtxt"><input name="DbName" type="text" value="Fl<%=Day(Now())%>#Ek_<%=Left(UCase(MD5(Now())),10)%>.mdb" /> ����д���ݿ�����</td>
            </tr>
            <tr>
                <td colspan="2" class="btn"> <input type="button" value="��һ��" onclick="go('step1');" /> &nbsp;&nbsp;&nbsp;&nbsp;
        		<input type="submit" value="��һ��" /></td>
            </tr>
        </table>
        </form>
    </div>
<script language="javascript">
<!--
function chkForm(){
	var form = document.forms["form1"];
	if( form.elements["Username"].value == ""){
		alert("�˺Ų���Ϊ��");
		return false;
	}
	if( form.elements["Password"].value == ""){
		alert("���벻��Ϊ��");
		return false;
	 }
	if( form.elements["Password"].value.length < 6){
		alert("���벻������6λ");
		return false;
	 }
	if( form.elements["RePassword"].value != form.elements["Password"].value){
		alert("�������벻һ�£�");
		return false;
	 }
	if( form.elements["DbName"].value.length < 6){
		alert("���ݿ�����������׺����������6λ!");
		return false;
	 }
	return true;
}
//-->
</script>

<%
 End Sub
 
 
 Sub Step3()
%>

   <div id="step3">
    		<div class="title">����������ɱ�ϵͳ�İ�װ</div>
			<div class="state">
            <dl>
            	<dt> ϵͳ��װ�ɹ�</dt>
            	<dd> �ʺţ�<span class="success"><%=Request("Username")%></span></dd>
                <dd> ���룺<span class="success"><%=Request("Password")%></span></dd>
            	<dd> ��ǰ��װĿ¼�� <%=Replace(Request.ServerVariables("Path_Info"),"/install/index.asp", "")%> </dd>
                <dd> <a href="../index.asp">������ҳ</a></dd>
                <dd> <a href="../admin/login.asp">�����̨����</a> </dd>
              <dd>&nbsp; </dd>
                <dd>   <span class="warn">��ע�⣺</span> ��������а�װ�����밴�հ�װ˵���������������ϵͳ��</dd>
               <dd> <span class="warn">���飺</span>Ϊ�˱�ϵͳ��ȫ����ֹ�������֣��������̰�InstallĿ¼������������Խ���˲µ�Խ�ã�������ֱ��ɾ�����ǳ���л��ʹ�ñ�ϵͳ��ף����;��죡</dd>
                <dd class="btn"> <input type="button" value="���" onclick="go('hasinstall');"   /></dd>
            </dl>
            </div>
    </div>
    
<%
 End Sub
 
 
 Sub Cancel()
%>
    <div id="cancel">
    	<div class="title">���Ѿ�ȡ���˱���ϵͳ�İ�װ��</div>
        <div class="cancel">
        	<div class="ourinfo">
            &nbsp;&nbsp;<b>Author: Foolin(����) </b><br />
            	&nbsp;&nbsp;&nbsp;&nbsp; QQ��970026999<br />
                &nbsp;&nbsp;&nbsp;&nbsp; E-mail��Foolin@126.com<br /><br />
                &nbsp;&nbsp;&nbsp;&nbsp; Home Page: http://ling.liufu.org<br /><br />
            ��ַ��http://www.eekku.com ��E�Ṥ���ң�<br /><br />
            ��ϵͳ��Foolin����������������κ�������߽��飬����ϵ���ߣ���ָ�л��<br />
            </div>
        </div>
    </div>
    
<%
 End Sub
 
 
 Sub HasInstall()
%>
    
    <div id="hasInstall">
    	<div class="title">���Ѿ���װ�˱�ϵͳ��</div>
        <div class="box">
            <div class="error">�����Ҫ���°�װ�����ֹ�ɾ��Install/Ŀ¼�µ�Install.lock�ļ���Ȼ��<a href="index.asp">�������</a>�����а�װ��</div>
           <p>
            	>> <a href="../index.asp">������ҳ</a> <br />
              	>> <a href="../admin/login.asp">�����̨����</a>
                
                <div class="warn">!ע�⣺�������ҳ���������а���˵�������ֶ�����inc/config.asp��</div>
           </p>
        	<br />
        	<div class="ourinfo">
            &nbsp;&nbsp;<b>Author: Foolin(����) </b><br />
            	&nbsp;&nbsp;&nbsp;&nbsp; QQ��970026999<br />
                &nbsp;&nbsp;&nbsp;&nbsp; E-mail��Foolin@126.com<br /><br />
                &nbsp;&nbsp;&nbsp;&nbsp; Home Page: http://ling.liufu.org<br /><br />
            ��ַ��http://www.eekku.com ��E�Ṥ���ң�<br /><br />
            ��ϵͳ��Foolin����������������κ�������߽��飬����ϵ���ߣ���ָ�л��<br />
            </div>
        </div>
    </div>
    
<%
 End Sub
 

'��鰲װ����
Function ChkInstall()
	If ExistFile("install.lock") Then
		Response.Redirect("?action=hasinstall")
		ChkInstall = True
	End If
	ChkInstall = True
End Function

Function Install()
	Call ChkInstall()	'����Ƿ��Ѿ���װ
	Dim strUsername, strPassword
	strUsername = Request("Username")
	strPassword = Request("Password")
	If Len(strUsername) < 3 Then Call MsgBox("����Ա�û����������������ַ�","BACK")
	If Len(strPassword) < 6 Then Call MsgBox("���벻������6���ַ�","BACK")
	If Request("RePassword")<>strPassword Then Call MsgBox("�������벻һ�£�","BACK")
	If Len(Request("DbName")) < 6 Then Call MsgBox("Ϊ�����ݿⰲȫ�����ݿ�����������6���ַ�","BACK")
	If (Instr(Request("DbName"),"/") > 0) Or (Instr(Request("DbName"),"\") > 0) Then
		Call MsgBox("���ݿ������벻Ҫ���֡�/����\��������·����","BACK")
	End If
	

	strPassword = MD5(strPassword)
	'�������ݿ�
	Call CreateDB(Request("DbName"))
	'�������ݿ��	
	Call CreateTable(Request("DbName"), strUsername, strPassword)	
	'���������ļ�
	Call CreateConfig(Request("DbName"))
	'������װ�ļ�	
	Call LockInstall()	
	Install = True
End Function

'����SQL���
Function CreateTable(strDbName, strUsername, strPassword)
	Dim Conn: Set Conn=Server.CreateObject("ADODB.Connection")
	If Instr(DbName,":\")=0 And Instr(DbName,":/")=0 Then
		strDbName = Server.MapPath("../database/" & strDbName)
	End If 
	Conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDbName
	
	'[Admin]:
	Conn.execute("CREATE TABLE [Admin] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[Username] varchar(20),"&_
		"[Nickname] varchar(20),"&_
		"[Password] varchar(50),"&_
		"[LoginTime] datetime Default Now(),"&_
		"[LoginCount] integer Default 0,"&_
		"[Level] integer Default 0,"&_
		"[LoginIP] varchar(50)"&_
		")")
	Conn.execute("CREATE Unique INDEX [PrimaryKey] on [Admin]([ID] ) with Primary")
	
	'[ArtColumn]:
	Conn.execute("CREATE TABLE [ArtColumn] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[Name] varchar(50),"&_
		"[Info] varchar(250),"&_
		"[ParentID] integer Default 0,"&_
		"[Template] varchar(20)"&_
		")")
	Conn.execute("CREATE Unique INDEX [PrimaryKey] on [ArtColumn]([ID] ) with Primary")
	Conn.execute("CREATE INDEX [ArtClassID] on [ArtColumn]([ID] )")
	
	'[Article]:
	Conn.execute("CREATE TABLE [Article] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[ColID] integer Default 0,"&_
		"[Title] varchar(50),"&_
		"[Author] varchar(20),"&_
		"[Source] varchar(250),"&_
		"[JumpUrl] varchar(250),"&_
		"[Hits] integer Default 0,"&_
		"[FocusPic] varchar(250),"&_
		"[Content] text,"&_
		"[KeyWords] varchar(250),"&_
		"[IsTop] integer Default 0,"&_
		"[IsFocusPic] integer Default 0,"&_
		"[State] integer Default 0,"&_
		"[CreateTime] datetime Default Now(),"&_
		"[ModifyTime] datetime Default Now()"&_
		")")
	Conn.execute("CREATE Unique INDEX [PrimaryKey] on [Article]([ID] ) with Primary")
	Conn.execute("CREATE INDEX [KeyWords] on [Article]([KeyWords] )")
	
	'[Comment]:
	Conn.execute("CREATE TABLE [Comment] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[CType] varchar(20),"&_
		"[User] varchar(20),"&_
		"[Email] varchar(20),"&_
		"[HomePage] varchar(250),"&_
		"[Title] varchar(50),"&_
		"[Content] varchar(250),"&_
		"[Ip] varchar(50),"&_
		"[CreateTime] datetime Default now(),"&_
		"[State] integer Default 0"&_
		")")
	Conn.execute("CREATE Unique INDEX [Index_0D7D3CC3_EF47_4F3A] on [Comment]([ID] ) with Primary")
	
	'[DiyPage]:
	Conn.execute("CREATE TABLE [DiyPage] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[Title] varchar(50),"&_
		"[PageName] varchar(20),"&_
		"[Keywords] varchar(250),"&_
		"[Template] varchar(20),"&_
		"[Code] text,"&_
		"[State] integer Default 0,"&_
		"[IsSystem] integer Default 0"&_
		")")
	Conn.execute("CREATE Unique INDEX [PrimaryKey] on [DiyPage]([ID] ) with Primary")
	Conn.execute("CREATE INDEX [ChannelID] on [DiyPage]([ID] )")
	
	'[GuestBook]:
	Conn.execute("CREATE TABLE [GuestBook] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[User] varchar(20),"&_
		"[Email] varchar(20),"&_
		"[HomePage] varchar(50),"&_
		"[Ip] varchar(50),"&_
		"[Title] varchar(50),"&_
		"[Content] varchar(250),"&_
		"[CreateTime] datetime Default now(),"&_
		"[Recomment] varchar(250),"&_
		"[ReUser] varchar(20),"&_
		"[ReTime] datetime,"&_
		"[State] integer Default 0"&_
		")")
	Conn.execute("CREATE Unique INDEX [Index_0D7D3CC3_EF47_4F3A] on [GuestBook]([ID] ) with Primary")
	
	'[MyTags]:
	Conn.execute("CREATE TABLE [MyTags] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[Name] varchar(50),"&_
		"[Info] varchar(250),"&_
		"[Code] text"&_
		")")
	Conn.execute("CREATE Unique INDEX [Index_621A58EA_0C67_4CE4] on [MyTags]([ID] ) with Primary")
	
	'[PicColumn]:
	Conn.execute("CREATE TABLE [PicColumn] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[Name] varchar(50),"&_
		"[Info] varchar(250),"&_
		"[ParentID] integer Default 0,"&_
		"[Template] varchar(20)"&_
		")")
	Conn.execute("CREATE Unique INDEX [PrimaryKey] on [PicColumn]([ID] ) with Primary")
	Conn.execute("CREATE INDEX [ArtClassID] on [PicColumn]([ID] )")
	
	'[Picture]:
	Conn.execute("CREATE TABLE [Picture] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[Title] varchar(50),"&_
		"[ColID] integer Default 0,"&_
		"[Author] varchar(20),"&_
		"[Source] varchar(250),"&_
		"[SmallPicPath] varchar(250),"&_
		"[PicPath] varchar(250),"&_
		"[IsTop] integer Default 0,"&_
		"[State] integer Default 0,"&_
		"[Intro] varchar(250),"&_
		"[Hits] integer Default 0,"&_
		"[CreateTime] datetime Default Now()"&_
		")")
	Conn.execute("CREATE Unique INDEX [PrimaryKey] on [Picture]([ID] ) with Primary")
	Conn.execute("CREATE INDEX [PhotoID] on [Picture]([ID] )")
	
	'[UploadFile]:
	Conn.execute("CREATE TABLE [UploadFile] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[Title] varchar(50),"&_
		"[Path] varchar(250),"&_
		"[Info] varchar(250),"&_
		"[Ext] varchar(20),"&_
		"[Size] integer,"&_
		"[Author] varchar(20),"&_
		"[DownCount] integer Default 0,"&_
		"[CreateTime] datetime Default Now()"&_
		")")
	Conn.execute("CREATE Unique INDEX [PrimaryKey] on [UploadFile]([ID] ) with Primary")
	Conn.execute("CREATE INDEX [FileID] on [UploadFile]([ID] )")
	
	'[WebLog]:
	Conn.execute("CREATE TABLE [WebLog] ("&_
		"[ID] integer IDENTITY (1,1) not null,"&_
		"[Username] varchar(20),"&_
		"[UserAction] varchar(250),"&_
		"[UserIP] varchar(50),"&_
		"[ActionUrl] varchar(250),"&_
		"[CreateTime] datetime Default Now()"&_
		")")
	Conn.execute("CREATE Unique INDEX [PrimaryKey] on [WebLog]([ID] ) with Primary")
	Conn.execute("CREATE INDEX [ID] on [WebLog]([ID] )")
	'��������Ա��ʼ����
	Conn.execute("INSERT INTO [Admin] ([Username],[Nickname],[Password],[Level],[LoginCount],[LoginTime],[LoginIP]) VALUES('"& strUsername &"','"& strUsername &"','"& strPassword &"',3,0,'"& Now() &"','"& GetIP() &"')")
	'������ʼ���Զ����ǩ
	Conn.execute("INSERT INTO [MyTags] ([Name],[Info],[Code]) VALUES('FirstTag','��һ���Զ����ǩ','��һ���Զ����ǩ��(��������)')")
	'ϵͳ��Ϣ
	Conn.execute("INSERT INTO [MyTags] ([Name],[Info],[Code]) VALUES('AboutSys','����ϵͳ��Ϣ','{sys:sys /}<br />���ߣ�Foolin<br /> Email: Foolin@126.com <br /> ��ҳ��http://www.LiuFu.org/Ling<br />������Http://www.eekku.com<br />')")
	'�Զ����ǩ����������
	Conn.execute("INSERT INTO [MyTags] ([Name],[Info],[Code]) VALUES('FriendLinks','��������','<a href=""http://www.eekku.com""> -==- E���� -==- </a><br />"& Chr(10) & Chr(9) &"<a href=""http://www.liufu.org/ling/""> -==- �������� -==- </a><br />"& Chr(10) & Chr(9) &"')")
	'����ҳ��
	Conn.execute("INSERT INTO [DiyPage] ([Title],[PageName],[Code],[State],[IsSystem]) VALUES('�����ĵ�','help.html','<p>�����ҵĵ�һ���Զ���ҳ�棬��ӭ��ҹ��١�</p><p>��ֻҪ�ں�̨���һ��ҳ�棬Ȼ�����ø����ӣ�<font color=""red"">diypage.asp?id=<font color=""blue"">[������ҳ���ID]</font></font>����<font color=""red"">diypage.asp?url=<font color=""blue"">[������ҳ�������]</font></font>���ɴ�����Զ���ҳ�棡</p><p>�������������ʣ�http://www.eekku.com��E�����绶ӭ�㡣</p><p>������ʹ��E��CMSϵͳ��</p><p><br><br><a href=""http://www.eekku.com""> -==- E���� -==- </a></p><p><br /></p><p><a href=""http://www.liufu.org/ling/""> -==- �������� -==- </a><br /></p>',1,0)")
	'��������
	Conn.execute("INSERT INTO [DiyPage] ([Title],[PageName],[Code],[State],[IsSystem]) VALUES('��������','links.html','{my:friendlinks /}',1,0)")
	'����ҳ��
	Conn.execute("INSERT INTO [DiyPage] ([Title],[PageName],[Code],[State],[IsSystem]) VALUES('��Ʒ����','download.html','<p>--------------</p><p>E��CMS��Ʒ��飺</p><p>Eekku Cms(E��Cms)���ҵĵ�һ��Cms��Ʒ���书���У�</p><p>&nbsp;&nbsp;&nbsp; 1����ϵͳ�ǲ��� ASP + Access ����ʵ��<br />&nbsp;&nbsp;&nbsp;2������Ĺ��������¡���ᡢ���Ժ����۵Ȼ�������<br />&nbsp;&nbsp;&nbsp; 3��������ģ���ASP���뽫100%��ȫ���롣<br />&nbsp;&nbsp;&nbsp;4����ϵͳ�Դ�ϵͳ��ǩ����ǩ�﷨����HTML��ǩ�﷨������׶��������Զ����ǩ���ܡ�<br />&nbsp;&nbsp;&nbsp; 5����ϵͳ���Զ���ҳ�湦�ܡ�<br />&nbsp;&nbsp;&nbsp; 6�����๦�ܵȴ���������....</p><p>Ŀǰ���ǿ����棬�����ע��</p><p>���°汾���أ�<a href=""http://code.google.com/p/foolincms/downloads/list"" target=""_blank"">�����������ҳ��</a></p><p>--------------</p>',1,0)")
	Conn.Close: Set Conn = Nothing
	CreateTable = True
End Function

'�������ݿ�
Function CreateDB(Byval DbName)
	'On Error Resume Next		'�ݴ���
	DbName = "../database/" & DbName
	'�ж����ݿ��Ƿ����
	If ExistFile(DbName)= True Then
		Call DeleteFile(DbName)
	End If
	If Instr(DbName,":\")=0 And Instr(DbName,":/")=0 Then
		DbName = Server.MapPath(DbName)
	End If
	'�ж��Ƿ�������ݿ��ļ���
	If ExistFolder("../database/")= False Then	Call CreateFolder("../database/")
	'�������ݿ�
	Dim Cat: Set Cat=Server.CreateObject("ADOX.Catalog") 
	Cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DbName
	Set Cat = Nothing
	'If CreateTable(DbName) = False Then CreateDB = False: Exit Function	'������
	If Err Then Response.Write(Err): Response.End()
	CreateDB = True
End Function


Function CreateConfig(DbName)
 	Dim strTemp, keyTab, keyEnter
	keyTab = Chr(9) & Chr(9)
	keyEnter = vbcrlf & vbcrlf
	'ϵͳ��Ϣ
	strTemp =  Chr(60) & "%@LANGUAGE=""VBSCRIPT"" CODEPAGE=""936""%" & Chr(62) & Chr(10)
	strTemp = strTemp & Chr(60) & "%" & Chr(10)
	strTemp = strTemp & "'Option Explicit" & keyTab & "'ǿ������" & Chr(10)
	strTemp = strTemp & "On Error Resume Next" & keyTab & "'�ݴ���" & Chr(10)
	strTemp = strTemp & "'=========================================================" & Chr(10)
	strTemp = strTemp & "' File Name��	config.asp" & Chr(10)
	strTemp = strTemp & "' Purpose��		ϵͳ�����ļ�" & Chr(10)
	strTemp = strTemp & "' Auhtor: 		Foolin" & Chr(10)
	strTemp = strTemp & "' E-mail: 		Foolin@126.com" & Chr(10)
	strTemp = strTemp & "' Created on: 	"& Now() & Chr(10)
	strTemp = strTemp & "' Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved" & Chr(10)
	strTemp = strTemp & "'=========================================================" & keyEnter
	' DBPATH����
	strTemp= strTemp & "Dim DBPATH" & keyTab & "'Access���ݿ�·��" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DBPATH = " & Chr(34) & "database/" & DbName & Chr(34) & keyEnter
	' SITENAME����
	strTemp= strTemp & "Dim SITENAME" & keyTab & "'��վ����" & Chr(10) & Chr(9) 
	strTemp= strTemp & "SITENAME = " & Chr(34) & "E�Ṥ����" & Chr(34) & keyEnter
	' HTTPURL����
	strTemp= strTemp & "Dim HTTPURL" & keyTab & "'��վ��ַǰ׺" & Chr(10) & Chr(9) 
	strTemp= strTemp & "HTTPURL = " & Chr(34) & "http://" & Request.ServerVariables("Http_Host") & Chr(34) & keyEnter
	' INSTALLDIR����
	strTemp= strTemp & "Dim INSTALLDIR" & keyTab & "'��վ��װĿ¼����Ŀ¼��Ϊ��/" & Chr(10) & Chr(9) 
	strTemp= strTemp & "INSTALLDIR = " & Chr(34) & Replace(Request.ServerVariables("Path_Info"),"/install/index.asp", "") & Chr(34) & keyEnter
	' SITEKEYWORDS����
	strTemp= strTemp & "Dim SITEKEYWORDS" & keyTab & "'��վ�ؼ���" & Chr(10) & Chr(9) 
	strTemp= strTemp & "SITEKEYWORDS = " & Chr(34) & "E������E��Cms��E�Ṥ����,www.eekku.com���������£�ling.liufu.org" & Chr(34) & keyEnter
	' TEMPLATEDIR����
	strTemp= strTemp & "Dim TEMPLATEDIR" & keyTab & "'��վģ��·�������磺default��ʾtemplate/default/" & Chr(10) & Chr(9) 
	strTemp= strTemp & "TEMPLATEDIR = " & Chr(34) & "default" & Chr(34) & keyEnter
	' ISHIDETEMPPATH����
	strTemp= strTemp & "Dim ISHIDETEMPPATH" & keyTab & "'�Ƿ�����ģ��·�����������Ӱ�������ٶ�" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISHIDETEMPPATH = " & "0" & keyEnter
	' ISOPENGBOOK����
	strTemp= strTemp & "Dim ISOPENGBOOK" & keyTab & "'�Ƿ񿪷����ԣ�Ĭ�Ͽ���" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISOPENGBOOK = 1" & keyEnter
	' ISAUDITGBOOK����
	strTemp= strTemp & "Dim ISAUDITGBOOK" & keyTab & "'�Ƿ���Ҫ������ԣ���-1����-0" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISAUDITGBOOK = 0" & keyEnter
	' GBOOKTIME����
	strTemp= strTemp & "Dim GBOOKTIME" & keyTab & "'�����������ʱ��������λ�룬Ĭ��60��" & Chr(10) & Chr(9) 
	strTemp= strTemp & "GBOOKTIME = 60" keyEnter
	' ISCACHE����
	strTemp= strTemp & "Dim ISCACHE" & keyTab & "'�Ƿ񻺴棬�����ǣ����������������" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISCACHE = " & "1" & keyEnter
	' CACHEFLAG����
	strTemp= strTemp & "Dim CACHEFLAG" & keyTab & "'�����־����������Ӣ����ĸ" & Chr(10) & Chr(9) 
	strTemp= strTemp & "CACHEFLAG = " & Chr(34) & "EekkuCms_" & Chr(34) & keyEnter
	' CACHETIME����
	strTemp= strTemp & "Dim CACHETIME" & keyTab & "'����ʱ�䣬Ĭ����60��" & Chr(10) & Chr(9) 
	strTemp= strTemp & "CACHETIME = " & "60" & keyEnter
	' ISWEBLOG����
	strTemp= strTemp & "Dim ISWEBLOG" & keyTab & "'�Ƿ��¼��̨���������¼" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISWEBLOG = " & "1" & keyEnter
	' LIMITIP����
	strTemp= strTemp & "Dim LIMITIP" & keyTab & "'����IP������|���зָ�" & Chr(10) & Chr(9) 
	strTemp= strTemp & "LIMITIP = " & Chr(34) & "" & Chr(34) & keyEnter
	' DIRTYWORDS����
	strTemp= strTemp & "Dim DIRTYWORDS" & keyTab & "'�໰���ˣ�����|���зָ�" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DIRTYWORDS = " & Chr(34) & "fuck|sex" & Chr(34) & keyEnter
	'��ǽ���
	strTemp = strTemp & "%" & Chr(62) & Chr(10)
	If CreateFile(strTemp, "../inc/config.asp") = True Then
		CreateConfig = True
	Else
		CreateConfig = False
	End If
End Function

Function LockInstall()
 	Dim strTemp, keyTab, keyEnter
	keyTab = Chr(9) & Chr(9)
	keyEnter = vbcrlf & vbcrlf
	'ϵͳ��Ϣ
	strTemp = strTemp & "=========================================================" & Chr(10)
	strTemp = strTemp & " File Name��	install.lock" & Chr(10)
	strTemp = strTemp & " Purpose��		�����ļ�" & Chr(10)
	strTemp = strTemp & " Auhtor: 		Foolin" & Chr(10)
	strTemp = strTemp & " E-mail: 		Foolin@126.com" & Chr(10)
	strTemp = strTemp & " Created on: 	"& Now() & Chr(10)
	strTemp = strTemp & " Copyright (c) 2009 E�Ṥ���ң�Foolin��All Rights Reserved" & Chr(10)
	strTemp = strTemp & "=========================================================" & keyEnter
	Call CreateFile(strTemp, "install.lock")
End Function
%>
