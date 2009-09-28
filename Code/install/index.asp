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
   <title>E酷内容管理系统-安装</title>
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
            <p>版权所有 (c)2009-2010，E酷工作室 (www.eekku.com) 保留所有权利。 </p>
            <p>本系统由Foolin(负零)独立开发。Author: Foolin &nbsp;&nbsp; QQ: 970026999 E-mail: Foolin@126.com </p>
    </div>
    
</div>

</body>
</html>
<!-- 本安装过程制作：刘付灵   E-mail:Foolin@126.com   2009年7月28日15:14:40 -->

<%Sub Step1()%>

	<!--第一步:许可协议-->
    <div id="step1">
    	<div class="title">第一步：E酷内容管理系统（EekkuCMS）安装许可协议</div>
        <div class="license">
          <p>感谢您选择E酷内容管理系统（以下简称EekkuCMS），EekkuCMS是一个简单又实用的小型站点程序，基于 ASP + Access 的技术开发，全部源码开放。 </p>
          <p>EekkuCMS 的官方网址是： www.eekku.com</p>
          <p>为了使您正确并合法的使用本软件，请您在使用前务必阅读清楚下面的协议条款：</p>
          <p>一、本授权协议适用且仅适用于 EekkuCMS 1.x.x 版本，E酷工作室对本授权协议拥有最终解释权。</p>
          <p>二、协议许可的权利 <br />
            &nbsp;&nbsp;&nbsp;&nbsp; 1、您可以在完全遵守本最终用户授权协议的基础上，将本软件应用于非商业用途，而不必支付软件版权授权费用。 <br />
            &nbsp;&nbsp;&nbsp;&nbsp; 2、您可以在协议规定的约束和限制范围内修改 EekkuCMS 源代码或界面风格以适应您的网站要求。 <br />
            &nbsp;&nbsp;&nbsp;&nbsp;   3、您拥有使用本软件构建的网站全部内容所有权，并独立承担与这些内容的相关法律义务。 <br />
            &nbsp;&nbsp;&nbsp;&nbsp;   4、获得商业授权之后，您可以将本软件应用于商业用途，同时依据所购买的授权类型中确定的技术支持内容，自购买时刻起，在技术支持期限内拥有通过指定的方式获得指定范围内的技术支持服务。商业授权用户享有反映和提出意见的权力，相关意见将被作为首要考虑，但没有一定被采纳的承诺或保证。 </p>
<p><strong>三、协议规定的约束和限制 </strong><br />
   &nbsp;&nbsp;&nbsp;&nbsp; 1、不得将本软件用于国家不允许开设的网站（包括色情、反动、含有病毒，赌博类网站）。<br />
   &nbsp;&nbsp;&nbsp;&nbsp; 2、未经官方许可，不得对本软件或与之关联的商业授权进行出租、出售、抵押或发放子许可证。 <br />
   &nbsp;&nbsp;&nbsp;&nbsp; 3、未经官方许可，禁止在本软件的整体或任何部分基础上以发展任何派生版本、修改版本或第三方版本用于重新分发。 <br />
  &nbsp;&nbsp;&nbsp;&nbsp;  4、如果您未能遵守本协议的条款，您的授权将被终止，所被许可的权利将被收回，并承担相应法律责任。 </p>
<p><strong>四、有限担保和免责声明 </strong><br />
   &nbsp;&nbsp;&nbsp;&nbsp; 1、本软件及所附带的文件是作为不提供任何明确的或隐含的赔偿或担保的形式提供的。 <br />
   &nbsp;&nbsp;&nbsp;&nbsp; 2、用户出于自愿而使用本软件，您必须了解使用本软件的风险，在尚未购买产品技术服务之前，我们不承诺对免费用户提供任何形式的技术支持、使用担保，也不承担任何因使用本软件而产生问题的相关责任。 <br />
   &nbsp;&nbsp;&nbsp;&nbsp; 3、电子文本形式的授权协议如同双方书面签署的协议一样，具有完全的和等同的法律效力。您一旦开始确认本协议并安装本系统，即被视为完全理解并接受本协议的各项条款，在享有上述条款授予的权力的同时，受到相关的约束和限制。协议许可范围以外的行为，将直接违反本授权协议并构成侵权，我们有权随时终止授权，责令停止损害，并保留追究相关责任的权力。 <br />
   &nbsp;&nbsp;&nbsp;&nbsp; 4、如果本软件带有其它软件的整合API示范例子包，这些文件版权不属于本软件官方，并且这些文件是没经过授权发布的，请参考相关软件的使用许可合法的使用。</p>
<p>版权所有 (c)2009-2010，E酷工作室 保留所有权利。 </p>
<p>协议发布时间：  2009年9月21日 By Foolin </p>
<p>版本最新更新：2009年9月21日 By Foolin </p>
        </div>
        <div class="btn">
            <input type="button" value="同意"  onclick="go('step2');"/> &nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="不同意" onclick="go('cancel');" />
        </div>
    </div>
<%
 End Sub
 
 
 Sub Step2()
%>
    <div id="step2">
    	<div class="title">第二步：E酷内容管理系统（EekkuCMS）安装</div>
        <form action="?action=step3" method="post" name="form1" onsubmit="return chkForm();">
        <table class="form">
            <tr>
                <td class="name">登录帐号：</td><td class="inputtxt"><input name="Username" type="text" /> 请填写管理员账号</td>
            </tr>
            <tr>
                <td class="name">登录密码：</td><td  class="inputtxt"><input name="Password" type="password" /> 填写管理员密码</td>
            </tr>
            <tr>
                <td class="name">重复密码：</td><td  class="inputtxt"><input name="RePassword" type="password" /></td>
            </tr>
            <tr>
                <td class="name">数据库名称：</td><td  class="inputtxt"><input name="DbName" type="text" value="Fl<%=Day(Now())%>#Ek_<%=Left(UCase(MD5(Now())),10)%>.mdb" /> 请填写数据库名称</td>
            </tr>
            <tr>
                <td colspan="2" class="btn"> <input type="button" value="上一步" onclick="go('step1');" /> &nbsp;&nbsp;&nbsp;&nbsp;
        		<input type="submit" value="下一步" /></td>
            </tr>
        </table>
        </form>
    </div>
<script language="javascript">
<!--
function chkForm(){
	var form = document.forms["form1"];
	if( form.elements["Username"].value == ""){
		alert("账号不能为空");
		return false;
	}
	if( form.elements["Password"].value == ""){
		alert("密码不能为空");
		return false;
	 }
	if( form.elements["Password"].value.length < 6){
		alert("密码不能少于6位");
		return false;
	 }
	if( form.elements["RePassword"].value != form.elements["Password"].value){
		alert("两次密码不一致！");
		return false;
	 }
	if( form.elements["DbName"].value.length < 6){
		alert("数据库名（包括后缀）不能少于6位!");
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
    		<div class="title">第三步：完成本系统的安装</div>
			<div class="state">
            <dl>
            	<dt> 系统安装成功</dt>
            	<dd> 帐号：<span class="success"><%=Request("Username")%></span></dd>
                <dd> 密码：<span class="success"><%=Request("Password")%></span></dd>
            	<dd> 当前安装目录： <%=Replace(Request.ServerVariables("Path_Info"),"/install/index.asp", "")%> </dd>
                <dd> <a href="../index.asp">进入首页</a></dd>
                <dd> <a href="../admin/login.asp">进入后台管理</a> </dd>
              <dd>&nbsp; </dd>
                <dd>   <span class="warn">！注意：</span> 如果出现有安装错误，请按照安装说明书进行自行配置系统。</dd>
               <dd> <span class="warn">建议：</span>为了本系统安全，防止外人入侵，建议立刻把Install目录重命名（名字越少人猜到越好），或者直接删除。非常感谢您使用本系统！祝你旅途愉快！</dd>
                <dd class="btn"> <input type="button" value="完成" onclick="go('hasinstall');"   /></dd>
            </dl>
            </div>
    </div>
    
<%
 End Sub
 
 
 Sub Cancel()
%>
    <div id="cancel">
    	<div class="title">您已经取消了本次系统的安装！</div>
        <div class="cancel">
        	<div class="ourinfo">
            &nbsp;&nbsp;<b>Author: Foolin(负零) </b><br />
            	&nbsp;&nbsp;&nbsp;&nbsp; QQ：970026999<br />
                &nbsp;&nbsp;&nbsp;&nbsp; E-mail：Foolin@126.com<br /><br />
                &nbsp;&nbsp;&nbsp;&nbsp; Home Page: http://ling.liufu.org<br /><br />
            网址：http://www.eekku.com （E酷工作室）<br /><br />
            本系统由Foolin独立开发。如果有任何问题或者建议，请联系作者，万分感谢！<br />
            </div>
        </div>
    </div>
    
<%
 End Sub
 
 
 Sub HasInstall()
%>
    
    <div id="hasInstall">
    	<div class="title">你已经安装了本系统！</div>
        <div class="box">
            <div class="error">如果需要重新安装，请手工删除Install/目录下的Install.lock文件，然后【<a href="index.asp">点击这里</a>】进行安装！</div>
           <p>
            	>> <a href="../index.asp">进入首页</a> <br />
              	>> <a href="../admin/login.asp">进入后台管理</a>
                
                <div class="warn">!注意：如果打开首页出错，请自行按照说明进行手动配置inc/config.asp。</div>
           </p>
        	<br />
        	<div class="ourinfo">
            &nbsp;&nbsp;<b>Author: Foolin(负零) </b><br />
            	&nbsp;&nbsp;&nbsp;&nbsp; QQ：970026999<br />
                &nbsp;&nbsp;&nbsp;&nbsp; E-mail：Foolin@126.com<br /><br />
                &nbsp;&nbsp;&nbsp;&nbsp; Home Page: http://ling.liufu.org<br /><br />
            网址：http://www.eekku.com （E酷工作室）<br /><br />
            本系统由Foolin独立开发。如果有任何问题或者建议，请联系作者，万分感谢！<br />
            </div>
        </div>
    </div>
    
<%
 End Sub
 

'检查安装函数
Function ChkInstall()
	If ExistFile("install.lock") Then
		Response.Redirect("?action=hasinstall")
		ChkInstall = True
	End If
	ChkInstall = True
End Function

Function Install()
	Call ChkInstall()	'检查是否已经安装
	Dim strUsername, strPassword
	strUsername = Request("Username")
	strPassword = Request("Password")
	If Len(strUsername) < 3 Then Call MsgBox("管理员用户名不能少于三个字符","BACK")
	If Len(strPassword) < 6 Then Call MsgBox("密码不能少于6个字符","BACK")
	If Request("RePassword")<>strPassword Then Call MsgBox("两次密码不一致！","BACK")
	If Len(Request("DbName")) < 6 Then Call MsgBox("为了数据库安全，数据库名不能少于6个字符","BACK")
	If (Instr(Request("DbName"),"/") > 0) Or (Instr(Request("DbName"),"\") > 0) Then
		Call MsgBox("数据库名称请不要出现“/”或“\”这样的路径名","BACK")
	End If
	

	strPassword = MD5(strPassword)
	'创建数据库
	Call CreateDB(Request("DbName"))
	'创建数据库表	
	Call CreateTable(Request("DbName"), strUsername, strPassword)	
	'生成配置文件
	Call CreateConfig(Request("DbName"))
	'锁定安装文件	
	Call LockInstall()	
	Install = True
End Function

'导入SQL语句
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
	'创建管理员初始密码
	Conn.execute("INSERT INTO [Admin] ([Username],[Nickname],[Password],[Level],[LoginCount],[LoginTime],[LoginIP]) VALUES('"& strUsername &"','"& strUsername &"','"& strPassword &"',3,0,'"& Now() &"','"& GetIP() &"')")
	'创建初始化自定义标签
	Conn.execute("INSERT INTO [MyTags] ([Name],[Info],[Code]) VALUES('FirstTag','第一个自定义标签','第一个自定义标签：(零星碎事)')")
	'系统信息
	Conn.execute("INSERT INTO [MyTags] ([Name],[Info],[Code]) VALUES('AboutSys','关于系统信息','{sys:sys /}<br />作者：Foolin<br /> Email: Foolin@126.com <br /> 主页：http://www.LiuFu.org/Ling<br />官网：Http://www.eekku.com<br />')")
	'自定义标签：友情链接
	Conn.execute("INSERT INTO [MyTags] ([Name],[Info],[Code]) VALUES('FriendLinks','友情链接','<a href=""http://www.eekku.com""> -==- E酷网 -==- </a><br />"& Chr(10) & Chr(9) &"<a href=""http://www.liufu.org/ling/""> -==- 零星碎事 -==- </a><br />"& Chr(10) & Chr(9) &"')")
	'帮助页面
	Conn.execute("INSERT INTO [DiyPage] ([Title],[PageName],[Code],[State],[IsSystem]) VALUES('帮助文档','help.html','<p>这是我的第一个自定义页面，欢迎大家光临。</p><p>你只要在后台添加一个页面，然后引用该链接：<font color=""red"">diypage.asp?id=<font color=""blue"">[您建立页面的ID]</font></font>或者<font color=""red"">diypage.asp?url=<font color=""blue"">[您建立页面的名称]</font></font>即可打开这个自定义页面！</p><p>如需帮助，请访问：http://www.eekku.com，E酷网络欢迎你。</p><p>你正在使用E酷CMS系统！</p><p><br><br><a href=""http://www.eekku.com""> -==- E酷网 -==- </a></p><p><br /></p><p><a href=""http://www.liufu.org/ling/""> -==- 零星碎事 -==- </a><br /></p>',1,0)")
	'友情链接
	Conn.execute("INSERT INTO [DiyPage] ([Title],[PageName],[Code],[State],[IsSystem]) VALUES('友情链接','links.html','{my:friendlinks /}',1,0)")
	'下载页面
	Conn.execute("INSERT INTO [DiyPage] ([Title],[PageName],[Code],[State],[IsSystem]) VALUES('作品下载','download.html','<p>--------------</p><p>E酷CMS作品简介：</p><p>Eekku Cms(E酷Cms)是我的第一个Cms作品，其功能有：</p><p>&nbsp;&nbsp;&nbsp; 1、本系统是采用 ASP + Access 技术实现<br />&nbsp;&nbsp;&nbsp;2、程序的功能有文章、相册、留言和评论等基本功能<br />&nbsp;&nbsp;&nbsp; 3、程序中模板和ASP代码将100%完全分离。<br />&nbsp;&nbsp;&nbsp;4、本系统自带系统标签，标签语法类似HTML标签语法，简洁易懂，还有自定义标签功能。<br />&nbsp;&nbsp;&nbsp; 5、本系统有自定义页面功能。<br />&nbsp;&nbsp;&nbsp; 6、更多功能等待你来发现....</p><p>目前还是开发版，敬请关注！</p><p>最新版本下载：<a href=""http://code.google.com/p/foolincms/downloads/list"" target=""_blank"">点击进入下载页面</a></p><p>--------------</p>',1,0)")
	Conn.Close: Set Conn = Nothing
	CreateTable = True
End Function

'创建数据库
Function CreateDB(Byval DbName)
	'On Error Resume Next		'容错处理
	DbName = "../database/" & DbName
	'判断数据库是否存在
	If ExistFile(DbName)= True Then
		Call DeleteFile(DbName)
	End If
	If Instr(DbName,":\")=0 And Instr(DbName,":/")=0 Then
		DbName = Server.MapPath(DbName)
	End If
	'判断是否存在数据库文件件
	If ExistFolder("../database/")= False Then	Call CreateFolder("../database/")
	'创建数据库
	Dim Cat: Set Cat=Server.CreateObject("ADOX.Catalog") 
	Cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DbName
	Set Cat = Nothing
	'If CreateTable(DbName) = False Then CreateDB = False: Exit Function	'创建表
	If Err Then Response.Write(Err): Response.End()
	CreateDB = True
End Function


Function CreateConfig(DbName)
 	Dim strTemp, keyTab, keyEnter
	keyTab = Chr(9) & Chr(9)
	keyEnter = vbcrlf & vbcrlf
	'系统信息
	strTemp =  Chr(60) & "%@LANGUAGE=""VBSCRIPT"" CODEPAGE=""936""%" & Chr(62) & Chr(10)
	strTemp = strTemp & Chr(60) & "%" & Chr(10)
	strTemp = strTemp & "'Option Explicit" & keyTab & "'强制声明" & Chr(10)
	strTemp = strTemp & "On Error Resume Next" & keyTab & "'容错处理" & Chr(10)
	strTemp = strTemp & "'=========================================================" & Chr(10)
	strTemp = strTemp & "' File Name：	config.asp" & Chr(10)
	strTemp = strTemp & "' Purpose：		系统配置文件" & Chr(10)
	strTemp = strTemp & "' Auhtor: 		Foolin" & Chr(10)
	strTemp = strTemp & "' E-mail: 		Foolin@126.com" & Chr(10)
	strTemp = strTemp & "' Created on: 	"& Now() & Chr(10)
	strTemp = strTemp & "' Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved" & Chr(10)
	strTemp = strTemp & "'=========================================================" & keyEnter
	' DBPATH变量
	strTemp= strTemp & "Dim DBPATH" & keyTab & "'Access数据库路径" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DBPATH = " & Chr(34) & "database/" & DbName & Chr(34) & keyEnter
	' SITENAME变量
	strTemp= strTemp & "Dim SITENAME" & keyTab & "'网站名称" & Chr(10) & Chr(9) 
	strTemp= strTemp & "SITENAME = " & Chr(34) & "E酷工作室" & Chr(34) & keyEnter
	' HTTPURL变量
	strTemp= strTemp & "Dim HTTPURL" & keyTab & "'网站网址前缀" & Chr(10) & Chr(9) 
	strTemp= strTemp & "HTTPURL = " & Chr(34) & "http://" & Request.ServerVariables("Http_Host") & Chr(34) & keyEnter
	' INSTALLDIR变量
	strTemp= strTemp & "Dim INSTALLDIR" & keyTab & "'网站安装目录，根目录则为：/" & Chr(10) & Chr(9) 
	strTemp= strTemp & "INSTALLDIR = " & Chr(34) & Replace(Request.ServerVariables("Path_Info"),"/install/index.asp", "") & Chr(34) & keyEnter
	' SITEKEYWORDS变量
	strTemp= strTemp & "Dim SITEKEYWORDS" & keyTab & "'网站关键词" & Chr(10) & Chr(9) 
	strTemp= strTemp & "SITEKEYWORDS = " & Chr(34) & "E酷网，E酷Cms，E酷工作室,www.eekku.com，零星碎事，ling.liufu.org" & Chr(34) & keyEnter
	' TEMPLATEDIR变量
	strTemp= strTemp & "Dim TEMPLATEDIR" & keyTab & "'网站模板路径，例如：default表示template/default/" & Chr(10) & Chr(9) 
	strTemp= strTemp & "TEMPLATEDIR = " & Chr(34) & "default" & Chr(34) & keyEnter
	' ISHIDETEMPPATH变量
	strTemp= strTemp & "Dim ISHIDETEMPPATH" & keyTab & "'是否隐藏模板路径，隐藏则会影响载入速度" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISHIDETEMPPATH = " & "0" & keyEnter
	' ISOPENGBOOK变量
	strTemp= strTemp & "Dim ISOPENGBOOK" & keyTab & "'是否开放留言，默认开放" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISOPENGBOOK = 1" & keyEnter
	' ISAUDITGBOOK变量
	strTemp= strTemp & "Dim ISAUDITGBOOK" & keyTab & "'是否需要审核留言，是-1，否-0" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISAUDITGBOOK = 0" & keyEnter
	' GBOOKTIME变量
	strTemp= strTemp & "Dim GBOOKTIME" & keyTab & "'允许留言最短时间间隔，单位秒，默认60秒" & Chr(10) & Chr(9) 
	strTemp= strTemp & "GBOOKTIME = 60" keyEnter
	' ISCACHE变量
	strTemp= strTemp & "Dim ISCACHE" & keyTab & "'是否缓存，建议是，减轻服务器负载量" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISCACHE = " & "1" & keyEnter
	' CACHEFLAG变量
	strTemp= strTemp & "Dim CACHEFLAG" & keyTab & "'缓存标志，可以任意英文字母" & Chr(10) & Chr(9) 
	strTemp= strTemp & "CACHEFLAG = " & Chr(34) & "EekkuCms_" & Chr(34) & keyEnter
	' CACHETIME变量
	strTemp= strTemp & "Dim CACHETIME" & keyTab & "'缓存时间，默认是60分" & Chr(10) & Chr(9) 
	strTemp= strTemp & "CACHETIME = " & "60" & keyEnter
	' ISWEBLOG变量
	strTemp= strTemp & "Dim ISWEBLOG" & keyTab & "'是否记录后台管理操作记录" & Chr(10) & Chr(9) 
	strTemp= strTemp & "ISWEBLOG = " & "1" & keyEnter
	' LIMITIP变量
	strTemp= strTemp & "Dim LIMITIP" & keyTab & "'限制IP，多用|进行分割" & Chr(10) & Chr(9) 
	strTemp= strTemp & "LIMITIP = " & Chr(34) & "" & Chr(34) & keyEnter
	' DIRTYWORDS变量
	strTemp= strTemp & "Dim DIRTYWORDS" & keyTab & "'脏话过滤，多用|进行分割" & Chr(10) & Chr(9) 
	strTemp= strTemp & "DIRTYWORDS = " & Chr(34) & "fuck|sex" & Chr(34) & keyEnter
	'标记结束
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
	'系统信息
	strTemp = strTemp & "=========================================================" & Chr(10)
	strTemp = strTemp & " File Name：	install.lock" & Chr(10)
	strTemp = strTemp & " Purpose：		锁定文件" & Chr(10)
	strTemp = strTemp & " Auhtor: 		Foolin" & Chr(10)
	strTemp = strTemp & " E-mail: 		Foolin@126.com" & Chr(10)
	strTemp = strTemp & " Created on: 	"& Now() & Chr(10)
	strTemp = strTemp & " Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved" & Chr(10)
	strTemp = strTemp & "=========================================================" & keyEnter
	Call CreateFile(strTemp, "install.lock")
End Function
%>
