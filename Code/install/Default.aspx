<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="Install_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>E酷校报管理系统-安装</title>
<style type="text/css">
<!--
body{
	font-family:"微软雅黑",Georgia, "Times New Roman", Times, serif;
	font-size:14px;
}

.title{
	font-size:26px;
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
	font-size:16px;
}


#step1{ display:none;}
.license {
	margin:0px auto;
	line-height:22px;
	height:450px;
	width:70%;
	padding:10px 20px;
	border:#EBEBEB dotted solid;
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


#step2{}

.form{
	margin:0px auto;
	border:#EBEBEB 5px solid;
}
table.form { width:550px;}
table.form tr{background:#F3F3F3;}
table.form td{ padding:5px;}
td.name { text-align:right;}
td.inputtxt { width:75%; text-align:left; color:#666;}
.inputtxt input{ width:250px; height:22px; line-height:22px;}

-->
</style>
</head>
<body>
    <form id="form1" runat="server">
    <div class="title">E酷校报管理系统安装</div>
    <div id="step1">
        <div class="license">
            <p>感谢您选择E酷校报管理系统，本系统由刘付灵、朱松辉两人基于ASP.Net+MSSQL 技术开发。</p>
            <p>为了使你正确并合法的使用本软件，请你在使用前务必阅读清楚下面的协议条款：</p>
            <p><strong>一、本协议仅适用于E酷校报管理系统1.x.x 版本，E酷工作室对本协议有最终解释权。 </strong></p>
            <p><strong>二、协议许可的权利 </strong><br />
              1、您可以在完全遵守本最终用户授权协议的基础上，将本软件应用于商业或非商业用途。 <br />
              2、您可以在协议规定的约束和限制范围内修改 本系统 源代码或界面风格以适应您的网站要求。 <br />
              3、您拥有使用本软件构建的网站全部内容所有权，并独立承担与这些内容的相关法律义务。 <br />
              4、获得商业授权之后，您可以依据所购买的授权类型中确定的技术支持内容，自购买时刻起，在技术支持期限内拥有通过指定的方式获得指定范围内的技术支持服务。商业授权用户享有反映和提出意见的权力，相关意见将被作为首要考虑，但没有一定被采纳的承诺或保证。 </p>
            <p><strong>二、协议规定的约束和限制 </strong><br />
              1、不得将本软件用于国家不允许开设的网站（包括色情、反动、含有病毒，赌博类网站）。<br />
              2、未经官方许可，不得对本软件或与之关联的商业授权进行出租、出售、抵押或发放子许可证。 <br />
              3、未经官方许可，禁止在本软件的整体或任何部分基础上以发展任何派生版本、修改版本或第三方版本用于重新分发。 <br />
              4、如果您未能遵守本协议的条款，您的授权将被终止，所被许可的权利将被收回，并承担相应法律责任。 </p>
            <p><strong>三、有限担保和免责声明 </strong><br />
              1、本软件及所附带的文件是作为不提供任何明确的或隐含的赔偿或担保的形式提供的。 <br />
              2、用户出于自愿而使用本软件，您必须了解使用本软件的风险，在尚未购买产品技术服务之前，我们不承诺对免费用户提供任何形式的技术支持、使用担保，也不承担任何因使用本软件而产生问题的相关责任。 <br />
              3、电子文本形式的授权协议如同双方书面签署的协议一样，具有完全的和等同的法律效力。您一旦开始确认本协议并安装本系统，即被视为完全理解并接受本协议的各项条款，在享有上述条款授予的权力的同时，受到相关的约束和限制。协议许可范围以外的行为，将直接违反本授权协议并构成侵权，我们有权随时终止授权，责令停止损害，并保留追究相关责任的权力。 <br />
              4、如果本软件带有其它软件的整合API示范例子包，这些文件版权不属于本软件官方，并且这些文件是没经过授权发布的，请参考相关软件的使用许可合法的使用。</p>
            <p>版权所有 (c)2009-2010，E酷工作室 保留所有权利。 </p>
            <p>协议发布时间：  2009年7月26日 By Foolin (www.eekku.com) </p>
        </div>
        <div class="btn">
            <input type="button" value="同意" /> &nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="不同意" />
        </div>
    </div>
    
    <div id="step2">
        <form action="" method="post">
        <table class="form">
            <tr>
                <td class="name">登录帐号：</td><td class="inputtxt"><input name="username" type="text" /></td>
            </tr>
            <tr>
                <td class="name">登录密码：</td><td  class="inputtxt"><input name="passsword" type="text" /></td>
            </tr>
            <tr>
                <td class="name">数据库服务器：</td><td class="inputtxt"><input name="DBserver" type="text" /> (包括端口)</td>
            </tr>
            <tr>
                <td class="name">数据库名：</td><td  class="inputtxt"><input name="DBname" type="text" /> (请先建数据库)</td>
            </tr>
            <tr>
                <td class="name">数据库用户名：</td><td  class="inputtxt"><input name="DBuser" type="text" /></td>
            </tr>
            <tr>
                <td class="name">数据库密码：</td><td  class="inputtxt"><input name="DBpwd" type="text" /></td>
            </tr>
            <tr>
                <td colspan="2" class="btn"> <input type="button" value="上一步" /> &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="button" value="下一步" /></td>
            </tr>
        </form>
    </div>
        <asp:Label ID="State" runat="server" Text="Label"></asp:Label>
    
    
    </form>
</body>
</html>
