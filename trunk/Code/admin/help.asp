<!--#include file="inc/admin.include.asp"-->
<%
'===========================================
'File Name：	help.asp
'Purpose：		管理帮助文档
'Auhtor: 		Foolin
'Create on:		2009-9-19 17:39:51
'Copyright:		E酷工作室(www.eekku.com)
'===========================================
Call ChkLogin()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>标签代码参考-用户帮助</title>
<link href="css/common.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	font-size: 12px;
}
.clear{ clear:both;}
.subcontainer{
	border:1px #88C4FF solid;
	background:#E6F2FF;
	margin:5px;
}


.tagNav{
	background:#71B8FF;
	line-height:30px;
	height:30px;
	font-size:14px;
	font-weight:bold;
	color:#39F;
	color:#FFF;
}

.tagNav ul{
	margin:0px;
	padding:0px;
	padding-left:20px;
	list-style:none;
}

.tagNav li{
	float:left;
	margin-left:8px;
	margin-right:8px;
	text-align:center;
	padding-bottom:0px;
	cursor:pointer;
}

.tagNav li a{
	color:#FFF;
	padding:8px;
	text-decoration:none;
}

.tagNav li a:hover{
	background:#E6F2FF;
	color:#39F;
	padding:8px;
	text-decoration:none;
}

.tagNav .on a{
	background:#E6F2FF;
	color:#39F;
	padding:8px;
	text-decoration:none;
}

.tagNav .off a{
	color:#FFF;
	padding:8px;
	text-decoration:none;
}
	

.area{
	margin:4px;
	padding:5px;
	font-size:14px;
	background:#FFF;
	line-height:25px;
}
#control{ margin:10px 5px;
}
#view{ margin:10px 5px;}
legend{ font-size:14px; font-weight:bold; color:#09F;}
#tagCode{ color:#F00;}
.intro{ color:#999;}
.codeTipTit{color:#000;width:60%; padding-left:10px;}
-->
</style>
<script type="text/javascript" language="javascript">
<!--
//获取ID元素
function $(o){ return typeof(o)=="string" ? document.getElementById(o) : o;}
//显示节点内容
function S(o){ $(o).style.display = "block";}
//隐藏节点内容
function H(o){ $(0).style.display = "none";}
//获取URL
function GetUrl(ProtoString){
    var paraString =  ProtoString.split('#');
    if(!paraString[1]){
        return null;
    }
    var paras = paraString[1].split('&');
    var allParas = new Array(paras.length);
    for(var i = 0; i<paras.length; i++){
           allParas[GetPara(paras[i])[0]] = GetPara(paras[i])[1];
    }
    return allParas;
}

//获取Para
function GetPara(word){
    if(!word){
        return null;
    }
    var onePara = word.split('=');
    return onePara;
}

//标签显示分析
function HightLightTag(){
	var strList = $("tagCode").innerHTML;	//列表内容
	var listName;	//列表名称
	var regExp;		//正则表达式
	//绿色高亮{tag:name }中的name
	regExp = /\{(.+?):([^\s\]\}]+)/ig;
	strList = strList.replace(regExp, "{$1:<font color='blue'>$2</font>");
	//绿色高亮[list:name]中的name
	regExp = /\[(.+?):([^\s\]]+)/ig;
	strList = strList.replace(regExp, " &nbsp;&nbsp;[<font color='blue'>$1</font>:<font color='green'>$2</font>");
	//蓝色高亮属性名，绿色高亮属性值
	regExp = /\s(\S+)=\"(\S*)\"/ig;
	strList = strList.replace(regExp, "  <font color='blue'>$1</font>=\"<font color='green'>$2</font>\"");
	//绿色高亮IF条件底层标签
	regExp = /\{if:(.+?)\}(.+?)/ig;
	strList = strList.replace(regExp, "{if:$1}<font color='green'>$2");
	regExp = /\{else\}/ig;
	strList = strList.replace(regExp, "</font>{else}<font color='green'>");
	regExp = /\{\/if\}/ig;
	strList = strList.replace(regExp, "</font>{/if}");
	$("tagCode").innerHTML = strList;
}

function onNav(opt)
{
	var _Navs = $("tagNav").getElementsByTagName("li");
	for(var i = 0; i < _Navs.length; i++){
		_Navs[i].className = "off";
	}
	if (opt){
		$("nav" + opt).className = "on";
		return false;
	}
    var url = document.location.href;
    var paras = new Array();
    paras = GetUrl(url);
	if (paras == null){
		$("navList").className = "on";
	}
	else if (paras["tag"]){
		$("nav" + paras["tag"]).className = "on";
	}
	else{
		$("navList").className = "on";
	}
}

-->
</script>
</head>

<body>
<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td colspan="2" id="topNav">
			<%Call TopNav("index")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">管理首页</a></li>
                </ul>
                <ul class="menu">
                	<li class="mTitle">系统帮助</li>
                    <li class="on"><a href="index.asp">帮助首页</a></li>
                    <li><a href="#list">列表</a></li>
                    <li><a  onclick="onNav('Content');">内容</a></li>
                    <li><a onclick="onNav('MyTag');">自定义</a></li>
                    <li><a onclick="onNav('Sys');">系统</a></li>
                    <li><a onclick="onNav('Include');">包含文件</a></li>
                    <li><a onclick="onNav('DiyPage');">自定义页面</a></li>
                    <li><a onclick="onNav('If');">判断</a></li>
                    <li><a onclick="onNav('DbDoc');">数据库表</a></li>
                    <li><a onclick="onNav('TagDoc');">帮助</a></li>
                </ul>
                
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top" height="100%">
                        <div class="subcontainer">
                            <div class="tagNav" id="tagNav">
                                <ul>
                                <li id="navList" class="on"><a onclick="onNav('List');">列表</a></li>
                                <li id="navContent"><a  onclick="onNav('Content');">内容</a></li>
                                <li id="navMytag"><a onclick="onNav('MyTag');">自定义</a></li>
                                <li id="navSys"><a onclick="onNav('Sys');">系统</a></li>
                                <li id="navInclude"><a onclick="onNav('Include');">包含文件</a></li>
                                <li id="navDiyPage"><a onclick="onNav('DiyPage');">自定义页面</a></li>
                                <li id="navIf"><a onclick="onNav('If');">判断</a></li>
                                <li id="navDbDoc"><a onclick="onNav('DbDoc');">数据库表</a></li>
                                <li id="navTagDoc"><a onclick="onNav('TagDoc');">帮助</a></li>
                                </ul>
                            </div>
                            <div class="area">
                                    <div id="control">
                                        <fieldset>
                                        <legend>标签操作选项</legend>
                                        <div id="ctrlOpt">
                                        </div>
                                        </fieldset>
                                    </div>
                                    <div id="view">
                                        <fieldset>
                                        <legend>标签参考代码</legend>
                                        <div id="tagCode">
                                             <div class="codeTipTit">========== 列表标签代码 ==========</div>
                                            {list:myname mode="sql" row="10" Col="1" width="100%" class="list"}<br />
                                            [myname:字段名  len="" lenext=""] <span class="intro">标题，长度</span><br />
                                            [myname:字段名]<br />
                                            {/list:myname}<br />
                                            
                                            <br /><div class="codeTipTit">========== 包含文件标签代码 ==========</div>
                                            {include file="header.html" /}<br />
                                            
                                            <br /><div class="codeTipTit">========== 自定义标签 ==========</div>
                                            {my:自定义标签名 /} <br />
                                            
                                            <br /><div class="codeTipTit">========== 系统标签 ==========</div>
                                            {sys:变量名 /}<br />
                                            
                                            <br /><div class="codeTipTit">========== 内容标签 ==========</div>
                                            {field:字段名 /}<br />
                                            
                                            <br /><div class="codeTipTit">========== 自定义页面标签 ==========</div>
                                            {diypage:字段名 /}<br />
                                            
                                            <br /><div class="codeTipTit">========== 判断标签 ==========</div>
                                            {if:表达式} ### 表达式成立的值 ### {else} ### 表达式不成立的值 ### {/if}
                                        </div>
                                        </fieldset>
                                    </div>
                            </div>
                            <div class="clear"></div>
                        </div>
                        
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>

</div>
</body>
</html>
<script type="text/javascript" language="javascript">
<!--
HightLightTag();
onNav();
-->
</script>
