<!--#include file="inc/admin.include.asp"-->
<%
'===========================================
'File Name��	help.asp
'Purpose��		��������ĵ�
'Auhtor: 		Foolin
'Create on:		2009-9-19 17:39:51
'Copyright:		E�Ṥ����(www.eekku.com)
'===========================================
Call ChkLogin()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ǩ����ο�-�û�����</title>
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
//��ȡIDԪ��
function $(o){ return typeof(o)=="string" ? document.getElementById(o) : o;}
//��ʾ�ڵ�����
function S(o){ $(o).style.display = "block";}
//���ؽڵ�����
function H(o){ $(0).style.display = "none";}
//��ȡURL
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

//��ȡPara
function GetPara(word){
    if(!word){
        return null;
    }
    var onePara = word.split('=');
    return onePara;
}

//��ǩ��ʾ����
function HightLightTag(){
	var strList = $("tagCode").innerHTML;	//�б�����
	var listName;	//�б�����
	var regExp;		//������ʽ
	//��ɫ����{tag:name }�е�name
	regExp = /\{(.+?):([^\s\]\}]+)/ig;
	strList = strList.replace(regExp, "{$1:<font color='blue'>$2</font>");
	//��ɫ����[list:name]�е�name
	regExp = /\[(.+?):([^\s\]]+)/ig;
	strList = strList.replace(regExp, " &nbsp;&nbsp;[<font color='blue'>$1</font>:<font color='green'>$2</font>");
	//��ɫ��������������ɫ��������ֵ
	regExp = /\s(\S+)=\"(\S*)\"/ig;
	strList = strList.replace(regExp, "  <font color='blue'>$1</font>=\"<font color='green'>$2</font>\"");
	//��ɫ����IF�����ײ��ǩ
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
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                	<li class="mTitle">ϵͳ����</li>
                    <li class="on"><a href="index.asp">������ҳ</a></li>
                    <li><a href="#list">�б�</a></li>
                    <li><a  onclick="onNav('Content');">����</a></li>
                    <li><a onclick="onNav('MyTag');">�Զ���</a></li>
                    <li><a onclick="onNav('Sys');">ϵͳ</a></li>
                    <li><a onclick="onNav('Include');">�����ļ�</a></li>
                    <li><a onclick="onNav('DiyPage');">�Զ���ҳ��</a></li>
                    <li><a onclick="onNav('If');">�ж�</a></li>
                    <li><a onclick="onNav('DbDoc');">���ݿ��</a></li>
                    <li><a onclick="onNav('TagDoc');">����</a></li>
                </ul>
                
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top" height="100%">
                        <div class="subcontainer">
                            <div class="tagNav" id="tagNav">
                                <ul>
                                <li id="navList" class="on"><a onclick="onNav('List');">�б�</a></li>
                                <li id="navContent"><a  onclick="onNav('Content');">����</a></li>
                                <li id="navMytag"><a onclick="onNav('MyTag');">�Զ���</a></li>
                                <li id="navSys"><a onclick="onNav('Sys');">ϵͳ</a></li>
                                <li id="navInclude"><a onclick="onNav('Include');">�����ļ�</a></li>
                                <li id="navDiyPage"><a onclick="onNav('DiyPage');">�Զ���ҳ��</a></li>
                                <li id="navIf"><a onclick="onNav('If');">�ж�</a></li>
                                <li id="navDbDoc"><a onclick="onNav('DbDoc');">���ݿ��</a></li>
                                <li id="navTagDoc"><a onclick="onNav('TagDoc');">����</a></li>
                                </ul>
                            </div>
                            <div class="area">
                                    <div id="control">
                                        <fieldset>
                                        <legend>��ǩ����ѡ��</legend>
                                        <div id="ctrlOpt">
                                        </div>
                                        </fieldset>
                                    </div>
                                    <div id="view">
                                        <fieldset>
                                        <legend>��ǩ�ο�����</legend>
                                        <div id="tagCode">
                                             <div class="codeTipTit">========== �б��ǩ���� ==========</div>
                                            {list:myname mode="sql" row="10" Col="1" width="100%" class="list"}<br />
                                            [myname:�ֶ���  len="" lenext=""] <span class="intro">���⣬����</span><br />
                                            [myname:�ֶ���]<br />
                                            {/list:myname}<br />
                                            
                                            <br /><div class="codeTipTit">========== �����ļ���ǩ���� ==========</div>
                                            {include file="header.html" /}<br />
                                            
                                            <br /><div class="codeTipTit">========== �Զ����ǩ ==========</div>
                                            {my:�Զ����ǩ�� /} <br />
                                            
                                            <br /><div class="codeTipTit">========== ϵͳ��ǩ ==========</div>
                                            {sys:������ /}<br />
                                            
                                            <br /><div class="codeTipTit">========== ���ݱ�ǩ ==========</div>
                                            {field:�ֶ��� /}<br />
                                            
                                            <br /><div class="codeTipTit">========== �Զ���ҳ���ǩ ==========</div>
                                            {diypage:�ֶ��� /}<br />
                                            
                                            <br /><div class="codeTipTit">========== �жϱ�ǩ ==========</div>
                                            {if:���ʽ} ### ���ʽ������ֵ ### {else} ### ���ʽ��������ֵ ### {/if}
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
