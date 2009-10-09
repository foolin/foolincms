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
Dim act: act = Request("action")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>标签代码参考-用户帮助</title>
<link href="images/common.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
body {
	font-size: 12px;
	font-family:"Times New Roman", Times, serif;
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
	font-size:12px;
	background:#FFF;
	line-height:25px;
}
#control{ margin:5px;
}
#control table{ border:dashed 1px #EEE; border-collapse:collapse;}
#control table td{ border:dashed 1px #EEE; padding:2px 5px;}
#control form{ margin:2px;}
#view{ margin:10px 5px;}
legend{ font-size:14px; font-weight:bold; color:#09F;}
fieldset{ font-size:14px;}
#tagCode{ color:#F00; font-size:14px;}
#tagCode a{ color:#099;}
#tagCode a:hover{ color:#F00;}
.intro{ color:#999;}
.codeTipTit{color:#000;width:60%; padding-left:10px;}
.btn{ border:#C4E1FF 1px solid; background:#F4FAFF; padding:5px; color:#090; font-weight:bold;}
.blue{ color:blue;}
.green{ color:green;}
.gray{ color:gray;}
.red{ color:red;}
.black{ color:#000;}
h3{ margin:3px;}
.desc{ color:#999; padding-left:5px; font-size:13px;}
-->
</style>
<script type="text/javascript" language="javascript">
<!--
//获取ID元素
function $(o){ return typeof(o)=="string" ? document.getElementById(o) : o;}
//标签高亮
function HightLightTag(id){
	var strList = $(id).innerHTML;//列表内容
	
	var listName;	//列表名称
	var regExp;		//正则表达式
	//绿色高亮{tag:name }中的name
	regExp = /\{(.+?):([^\s\]\}]+)/ig;
	strList = strList.replace(regExp, "{$1:<font color='blue'>$2</font>");
	//绿色高亮[list:name]中的name
	//regExp = /\[(.+?):([^\s\]]+)\]/ig;
	regExp = /\[(.+?):(.+?)\]/ig;
	strList = strList.replace(regExp, " [<font color='blue'>$1</font>:<font color='green'>$2</font>\]");
	//蓝色高亮属性名，绿色高亮属性值
	//alert(strList);
	//regExp = /\s(\S+)=\"(\S*)\"/ig;
	regExp = /\s(\S+)=\"(.+?)\"/ig;
	strList = strList.replace(regExp, " <font color='blue'>$1</font>=\"<font color='green'>$2</font>\"");
	//alert(strList);
	//绿色高亮IF条件底层标签
	regExp = /\{if:(.+?)\}(.+?)/ig;
	strList = strList.replace(regExp, "{if:$1}<font color='green'>$2");
	regExp = /\{else\}/ig;
	strList = strList.replace(regExp, "</font>{else}<font color='green'>");
	regExp = /\{\/if\}/ig;
	strList = strList.replace(regExp, "</font>{/if}");
	//灰色高亮解析说明
	regExp = /&lt;!--(.+?)--&gt;/ig;
	strList = strList.replace(regExp, " &nbsp;<span class='desc'>&lt;!--$1--&gt;</span>");
	//alert(strList);
	$(id).innerHTML = strList;
}
//HightLightTag();
-->
</script>
</head>

<body>

<div id="wrapper">
	<%Call Header()%>
    <table id="container">
        <tr><td colspan="2" id="topNav">
			<%Call TopNav("help")%>
        </td></tr>
        <tr>
            <td id="sidebar" valign="top">
            	<ul class="menu">
                	<li><a href="index.asp">管理首页</a></li>
                </ul>
                <ul class="menu">
                	<li class="mTitle">帮助文章</li>
                    <li <%If Len(act) = 0 Or act = "index" Then Echo(" class=""on""")%>><a href="help.asp">帮助首页</a></li>
                    <li <%If act = "list" Then Echo(" class=""on""")%>><a href="?action=list">列表标签</a></li>
                    <li <%If act = "content" Then Echo(" class=""on""")%>><a href="?action=content">内容标签</a></li>
                    <li <%If act = "mytag" Then Echo(" class=""on""")%>><a href="?action=mytag">自定义标签</a></li>
                    <li <%If act = "sys" Then Echo(" class=""on""")%>><a href="?action=sys">系统标签</a></li>
                    <li <%If act = "include" Then Echo(" class=""on""")%>><a href="?action=include">包含文件标签</a></li>
                    <li <%If act = "diypage" Then Echo(" class=""on""")%>><a href="?action=diypage">DIY页面标签</a></li>
                    <li <%If act = "if" Then Echo(" class=""on""")%>><a href="?action=if">判断标签</a></li>
                    <li <%If act = "doorder" Then Echo(" class=""on""")%>><a href="?action=doorder">标签执行顺序</a></li>
                    <li><a href="http://www.liufu.org/ling" target="_blank">更多帮助</a></li>
                </ul>
                
                <%Call SysInfo()%>
                
            </td>
            <td id="content" valign="top" height="100%">
                      <div class="subcontainer">
                         <div class="area">
						<% 
							Select Case LCase(act)
								Case "index"
									Call Hindex()
								Case "list"
									Call Hlist()
								Case "content"
									Call Hcontent()
								Case "mytag"
									Call Hmytag()
								Case "sys"
									Call Hsys()
								Case "include"
									Call Hinclude()
								Case "diypage"
									Call Hdiypage()
								Case "if"
									Call Hif()
								Case "doorder"
									Call Hdoorder()
								Case Else
									Call Hindex()
							End Select
                        %>
                        </div>
                      </div>
                        
            </td>
        </tr>
    </table>
	<%Call Footer()%>
</div>
</body>
</html>


<%
 '首页内容
 Sub Hindex()
%>
 <div id="view">
    <fieldset>
    <legend>制作模板简单说明</legend>
    	<div style="padding:5px;">
<ol>
<li>制作的模板，你可以创建一个新目录（可自由命名，例如命名为newtpl），但该目录必须放在template目录下。</li>
<li>模板目录中，必须存在的模板页有：index.html（首页模板页）、artlist.html（文章列表模板页）、article.html（文章内容模板页）、piclist.html（图片列表模板页）、picture.html（图片内容模板页）、guestbook.html（留言模板页）和diypage.html（自定义页面模板页）。</li>
<li>模板里面的图片必须放在当前模板（例如：newtpl）中images目录下，css文件必须放在css目录下，js文件必须放在js或者scripts目录下。</li>
<li>在模板中使用的标签，标签语法请查看标签说明文档。<br />
<div style="font-size:12px; color:gray;">

本系统标签尽量采用与HTML接近的语法，但为了与HTML区分，本系统标签采用大括号{}代替HTML中的<>。<br />
例如：<br />
{list：name mode="default" src="article"}<br />
###内层循环标签###<br />
{/list: name }<br />
如果标签有多种写法，请自己选择适合标签，建议用第一个，或者与HTML语法（具有开始与闭合）相近写法。<br />
</div></li>
<li>制作模板完成之后，进入【管理后台】 → 【<a href="admin_config.asp">系统配置</a>】 → 选择【<a href="admin_config.asp">模板</a>】，选中你制作模板的目录，点击保存即可。如果你网站不能及时刷新，点击【<a href="index.asp?action=clearcache">更新缓存</a>】，然后刷新即可完成。</li>
<li>	制作过程可以参照系统自带的模板和本标签说明即可。</li>
<li>	若有任何疑问或者bug,请到官方http://www.eekku.com或者发送邮件到Foolin@126.com进行反馈。</li>
</ol>

<br />
注意：制作模板必须具备一定的HTML和CSS知识。网页默认是Gb2312编码。
<br /><br />

官方：<a href="http://www.eekku.com" target="_blank">http://www.eekku.com</a><br />
主页：<a href="http://www.liufu.org/ling" target="_blank">http://www.liufu.org/ling</a><br />
邮箱：Foolin@126.com<br />
<br />

        </div>
    </fieldset>
    
    <fieldset>
    <legend>标签执行顺序</legend>
    
	&nbsp;&nbsp; 包含标签 → 自定义标签 → 系统标签 → 列表标签 → 分页标签 → 判断标签

    </fieldset>
</div>
                                    
<script type="text/javascript" language="javascript">
<!--
//HightLightTag("tagCode");
-->
</script>
<%End Sub%>




<%
 '列表帮助
 Sub Hlist()
 Dim mode: mode = LCase(Request("mode"))
	If Len(mode) = 0 Then mode = "default"
%>
<div id="control">
    <fieldset>
    <legend>标签操作选项</legend>
    <div id="ctrlOpt">
	<form action="tags.asp?action=create" method="post" name="formList" target="showTags">
    <table style="color:green;">
    	<tr><td>模式（mode）：</td>
            <td>
            <select name="ListMode" onchange="changeMode(this);">
              <option value="default" <%If mode="default" Then Echo("selected=""selected""")%>>默认模式</option>
              <option value="table" <%If mode="table" Then Echo("selected=""selected""")%>>组合模式</option>
              <option value="sql" <%If mode="sql" Then Echo("selected=""selected""")%>>SQL模式</option>
            </select>
            </td>
       </tr>
    	<tr><td>标签名（name）：</td>
            <td><input name="ListName" type="text" value="MyList" /><font color="red">（ * 必填，可任意英文） </font>
            </td>
       </tr>
       
       
       
       <%If mode = "default" Then%>
       <!-- 默认模式 -->
    	<tr><td>类型（src）：</td>
            <td>
            <select name="ListSrc" onchange="doSubmit();">
              <option value="article">文章</option>
              <option value="imgart">文章[图]</option>
              <option value="picture">图片</option>
            </select><span class="gray">文章[图]：表示带焦点图片的文章</span>
            </td>
       </tr>
    	<tr><td>栏目（column）:</td>
            <td><input name="ListColumn" type="text" value="" /> <span class="gray">选项值：栏目id | auto| 缺省。多id用逗号分隔。auto则自动选择栏目，省略则全部栏目。</span>
            </td>
       </tr>
    	<tr><td>排序(Order)：</td>
            <td><select name="ListOrder" onchange="doSubmit();">
    	  <option value="asc">ID升序</option>
    	  <option value="desc">ID逆序</option>
          <option value="hot">热门</option>
          <option value="last">最新</option>
    	  <option value="asc">时间升序</option>
    	  <option value="desc">时间倒序</option>
        </select>
            </td>
       </tr>
       <%End If%>
       
       
       <%If mode = "table" Then%>
       <!-- 组合模式 -->
    	<tr><td>数据库表（table）：</td>
            <td>
            <select name="ListTable" onchange="doSubmit();">
              <option value="Article">文章[Article]</option>
              <option value="ArtColumn">文章栏目[ArtColumn]</option>
              <option value="Picture">图片[Picture]</option>
              <option value="PicColumn">图片栏目[PicColumn]</option>
              <option value="GuestBook">留言表[GuestBook]</option>
              <option value="MyTags">自定义标签表[MyTags]</option>
              <option value="DiyPage">DIY页面表[DiyPage]</option>
            </select>
            </td>
       </tr>
    	<tr><td>表字段（field）：</td>
            <td>
            <input name="ListField" type="text" value="" /> <span class="gray">选取字段，多个用逗号分隔（*表示全部）</span>
            </td>
       </tr>
    	<tr><td>条件（where）：</td>
            <td>
            <input name="ListWhere" type="text" value="" /> <span class="gray">选取条件。例如[文章表]：<font color="blue">State = 1 And IsTop = 1</font>，则表示选取置顶且审核文章</span>
            </td>
       </tr>
    	<tr><td>排序(Order)：</td>
            <td><select name="ListOrder" onchange="doSubmit();">
            <option value="">默认</option>
    	  <option value="ID ASC">ID升序</option>
    	  <option value="ID Desc">ID逆序</option>
        </select>
            </td>
       </tr>
       <%End If%>
       
       
       <%If mode = "sql" Then%>
       <!-- SQL模式 -->
    	<tr><td>SQL(sql)：</td>
            <td><input name="ListSql" type="text" value="SELECT * FROM [表名] WHERE 条件 ORDER BY 排序方式"  style="width:500px;" /><span class="gray">（<span class="red">* 必填</span>，不分大小写。如果你不熟悉SQL，建议不使用）</span>
            </td>
       </tr>
       <%End If%>
       
       
    	<tr><td>行数（row）：</td>
            <td><input name="ListRow" type="text" value="10" /> 
            </td>
       </tr>
    	<tr><td>列数(col)：</td>
            <td><input name="ListCol" type="text" value="1" /><span class="gray"> (当大于1，以表格形式输出)</span> 
            </td>
       </tr>
    	<tr><td>宽度(width)：</td>
            <td><input name="ListWidth" type="text" value="100%" /><span class="gray">（当col大于1时有效）</span> 
            </td>
       </tr>
    	<tr><td>CSS样式类(class)：</td>
            <td><input name="ListClass" type="text" value="" /><span class="gray">（当col大于1时有效）</span> 
            </td>
       </tr>
    	<tr><td>是否分页(isPage)：</td>
            <td>是：<input name="ListIspage"type="radio" value="true" /> 否：<input name="ListIspage" type="radio" value="false" checked="checked" /><span class="gray">（一个页面中只能用一次） </span>
            </td>
       </tr>
    	<tr><td colspan="2"><input type="button" onclick="doSubmit();" class="btn" value="生成列表" /></td></tr>
      </table>
     </form>
    </div>
    </fieldset>
</div>
<div id="view">
    <fieldset>
    <legend>标签参考代码</legend>
	<iframe src="tags.asp" name="showTags" width="100%" marginwidth="0" marginheight="0" scrolling="Auto" frameborder="0" id="showTags"></iframe>
    </fieldset>
    <div style="padding:5px; border:dashed 1px #CCC; margin-top:10px; color:gray;">
        <span style="font-size:13px; font-weight:bold;">内层标签属性：</span><br />
        1、<span class="red">len=""</span> 截取长度（值为数字），<span class="red">lenext=""</span>截取长度后扩展后缀（值为字符串）<br />
        &nbsp;&nbsp; &nbsp;&nbsp;例如：<br />
         &nbsp;&nbsp; &nbsp;&nbsp;<span class="green">[MyList:title len="10" lenext="..."]</span>	（此为读取文章标题，并截取前10个字符，添加省略号"..."为后缀。）<br />
        2、<span class="red">Format="yyyy-mm-dd"</span> 格式字时间，只对于<span class="blue">时间格式的字段</span>有效，如 Format="yyyy-mm-dd hh:nn:ss"，yy表示二位年份，yyyy表示四位年份，mm dd hh nn ss 都以二位表示。<br />
         &nbsp;&nbsp; &nbsp;&nbsp;例如：<br />
         &nbsp;&nbsp; &nbsp;&nbsp;<span class="green">[list:createtime format="yyyy-mmdd"]</span>	（此为把日期格式化成：2009-09-29这样的形式。）<br />
        3、<span class="red">clearhtml="true|false"</span> 是否去除HTML代码，当true时去除HTML代码。<br />
        &nbsp;&nbsp; &nbsp;&nbsp;例如：<br />
         &nbsp;&nbsp; &nbsp;&nbsp;<span class="green">[list:content clearhtml="true"]</span>	（此为把内容里面的HTML全部格式化成文本）<br />

	</div>
</div>
<script type="text/javascript">
<!--
var oInputs = document.getElementsByTagName("input");
for(var i = 0; i < oInputs.length; i++){
	 if(oInputs.item(i).name != "" && oInputs.item(i).type == "text"){
			oInputs.item(i).onmouseover = function(){
				this.style.background='#FF0';
				this.style.border = '#09F 2px solid';
			};  
			oInputs.item(i).onmouseout = function(){
				this.style.background='#FFF';
				this.style.border = '#C4E1FF 1px solid';
			};
			oInputs.item(i).onkeyup = function(){doSubmit();};
	 }
	 if(oInputs.item(i).name != "" && oInputs.item(i).type == "radio"){
		 oInputs.item(i).onclick = function(){doSubmit();};
	 }
}

function changeMode(objSel){
	this.top.location.href = '?action=list&mode=' + objSel.options[objSel.selectedIndex].value;
}

function doSubmit(){
	var chkFlag = true;
	var frm = document.forms["formList"];
	if(frm.elements["ListName"].value == ""){
		chkFlag = false;
		alert("列表名不能为空");
	}
	if( !(/^\d+$/g.test(frm.elements["ListRow"].value))){
		chkFlag = false;
		alert("行数(Row)必须为数字");
	}
	if( !(/^\d+$/g.test(frm.elements["ListCol"].value))){
		chkFlag = false;
		alert("列数(Col)必须为数字");
	}
	if(chkFlag) frm.submit();
}
//-->
</script>
<%End Sub%>




<%
 '内容标签帮助
 Sub Hcontent()
%>

<div id="view">
    <fieldset>
    <legend>内容标签语法</legend>
    <h4>1、基本语法：</h4>
<span class="red">{field:<u class="blue">字段名</u> /}  </span><span class="gray">(此标签只能在文章模板页[article.html]、图片模板页[picture.html]使用。)</span><br />
<b class="green">或者</b><br />
<span class="red">{art:<u class="blue">字段名</u>  /}</span>、 <span class="red">{article:<u class="blue">字段名</u> /}</span> <span class="gray">（此标签在文章模板页[article.html]中使用）</span><br />
<span class="red">{pic:<u class="blue">字段名</u>  /}</span> <span class="red">{picture:<u class="blue">字段名</u> /}{img:<u class="blue">字段名</u>  /}</span> <span class="red">{image:<u class="blue">字段名</u> /}</span>
<span class="gray">（此标签在图片模板页[picture.html]中使用）</span>
<br />
 <span class="green">
备注：字段名为某篇文章或者图片的所有字段名称。相应数据库字段名称，请查看数据库或者相关手册。</span><br />
<h4>2、上（下）篇特殊标签语法：</h4>
<b class="green">上一篇（文章| 图片）:</b><br />
<div class="red">
    {tag:pre type="<u class="blue">link|title|url</u>" /}<br />
    {tag:previous type="<u class="blue">link|title|url</u>" /}<br />
</div> 
<b class="green">下一篇（文章| 图片）:</b><br />
<div class="red">
	{tag:next type="<u class="blue">link|title|url</u>" /}<br />
</div> 
<b class="green">属性(可选):</b><br />
<div class="red">
	type="link" <span class="gray">文章（图片）链接。默认省略，即是{tag:pre /}等同于{tag:pre type="link"/}</span><br />
	type="title" <span class="gray">文章（图片）标题</span><br />
	type="url" <span class="gray">文章（图片）URL地址</span><br />
	type="id"	<span class="gray">文章（图片）的id</span><br />
</div> 
 <span class="green">
备注：<br />
1.文章模板页[<span class="blue">article.html</span>]中<span class="red">{tag:pre /}</span>可以写成<span class="red">{article:pre /}</span>、<span class="red">{art:pre /}</span>。<br />
2.图片模板页[<span class="blue">picture.html</span>]中<span class="red">{tag:pre /}</span>可以写成<span class="red">{picture:pre /}</span>、<span class="red">{pic:pre /}</span>、<span class="red">{image:pre /}</span>、<span class="red">{img:pre /}</span>。</span><br />

    </fieldset>
    <br />
    <fieldset>
    <legend>标签参考代码</legend>
    	选择内容标签：<select name="TagType" onchange="doSel(this);">
        	 <option value="artcommon">文章[公共](article.html)</option>
              <option value="article">文章(article.html)</option>
              <option value="piccommon">图片[公共](picture.html)</option>
              <option value="picture">图片(picture.html)</option>
            </select>
    <div id="tagCode">
    
    
    
		<div id="artcommon">
        
<h4 class="black">1、基本标签语法：</h4>
			&nbsp; {field:id /} &lt;!-- ID标识符（自动排序） --&gt;<br />&nbsp; {field:title /} &lt;!-- 文章标题 --&gt;<br />&nbsp; {field:content /} &lt;!-- 文章内容 --&gt;<br />&nbsp; {field:colid /} &lt;!-- 所属栏目ID --&gt;<br />&nbsp; {field:author /} &lt;!-- 作者 --&gt;<br />&nbsp; {field:source /} &lt;!-- 来源 --&gt;<br />&nbsp; {field:hits /} &lt;!-- 点击率 --&gt;<br />&nbsp; {field:focuspic /} &lt;!-- 焦点图片 --&gt;<br />&nbsp; {field:keywords /} &lt;!-- 关键词 --&gt;<br />&nbsp; {field:createtime /} &lt;!-- 创建时间 --&gt;<br />&nbsp; {field:modifytime /} &lt;!-- 修改时间 --&gt;<br />&nbsp; {field:istop /} &lt;!-- 是否置顶：1 - 置顶， 0 - 不置顶 --&gt;<br />&nbsp; {field:isfocuspic /} &lt;!-- 是否焦点图片：1 - 是， 0 - 否 --&gt;<br />&nbsp; {field:state /} &lt;!-- 文章状态：1 - 已经审核， 0 - 未审核， -1 - 已删除 --&gt;<br />
            <div class="green">备注：亦可把<span class="red">{field:id /}</span>简写成为<span class="red">{article:id /}、{art:id /}</span>， 其输出是等价的。</div>
            

        </div>
        
        
        
        <div id="article" style='display:none;'>
        
<h4 class="black">1、基本标签语法：</h4>
           &nbsp; {article:id /} &lt;!-- ID标识符（自动排序） --&gt;<br />&nbsp; {article:title /} &lt;!-- 文章标题 --&gt;<br />&nbsp; {article:content /} &lt;!-- 文章内容 --&gt;<br />&nbsp; {article:colid /} &lt;!-- 所属栏目ID --&gt;<br />&nbsp; {article:author /} &lt;!-- 作者 --&gt;<br />&nbsp; {article:source /} &lt;!-- 来源 --&gt;<br />&nbsp; {article:hits /} &lt;!-- 点击率 --&gt;<br />&nbsp; {article:focuspic /} &lt;!-- 焦点图片 --&gt;<br />&nbsp; {article:keywords /} &lt;!-- 关键词 --&gt;<br />&nbsp; {article:createtime /} &lt;!-- 创建时间 --&gt;<br />&nbsp; {article:modifytime /} &lt;!-- 修改时间 --&gt;<br />&nbsp; {article:istop /} &lt;!-- 是否置顶：1 - 置顶， 0 - 不置顶 --&gt;<br />&nbsp; {article:isfocuspic /} &lt;!-- 是否焦点图片：1 - 是， 0 - 否 --&gt;<br />&nbsp; {article:state /} &lt;!-- 文章状态：1 - 已经审核， 0 - 未审核， -1 - 已删除 --&gt;
            <br />
            <div class="green">备注：亦可把<span class="red">{article:id /}</span>简写成为<span class="red">{art:id /}、{field:id /}</span>， 其输出是等价的。</div>
            

        </div>
        
        
        
        <div id="piccommon" style="display:none">
<h4 class="black">1、基本标签语法：</h4>
			&nbsp; {field:id /} &lt;!-- ID标识符（自动排序） --&gt;<br />&nbsp; {field:title /} &lt;!-- 标题 --&gt;<br />&nbsp; {field:smallpicpath /} &lt;!-- 图片缩略图路径,有时为空，建议使用picpath --&gt;<br />&nbsp; {field:picpath /} &lt;!-- 图片路径 --&gt;<br />&nbsp; {field:intro /} &lt;!-- 图片介绍 --&gt;<br />&nbsp; {field:colid /} &lt;!-- 所属栏目ID --&gt;<br />&nbsp; {field:author /} &lt;!-- 作者 --&gt;<br />&nbsp; {field:source /} &lt;!-- 来源 --&gt;<br />&nbsp; {field:hits /} &lt;!-- 点击率 --&gt;<br />&nbsp; {field:createtime /} &lt;!-- 创建时间 --&gt;<br />&nbsp; {field:istop /} &lt;!-- 是否置顶：1 - 置顶， 0 - 不置顶 --&gt;<br />&nbsp; {field:state /} &lt;!-- 图片状态：1 - 已经审核， 0 - 未审核， -1 - 已删除 --&gt;<br />
			 <div class="green">备注：亦可把<span class="red">{field:id /}</span>写成为<span class="red">{picture:id /}、{pic:id /}、{image:id /}、{img:id /}</span>， 其输出是等价的。</div>
        </div>
        
        
        
        
        <div id="picture" style="display:none">
<h4 class="black">1、基本标签语法：</h4>
				&nbsp; {picture:id /} &lt;!-- ID标识符（自动排序） --&gt;<br />&nbsp; {picture:title /} &lt;!-- 标题 --&gt;<br />&nbsp; {picture:smallpicpath /} &lt;!-- 图片缩略图路径,有时为空，建议使用picpath --&gt;<br />&nbsp; {picture:picpath /} &lt;!-- 图片路径 --&gt;<br />&nbsp; {picture:intro /} &lt;!-- 图片介绍 --&gt;<br />&nbsp; {picture:colid /} &lt;!-- 所属栏目ID --&gt;<br />&nbsp; {picture:author /} &lt;!-- 作者 --&gt;<br />&nbsp; {picture:source /} &lt;!-- 来源 --&gt;<br />&nbsp; {picture:hits /} &lt;!-- 点击率 --&gt;<br />&nbsp; {picture:createtime /} &lt;!-- 创建时间 --&gt;<br />&nbsp; {picture:istop /} &lt;!-- 是否置顶：1 - 置顶， 0 - 不置顶 --&gt;<br />&nbsp; {picture:state /} &lt;!-- 图片状态：1 - 已经审核， 0 - 未审核， -1 - 已删除 --&gt;<br />
                <div class="green">备注：亦可把<span class="red">{picture:id /}</span>写成为<span class="red">{pic:id /}、{field:id /}、{image:id /}、{img:id /}</span>， 其输出是等价的。</div>
        </div>
    </div>
    </fieldset>
</div>
<script type="text/javascript" language="javascript">
<!--
var hlFlag = new Array(4);
for(var i = 0; i < 4; i++){
	hlFlag[i] = false;
}
function doSel(objSel){
	var _sel = objSel.options[objSel.selectedIndex].value;
	$("artcommon").style.display = "none";
	$("article").style.display = "none";
	$("piccommon").style.display = "none";
	$("picture").style.display = "none";
	$(_sel).style.display = "block";
	if(!hlFlag[objSel.selectedIndex]){
		HightLightTag(_sel);
		hlFlag[objSel.selectedIndex] = true;
	}
}
HightLightTag("artcommon");
hlFlag[0] = true;
-->
</script>
<%End Sub%>



<%
 '自定义标签帮助
 Sub Hmytag()
%>

<div id="view">
    <fieldset>
    <legend>My标签语法</legend>
    <h3>1、基本语法：</h3>
<span class="red">{my:<u class="blue">标签名</u> /}  </span><span class="gray">(此标签适应全部模板页。)</span><br />
 <span class="green">
 注意：自定义标签，必须再后台【标签管理】创建了之后，才能够使用。
。</span><br />

    </fieldset>
    <br />
    <fieldset>
    <legend>标签参考代码</legend>
    	
    <div id="tagCode">
		<% 
		   Dim myRs: Set myRs = DB("SELECT [Name],[Info] FROM [MyTags]", 2)
		   While Not myRs.Eof
		   		Echo("&nbsp; {my:"& myRs("Name") &" /} &lt;!-- "& myRs("Info") &" --&gt;<br />")
		   		myRs.MoveNext
		   Wend
		   myRs.Close: Set myRs = Nothing
		%>
           
      
    </div>
    </fieldset>
</div>
<script type="text/javascript" language="javascript">
<!--

HightLightTag("tagCode");
-->
</script>
<%End Sub%>


<%
 '系统变量标签帮助
 Sub Hsys()
%>

<div id="view">
    <fieldset>
    <legend>系统变量标签语法</legend>
    <h3>1、基本语法：</h3>
<span class="red">{sys:<u class="blue">变量名</u> /}  </span><span class="gray">(此标签适应全部模板页。)</span><br />
 <span class="green">
 温馨提示：系统标签可以在inc/config.asp里面查看对应的变量以及说明。
。</span><br />

    </fieldset>
    <br />
    <fieldset>
    <legend>标签参考代码</legend>
    	
    <div id="tagCode">
		&nbsp; {sys:sitename /} &lt;!-- 网站名称 --&gt;<br />
        &nbsp; {sys:siteurl  /} &lt;!-- 网站地址,相当于{sys:httpurl /} + {sys:installdir /} 的组合 --&gt;<br />
        &nbsp; {sys:skinurl /} &lt;!-- 皮肤路径 --&gt;<br />
        &nbsp; {sys:sitepath /} &lt;!-- 当前路径 --&gt;<br />
        &nbsp; {sys:installdir /} &lt;!-- 系统安装目录 --&gt;<br />
        &nbsp; {sys:httpurl /} &lt;!-- 站内链接前缀 --&gt;<br />
        &nbsp; {sys:templatedir /} &lt;!-- 当前模板目录 --&gt;<br />
        &nbsp; {sys:templatepath /} &lt;!-- 模板路径（相当于template/{sys:templatedir /}/） --&gt;<br />
        &nbsp; {sys:sitekeywords /} &lt;!-- 网站关键字 --&gt;<br />
        &nbsp; {sys:sitedesc /} &lt;!-- 网站描述 --&gt;<br />
        &nbsp; {sys:sysname /} &lt;!-- 系统名称 --&gt;<br />
        &nbsp; {sys:sysversion /} &lt;!-- 系统当前版本  --&gt;<br />
        &nbsp; {sys:sys /} &lt;!-- 系统信息（相当于{sys:sysname /} + {sys:sysversion /}）  --&gt;<br />
        
      
    </div>
    </fieldset>
</div>
<script type="text/javascript" language="javascript">
<!--

HightLightTag("tagCode");
-->
</script>
<%End Sub%>



<%
 '包含标签帮助
 Sub Hinclude()
%>

<div id="view">
    <fieldset>
    <legend>包含文件标签语法</legend>
    <h3>1、基本语法：</h3>
<span class="red">{include file="<u class="blue">文件名</u>" /}  </span><span class="gray">(此标签适应全部模板页。)</span><br />
 <span class="green">
 注意：包含的文件不能超越当前模板目录范围，即是不能包含该模板目录以外的任何文件。另外，包含不能嵌套本身，最多只能包含三层。
。</span><br />

    </fieldset>
    <br />
    <fieldset>
    <legend>标签参考代码</legend>
    	
    <div id="tagCode">
		&nbsp; {include file="header.html" /} &lt;!-- 包含header.html文件，该文件必须存在该模板目录下 --&gt;<br />
    </div>
    </fieldset>
</div>
<script type="text/javascript" language="javascript">
<!--

HightLightTag("tagCode");
-->
</script>
<%End Sub%>



<%
 'DIY页面标签标签帮助
 Sub Hdiypage()
%>

<div id="view">
    <fieldset>
    <legend>DIY页面标签语法</legend>
    <h3>1、基本语法：</h3>
<span class="red">{diypage:"<u class="blue">字段名</u>" /}  </span><span class="gray">(此标签只能在DIY模板页[diypage.html]中使用。)</span><br />
 <span class="green">
 注意：包含的文件不能超越当前模板目录范围，即是不能包含该模板目录以外的任何文件。另外，包含不能嵌套本身，最多只能包含三层。
。</span><br />

    </fieldset>
    <br />
    <fieldset>
    <legend>标签参考代码</legend>
    	
    <div id="tagCode">
		&nbsp; {diypage:id /} &lt;!-- ID标识符（自动排序） --&gt;<br />&nbsp; {diypage:title /} &lt;!-- 页面标题 --&gt;<br />&nbsp; {diypage:pagename /} &lt;!-- 该页面文件名 --&gt;<br />&nbsp; {diypage:keywords /} &lt;!-- 页面关键词 --&gt;<br />&nbsp; {diypage:template /} &lt;!-- 页面模板 --&gt;<br />&nbsp; {diypage:code /} &lt;!-- 页面代码 --&gt;<br />&nbsp; {diypage:state /} &lt;!-- 状态： 0 - 隐藏， 1 - 显示 --&gt;<br />&nbsp; {diypage:issystem /} &lt;!-- 是否是系统定义页面：0 - 否， 1 - 是 --&gt;<br />
    </div>
    </fieldset>
</div>
<script type="text/javascript" language="javascript">
<!--

HightLightTag("tagCode");
-->
</script>
<%End Sub%>



<%
 'If标签标签帮助
 Sub Hif()
%>

<div id="view">
    <fieldset>
    <legend>IF标签语法</legend>
    <h3>1、基本语法：</h3>
<div class="red">
	&nbsp;&nbsp;{if:<u class="blue">布尔表达</u>}<br />
    	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u class="blue">输出等式成立的值</u><br />
    &nbsp;&nbsp;{else}<br />
    	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u class="blue">输出表达式不成立的值</u><br />
    &nbsp;&nbsp;{/if}
    <div class="green" style="font-weight:bold;">或者：</div>
	&nbsp;&nbsp;{if:<u class="blue">布尔表达</u>}<br />
    	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u class="blue">输出表达式成立的值</u><br />
    &nbsp;&nbsp;{/if}
</div>


<div class="gray">(此标签适应全部模板页。)</div>
 <div class="green">
 注意：包含的文件不能超越当前模板目录范围，即是不能包含该模板目录以外的任何文件。另外，包含不能嵌套本身，最多只能包含三层。
。</div>

    </fieldset>
    <br />
    <fieldset>
    <legend>标签参考代码</legend>
    	
    <div id="tagCode">
		
        {list:art src="article" row="10" col="1" order="asc" ispage="false"} &lt;!--  文章art列表开始 --&gt;<br />
        	&nbsp;&nbsp;&nbsp;&nbsp;{if:[art:i] Mod 2 = 0} &lt;!-- 对[art:i]对2取余 --&gt;<br />
            	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[输出记录i]为偶数时输出文章 <br />
            &nbsp;&nbsp;{else} &lt;!-- 不符合取余为0条件 --&gt;<br />
            	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[输出记录i]为基数时输出文章<br />
            &nbsp;&nbsp;&nbsp;&nbsp;{/if} &lt;!-- If结束 --&gt;<br />
         {/list:art} &lt;!-- 列表art结束 --&gt;<br /><br />
         
         
         
          
    </div>
    <div class="gray"> 
    -----------------------<br />
         解析：由于列表执行顺序是：包含标签 → 自定义标签 → 系统标签 → 列表标签 → 分页标签 → 判断标签，所以会先执行列表标签，也只是把[art:i]变成数字1、2、3....9、10这样的数字，然后1 Mod 2（1对2取余） 自然为1，也就是偶数条件不成立，就输出{else}的内容。</div>
    </fieldset>
</div>
<script type="text/javascript" language="javascript">
<!--

HightLightTag("tagCode");
-->
</script>
<%End Sub%>




<%
 '标签执行顺序
 Sub Hdoorder()
%>

<div id="view">
    <fieldset>
    <legend>标签执行顺序</legend>
    
	&nbsp;&nbsp; 包含标签 → 自定义标签 → 系统标签 → 列表标签 → 分页标签 → 判断标签

    </fieldset>
</div>
<%End Sub%>

