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
Dim act: act = Request("action")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ǩ����ο�-�û�����</title>
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
//��ȡIDԪ��
function $(o){ return typeof(o)=="string" ? document.getElementById(o) : o;}
//��ǩ����
function HightLightTag(id){
	var strList = $(id).innerHTML;//�б�����
	
	var listName;	//�б�����
	var regExp;		//������ʽ
	//��ɫ����{tag:name }�е�name
	regExp = /\{(.+?):([^\s\]\}]+)/ig;
	strList = strList.replace(regExp, "{$1:<font color='blue'>$2</font>");
	//��ɫ����[list:name]�е�name
	//regExp = /\[(.+?):([^\s\]]+)\]/ig;
	regExp = /\[(.+?):(.+?)\]/ig;
	strList = strList.replace(regExp, " [<font color='blue'>$1</font>:<font color='green'>$2</font>\]");
	//��ɫ��������������ɫ��������ֵ
	//alert(strList);
	//regExp = /\s(\S+)=\"(\S*)\"/ig;
	regExp = /\s(\S+)=\"(.+?)\"/ig;
	strList = strList.replace(regExp, " <font color='blue'>$1</font>=\"<font color='green'>$2</font>\"");
	//alert(strList);
	//��ɫ����IF�����ײ��ǩ
	regExp = /\{if:(.+?)\}(.+?)/ig;
	strList = strList.replace(regExp, "{if:$1}<font color='green'>$2");
	regExp = /\{else\}/ig;
	strList = strList.replace(regExp, "</font>{else}<font color='green'>");
	regExp = /\{\/if\}/ig;
	strList = strList.replace(regExp, "</font>{/if}");
	//��ɫ��������˵��
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
                	<li><a href="index.asp">������ҳ</a></li>
                </ul>
                <ul class="menu">
                	<li class="mTitle">��������</li>
                    <li <%If Len(act) = 0 Or act = "index" Then Echo(" class=""on""")%>><a href="help.asp">������ҳ</a></li>
                    <li <%If act = "list" Then Echo(" class=""on""")%>><a href="?action=list">�б��ǩ</a></li>
                    <li <%If act = "content" Then Echo(" class=""on""")%>><a href="?action=content">���ݱ�ǩ</a></li>
                    <li <%If act = "mytag" Then Echo(" class=""on""")%>><a href="?action=mytag">�Զ����ǩ</a></li>
                    <li <%If act = "sys" Then Echo(" class=""on""")%>><a href="?action=sys">ϵͳ��ǩ</a></li>
                    <li <%If act = "include" Then Echo(" class=""on""")%>><a href="?action=include">�����ļ���ǩ</a></li>
                    <li <%If act = "diypage" Then Echo(" class=""on""")%>><a href="?action=diypage">DIYҳ���ǩ</a></li>
                    <li <%If act = "if" Then Echo(" class=""on""")%>><a href="?action=if">�жϱ�ǩ</a></li>
                    <li <%If act = "doorder" Then Echo(" class=""on""")%>><a href="?action=doorder">��ǩִ��˳��</a></li>
                    <li><a href="http://www.liufu.org/ling" target="_blank">�������</a></li>
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
 '��ҳ����
 Sub Hindex()
%>
 <div id="view">
    <fieldset>
    <legend>����ģ���˵��</legend>
    	<div style="padding:5px;">
<ol>
<li>������ģ�壬����Դ���һ����Ŀ¼����������������������Ϊnewtpl��������Ŀ¼�������templateĿ¼�¡�</li>
<li>ģ��Ŀ¼�У�������ڵ�ģ��ҳ�У�index.html����ҳģ��ҳ����artlist.html�������б�ģ��ҳ����article.html����������ģ��ҳ����piclist.html��ͼƬ�б�ģ��ҳ����picture.html��ͼƬ����ģ��ҳ����guestbook.html������ģ��ҳ����diypage.html���Զ���ҳ��ģ��ҳ����</li>
<li>ģ�������ͼƬ������ڵ�ǰģ�壨���磺newtpl����imagesĿ¼�£�css�ļ��������cssĿ¼�£�js�ļ��������js����scriptsĿ¼�¡�</li>
<li>��ģ����ʹ�õı�ǩ����ǩ�﷨��鿴��ǩ˵���ĵ���<br />
<div style="font-size:12px; color:gray;">

��ϵͳ��ǩ����������HTML�ӽ����﷨����Ϊ����HTML���֣���ϵͳ��ǩ���ô�����{}����HTML�е�<>��<br />
���磺<br />
{list��name mode="default" src="article"}<br />
###�ڲ�ѭ����ǩ###<br />
{/list: name }<br />
�����ǩ�ж���д�������Լ�ѡ���ʺϱ�ǩ�������õ�һ����������HTML�﷨�����п�ʼ��պϣ����д����<br />
</div></li>
<li>����ģ�����֮�󣬽��롾�����̨�� �� ��<a href="admin_config.asp">ϵͳ����</a>�� �� ѡ��<a href="admin_config.asp">ģ��</a>����ѡ��������ģ���Ŀ¼��������漴�ɡ��������վ���ܼ�ʱˢ�£������<a href="index.asp?action=clearcache">���»���</a>����Ȼ��ˢ�¼�����ɡ�</li>
<li>	�������̿��Բ���ϵͳ�Դ���ģ��ͱ���ǩ˵�����ɡ�</li>
<li>	�����κ����ʻ���bug,�뵽�ٷ�http://www.eekku.com���߷����ʼ���Foolin@126.com���з�����</li>
</ol>

<br />
ע�⣺����ģ�����߱�һ����HTML��CSS֪ʶ����ҳĬ����Gb2312���롣
<br /><br />

�ٷ���<a href="http://www.eekku.com" target="_blank">http://www.eekku.com</a><br />
��ҳ��<a href="http://www.liufu.org/ling" target="_blank">http://www.liufu.org/ling</a><br />
���䣺Foolin@126.com<br />
<br />

        </div>
    </fieldset>
    
    <fieldset>
    <legend>��ǩִ��˳��</legend>
    
	&nbsp;&nbsp; ������ǩ �� �Զ����ǩ �� ϵͳ��ǩ �� �б��ǩ �� ��ҳ��ǩ �� �жϱ�ǩ

    </fieldset>
</div>
                                    
<script type="text/javascript" language="javascript">
<!--
//HightLightTag("tagCode");
-->
</script>
<%End Sub%>




<%
 '�б����
 Sub Hlist()
 Dim mode: mode = LCase(Request("mode"))
	If Len(mode) = 0 Then mode = "default"
%>
<div id="control">
    <fieldset>
    <legend>��ǩ����ѡ��</legend>
    <div id="ctrlOpt">
	<form action="tags.asp?action=create" method="post" name="formList" target="showTags">
    <table style="color:green;">
    	<tr><td>ģʽ��mode����</td>
            <td>
            <select name="ListMode" onchange="changeMode(this);">
              <option value="default" <%If mode="default" Then Echo("selected=""selected""")%>>Ĭ��ģʽ</option>
              <option value="table" <%If mode="table" Then Echo("selected=""selected""")%>>���ģʽ</option>
              <option value="sql" <%If mode="sql" Then Echo("selected=""selected""")%>>SQLģʽ</option>
            </select>
            </td>
       </tr>
    	<tr><td>��ǩ����name����</td>
            <td><input name="ListName" type="text" value="MyList" /><font color="red">�� * ���������Ӣ�ģ� </font>
            </td>
       </tr>
       
       
       
       <%If mode = "default" Then%>
       <!-- Ĭ��ģʽ -->
    	<tr><td>���ͣ�src����</td>
            <td>
            <select name="ListSrc" onchange="doSubmit();">
              <option value="article">����</option>
              <option value="imgart">����[ͼ]</option>
              <option value="picture">ͼƬ</option>
            </select><span class="gray">����[ͼ]����ʾ������ͼƬ������</span>
            </td>
       </tr>
    	<tr><td>��Ŀ��column��:</td>
            <td><input name="ListColumn" type="text" value="" /> <span class="gray">ѡ��ֵ����Ŀid | auto| ȱʡ����id�ö��ŷָ���auto���Զ�ѡ����Ŀ��ʡ����ȫ����Ŀ��</span>
            </td>
       </tr>
    	<tr><td>����(Order)��</td>
            <td><select name="ListOrder" onchange="doSubmit();">
    	  <option value="asc">ID����</option>
    	  <option value="desc">ID����</option>
          <option value="hot">����</option>
          <option value="last">����</option>
    	  <option value="asc">ʱ������</option>
    	  <option value="desc">ʱ�䵹��</option>
        </select>
            </td>
       </tr>
       <%End If%>
       
       
       <%If mode = "table" Then%>
       <!-- ���ģʽ -->
    	<tr><td>���ݿ��table����</td>
            <td>
            <select name="ListTable" onchange="doSubmit();">
              <option value="Article">����[Article]</option>
              <option value="ArtColumn">������Ŀ[ArtColumn]</option>
              <option value="Picture">ͼƬ[Picture]</option>
              <option value="PicColumn">ͼƬ��Ŀ[PicColumn]</option>
              <option value="GuestBook">���Ա�[GuestBook]</option>
              <option value="MyTags">�Զ����ǩ��[MyTags]</option>
              <option value="DiyPage">DIYҳ���[DiyPage]</option>
            </select>
            </td>
       </tr>
    	<tr><td>���ֶΣ�field����</td>
            <td>
            <input name="ListField" type="text" value="" /> <span class="gray">ѡȡ�ֶΣ�����ö��ŷָ���*��ʾȫ����</span>
            </td>
       </tr>
    	<tr><td>������where����</td>
            <td>
            <input name="ListWhere" type="text" value="" /> <span class="gray">ѡȡ����������[���±�]��<font color="blue">State = 1 And IsTop = 1</font>�����ʾѡȡ�ö����������</span>
            </td>
       </tr>
    	<tr><td>����(Order)��</td>
            <td><select name="ListOrder" onchange="doSubmit();">
            <option value="">Ĭ��</option>
    	  <option value="ID ASC">ID����</option>
    	  <option value="ID Desc">ID����</option>
        </select>
            </td>
       </tr>
       <%End If%>
       
       
       <%If mode = "sql" Then%>
       <!-- SQLģʽ -->
    	<tr><td>SQL(sql)��</td>
            <td><input name="ListSql" type="text" value="SELECT * FROM [����] WHERE ���� ORDER BY ����ʽ"  style="width:500px;" /><span class="gray">��<span class="red">* ����</span>�����ִ�Сд������㲻��ϤSQL�����鲻ʹ�ã�</span>
            </td>
       </tr>
       <%End If%>
       
       
    	<tr><td>������row����</td>
            <td><input name="ListRow" type="text" value="10" /> 
            </td>
       </tr>
    	<tr><td>����(col)��</td>
            <td><input name="ListCol" type="text" value="1" /><span class="gray"> (������1���Ա����ʽ���)</span> 
            </td>
       </tr>
    	<tr><td>���(width)��</td>
            <td><input name="ListWidth" type="text" value="100%" /><span class="gray">����col����1ʱ��Ч��</span> 
            </td>
       </tr>
    	<tr><td>CSS��ʽ��(class)��</td>
            <td><input name="ListClass" type="text" value="" /><span class="gray">����col����1ʱ��Ч��</span> 
            </td>
       </tr>
    	<tr><td>�Ƿ��ҳ(isPage)��</td>
            <td>�ǣ�<input name="ListIspage"type="radio" value="true" /> ��<input name="ListIspage" type="radio" value="false" checked="checked" /><span class="gray">��һ��ҳ����ֻ����һ�Σ� </span>
            </td>
       </tr>
    	<tr><td colspan="2"><input type="button" onclick="doSubmit();" class="btn" value="�����б�" /></td></tr>
      </table>
     </form>
    </div>
    </fieldset>
</div>
<div id="view">
    <fieldset>
    <legend>��ǩ�ο�����</legend>
	<iframe src="tags.asp" name="showTags" width="100%" marginwidth="0" marginheight="0" scrolling="Auto" frameborder="0" id="showTags"></iframe>
    </fieldset>
    <div style="padding:5px; border:dashed 1px #CCC; margin-top:10px; color:gray;">
        <span style="font-size:13px; font-weight:bold;">�ڲ��ǩ���ԣ�</span><br />
        1��<span class="red">len=""</span> ��ȡ���ȣ�ֵΪ���֣���<span class="red">lenext=""</span>��ȡ���Ⱥ���չ��׺��ֵΪ�ַ�����<br />
        &nbsp;&nbsp; &nbsp;&nbsp;���磺<br />
         &nbsp;&nbsp; &nbsp;&nbsp;<span class="green">[MyList:title len="10" lenext="..."]</span>	����Ϊ��ȡ���±��⣬����ȡǰ10���ַ������ʡ�Ժ�"..."Ϊ��׺����<br />
        2��<span class="red">Format="yyyy-mm-dd"</span> ��ʽ��ʱ�䣬ֻ����<span class="blue">ʱ���ʽ���ֶ�</span>��Ч���� Format="yyyy-mm-dd hh:nn:ss"��yy��ʾ��λ��ݣ�yyyy��ʾ��λ��ݣ�mm dd hh nn ss ���Զ�λ��ʾ��<br />
         &nbsp;&nbsp; &nbsp;&nbsp;���磺<br />
         &nbsp;&nbsp; &nbsp;&nbsp;<span class="green">[list:createtime format="yyyy-mmdd"]</span>	����Ϊ�����ڸ�ʽ���ɣ�2009-09-29��������ʽ����<br />
        3��<span class="red">clearhtml="true|false"</span> �Ƿ�ȥ��HTML���룬��trueʱȥ��HTML���롣<br />
        &nbsp;&nbsp; &nbsp;&nbsp;���磺<br />
         &nbsp;&nbsp; &nbsp;&nbsp;<span class="green">[list:content clearhtml="true"]</span>	����Ϊ�����������HTMLȫ����ʽ�����ı���<br />

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
		alert("�б�������Ϊ��");
	}
	if( !(/^\d+$/g.test(frm.elements["ListRow"].value))){
		chkFlag = false;
		alert("����(Row)����Ϊ����");
	}
	if( !(/^\d+$/g.test(frm.elements["ListCol"].value))){
		chkFlag = false;
		alert("����(Col)����Ϊ����");
	}
	if(chkFlag) frm.submit();
}
//-->
</script>
<%End Sub%>




<%
 '���ݱ�ǩ����
 Sub Hcontent()
%>

<div id="view">
    <fieldset>
    <legend>���ݱ�ǩ�﷨</legend>
    <h4>1�������﷨��</h4>
<span class="red">{field:<u class="blue">�ֶ���</u> /}  </span><span class="gray">(�˱�ǩֻ��������ģ��ҳ[article.html]��ͼƬģ��ҳ[picture.html]ʹ�á�)</span><br />
<b class="green">����</b><br />
<span class="red">{art:<u class="blue">�ֶ���</u>  /}</span>�� <span class="red">{article:<u class="blue">�ֶ���</u> /}</span> <span class="gray">���˱�ǩ������ģ��ҳ[article.html]��ʹ�ã�</span><br />
<span class="red">{pic:<u class="blue">�ֶ���</u>  /}</span> <span class="red">{picture:<u class="blue">�ֶ���</u> /}{img:<u class="blue">�ֶ���</u>  /}</span> <span class="red">{image:<u class="blue">�ֶ���</u> /}</span>
<span class="gray">���˱�ǩ��ͼƬģ��ҳ[picture.html]��ʹ�ã�</span>
<br />
 <span class="green">
��ע���ֶ���Ϊĳƪ���»���ͼƬ�������ֶ����ơ���Ӧ���ݿ��ֶ����ƣ���鿴���ݿ��������ֲᡣ</span><br />
<h4>2���ϣ��£�ƪ�����ǩ�﷨��</h4>
<b class="green">��һƪ������| ͼƬ��:</b><br />
<div class="red">
    {tag:pre type="<u class="blue">link|title|url</u>" /}<br />
    {tag:previous type="<u class="blue">link|title|url</u>" /}<br />
</div> 
<b class="green">��һƪ������| ͼƬ��:</b><br />
<div class="red">
	{tag:next type="<u class="blue">link|title|url</u>" /}<br />
</div> 
<b class="green">����(��ѡ):</b><br />
<div class="red">
	type="link" <span class="gray">���£�ͼƬ�����ӡ�Ĭ��ʡ�ԣ�����{tag:pre /}��ͬ��{tag:pre type="link"/}</span><br />
	type="title" <span class="gray">���£�ͼƬ������</span><br />
	type="url" <span class="gray">���£�ͼƬ��URL��ַ</span><br />
	type="id"	<span class="gray">���£�ͼƬ����id</span><br />
</div> 
 <span class="green">
��ע��<br />
1.����ģ��ҳ[<span class="blue">article.html</span>]��<span class="red">{tag:pre /}</span>����д��<span class="red">{article:pre /}</span>��<span class="red">{art:pre /}</span>��<br />
2.ͼƬģ��ҳ[<span class="blue">picture.html</span>]��<span class="red">{tag:pre /}</span>����д��<span class="red">{picture:pre /}</span>��<span class="red">{pic:pre /}</span>��<span class="red">{image:pre /}</span>��<span class="red">{img:pre /}</span>��</span><br />

    </fieldset>
    <br />
    <fieldset>
    <legend>��ǩ�ο�����</legend>
    	ѡ�����ݱ�ǩ��<select name="TagType" onchange="doSel(this);">
        	 <option value="artcommon">����[����](article.html)</option>
              <option value="article">����(article.html)</option>
              <option value="piccommon">ͼƬ[����](picture.html)</option>
              <option value="picture">ͼƬ(picture.html)</option>
            </select>
    <div id="tagCode">
    
    
    
		<div id="artcommon">
        
<h4 class="black">1��������ǩ�﷨��</h4>
			&nbsp; {field:id /} &lt;!-- ID��ʶ�����Զ����� --&gt;<br />&nbsp; {field:title /} &lt;!-- ���±��� --&gt;<br />&nbsp; {field:content /} &lt;!-- �������� --&gt;<br />&nbsp; {field:colid /} &lt;!-- ������ĿID --&gt;<br />&nbsp; {field:author /} &lt;!-- ���� --&gt;<br />&nbsp; {field:source /} &lt;!-- ��Դ --&gt;<br />&nbsp; {field:hits /} &lt;!-- ����� --&gt;<br />&nbsp; {field:focuspic /} &lt;!-- ����ͼƬ --&gt;<br />&nbsp; {field:keywords /} &lt;!-- �ؼ��� --&gt;<br />&nbsp; {field:createtime /} &lt;!-- ����ʱ�� --&gt;<br />&nbsp; {field:modifytime /} &lt;!-- �޸�ʱ�� --&gt;<br />&nbsp; {field:istop /} &lt;!-- �Ƿ��ö���1 - �ö��� 0 - ���ö� --&gt;<br />&nbsp; {field:isfocuspic /} &lt;!-- �Ƿ񽹵�ͼƬ��1 - �ǣ� 0 - �� --&gt;<br />&nbsp; {field:state /} &lt;!-- ����״̬��1 - �Ѿ���ˣ� 0 - δ��ˣ� -1 - ��ɾ�� --&gt;<br />
            <div class="green">��ע����ɰ�<span class="red">{field:id /}</span>��д��Ϊ<span class="red">{article:id /}��{art:id /}</span>�� ������ǵȼ۵ġ�</div>
            

        </div>
        
        
        
        <div id="article" style='display:none;'>
        
<h4 class="black">1��������ǩ�﷨��</h4>
           &nbsp; {article:id /} &lt;!-- ID��ʶ�����Զ����� --&gt;<br />&nbsp; {article:title /} &lt;!-- ���±��� --&gt;<br />&nbsp; {article:content /} &lt;!-- �������� --&gt;<br />&nbsp; {article:colid /} &lt;!-- ������ĿID --&gt;<br />&nbsp; {article:author /} &lt;!-- ���� --&gt;<br />&nbsp; {article:source /} &lt;!-- ��Դ --&gt;<br />&nbsp; {article:hits /} &lt;!-- ����� --&gt;<br />&nbsp; {article:focuspic /} &lt;!-- ����ͼƬ --&gt;<br />&nbsp; {article:keywords /} &lt;!-- �ؼ��� --&gt;<br />&nbsp; {article:createtime /} &lt;!-- ����ʱ�� --&gt;<br />&nbsp; {article:modifytime /} &lt;!-- �޸�ʱ�� --&gt;<br />&nbsp; {article:istop /} &lt;!-- �Ƿ��ö���1 - �ö��� 0 - ���ö� --&gt;<br />&nbsp; {article:isfocuspic /} &lt;!-- �Ƿ񽹵�ͼƬ��1 - �ǣ� 0 - �� --&gt;<br />&nbsp; {article:state /} &lt;!-- ����״̬��1 - �Ѿ���ˣ� 0 - δ��ˣ� -1 - ��ɾ�� --&gt;
            <br />
            <div class="green">��ע����ɰ�<span class="red">{article:id /}</span>��д��Ϊ<span class="red">{art:id /}��{field:id /}</span>�� ������ǵȼ۵ġ�</div>
            

        </div>
        
        
        
        <div id="piccommon" style="display:none">
<h4 class="black">1��������ǩ�﷨��</h4>
			&nbsp; {field:id /} &lt;!-- ID��ʶ�����Զ����� --&gt;<br />&nbsp; {field:title /} &lt;!-- ���� --&gt;<br />&nbsp; {field:smallpicpath /} &lt;!-- ͼƬ����ͼ·��,��ʱΪ�գ�����ʹ��picpath --&gt;<br />&nbsp; {field:picpath /} &lt;!-- ͼƬ·�� --&gt;<br />&nbsp; {field:intro /} &lt;!-- ͼƬ���� --&gt;<br />&nbsp; {field:colid /} &lt;!-- ������ĿID --&gt;<br />&nbsp; {field:author /} &lt;!-- ���� --&gt;<br />&nbsp; {field:source /} &lt;!-- ��Դ --&gt;<br />&nbsp; {field:hits /} &lt;!-- ����� --&gt;<br />&nbsp; {field:createtime /} &lt;!-- ����ʱ�� --&gt;<br />&nbsp; {field:istop /} &lt;!-- �Ƿ��ö���1 - �ö��� 0 - ���ö� --&gt;<br />&nbsp; {field:state /} &lt;!-- ͼƬ״̬��1 - �Ѿ���ˣ� 0 - δ��ˣ� -1 - ��ɾ�� --&gt;<br />
			 <div class="green">��ע����ɰ�<span class="red">{field:id /}</span>д��Ϊ<span class="red">{picture:id /}��{pic:id /}��{image:id /}��{img:id /}</span>�� ������ǵȼ۵ġ�</div>
        </div>
        
        
        
        
        <div id="picture" style="display:none">
<h4 class="black">1��������ǩ�﷨��</h4>
				&nbsp; {picture:id /} &lt;!-- ID��ʶ�����Զ����� --&gt;<br />&nbsp; {picture:title /} &lt;!-- ���� --&gt;<br />&nbsp; {picture:smallpicpath /} &lt;!-- ͼƬ����ͼ·��,��ʱΪ�գ�����ʹ��picpath --&gt;<br />&nbsp; {picture:picpath /} &lt;!-- ͼƬ·�� --&gt;<br />&nbsp; {picture:intro /} &lt;!-- ͼƬ���� --&gt;<br />&nbsp; {picture:colid /} &lt;!-- ������ĿID --&gt;<br />&nbsp; {picture:author /} &lt;!-- ���� --&gt;<br />&nbsp; {picture:source /} &lt;!-- ��Դ --&gt;<br />&nbsp; {picture:hits /} &lt;!-- ����� --&gt;<br />&nbsp; {picture:createtime /} &lt;!-- ����ʱ�� --&gt;<br />&nbsp; {picture:istop /} &lt;!-- �Ƿ��ö���1 - �ö��� 0 - ���ö� --&gt;<br />&nbsp; {picture:state /} &lt;!-- ͼƬ״̬��1 - �Ѿ���ˣ� 0 - δ��ˣ� -1 - ��ɾ�� --&gt;<br />
                <div class="green">��ע����ɰ�<span class="red">{picture:id /}</span>д��Ϊ<span class="red">{pic:id /}��{field:id /}��{image:id /}��{img:id /}</span>�� ������ǵȼ۵ġ�</div>
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
 '�Զ����ǩ����
 Sub Hmytag()
%>

<div id="view">
    <fieldset>
    <legend>My��ǩ�﷨</legend>
    <h3>1�������﷨��</h3>
<span class="red">{my:<u class="blue">��ǩ��</u> /}  </span><span class="gray">(�˱�ǩ��Ӧȫ��ģ��ҳ��)</span><br />
 <span class="green">
 ע�⣺�Զ����ǩ�������ٺ�̨����ǩ����������֮�󣬲��ܹ�ʹ�á�
��</span><br />

    </fieldset>
    <br />
    <fieldset>
    <legend>��ǩ�ο�����</legend>
    	
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
 'ϵͳ������ǩ����
 Sub Hsys()
%>

<div id="view">
    <fieldset>
    <legend>ϵͳ������ǩ�﷨</legend>
    <h3>1�������﷨��</h3>
<span class="red">{sys:<u class="blue">������</u> /}  </span><span class="gray">(�˱�ǩ��Ӧȫ��ģ��ҳ��)</span><br />
 <span class="green">
 ��ܰ��ʾ��ϵͳ��ǩ������inc/config.asp����鿴��Ӧ�ı����Լ�˵����
��</span><br />

    </fieldset>
    <br />
    <fieldset>
    <legend>��ǩ�ο�����</legend>
    	
    <div id="tagCode">
		&nbsp; {sys:sitename /} &lt;!-- ��վ���� --&gt;<br />
        &nbsp; {sys:siteurl  /} &lt;!-- ��վ��ַ,�൱��{sys:httpurl /} + {sys:installdir /} ����� --&gt;<br />
        &nbsp; {sys:skinurl /} &lt;!-- Ƥ��·�� --&gt;<br />
        &nbsp; {sys:sitepath /} &lt;!-- ��ǰ·�� --&gt;<br />
        &nbsp; {sys:installdir /} &lt;!-- ϵͳ��װĿ¼ --&gt;<br />
        &nbsp; {sys:httpurl /} &lt;!-- վ������ǰ׺ --&gt;<br />
        &nbsp; {sys:templatedir /} &lt;!-- ��ǰģ��Ŀ¼ --&gt;<br />
        &nbsp; {sys:templatepath /} &lt;!-- ģ��·�����൱��template/{sys:templatedir /}/�� --&gt;<br />
        &nbsp; {sys:sitekeywords /} &lt;!-- ��վ�ؼ��� --&gt;<br />
        &nbsp; {sys:sitedesc /} &lt;!-- ��վ���� --&gt;<br />
        &nbsp; {sys:sysname /} &lt;!-- ϵͳ���� --&gt;<br />
        &nbsp; {sys:sysversion /} &lt;!-- ϵͳ��ǰ�汾  --&gt;<br />
        &nbsp; {sys:sys /} &lt;!-- ϵͳ��Ϣ���൱��{sys:sysname /} + {sys:sysversion /}��  --&gt;<br />
        
      
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
 '������ǩ����
 Sub Hinclude()
%>

<div id="view">
    <fieldset>
    <legend>�����ļ���ǩ�﷨</legend>
    <h3>1�������﷨��</h3>
<span class="red">{include file="<u class="blue">�ļ���</u>" /}  </span><span class="gray">(�˱�ǩ��Ӧȫ��ģ��ҳ��)</span><br />
 <span class="green">
 ע�⣺�������ļ����ܳ�Խ��ǰģ��Ŀ¼��Χ�����ǲ��ܰ�����ģ��Ŀ¼������κ��ļ������⣬��������Ƕ�ױ������ֻ�ܰ������㡣
��</span><br />

    </fieldset>
    <br />
    <fieldset>
    <legend>��ǩ�ο�����</legend>
    	
    <div id="tagCode">
		&nbsp; {include file="header.html" /} &lt;!-- ����header.html�ļ������ļ�������ڸ�ģ��Ŀ¼�� --&gt;<br />
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
 'DIYҳ���ǩ��ǩ����
 Sub Hdiypage()
%>

<div id="view">
    <fieldset>
    <legend>DIYҳ���ǩ�﷨</legend>
    <h3>1�������﷨��</h3>
<span class="red">{diypage:"<u class="blue">�ֶ���</u>" /}  </span><span class="gray">(�˱�ǩֻ����DIYģ��ҳ[diypage.html]��ʹ�á�)</span><br />
 <span class="green">
 ע�⣺�������ļ����ܳ�Խ��ǰģ��Ŀ¼��Χ�����ǲ��ܰ�����ģ��Ŀ¼������κ��ļ������⣬��������Ƕ�ױ������ֻ�ܰ������㡣
��</span><br />

    </fieldset>
    <br />
    <fieldset>
    <legend>��ǩ�ο�����</legend>
    	
    <div id="tagCode">
		&nbsp; {diypage:id /} &lt;!-- ID��ʶ�����Զ����� --&gt;<br />&nbsp; {diypage:title /} &lt;!-- ҳ����� --&gt;<br />&nbsp; {diypage:pagename /} &lt;!-- ��ҳ���ļ��� --&gt;<br />&nbsp; {diypage:keywords /} &lt;!-- ҳ��ؼ��� --&gt;<br />&nbsp; {diypage:template /} &lt;!-- ҳ��ģ�� --&gt;<br />&nbsp; {diypage:code /} &lt;!-- ҳ����� --&gt;<br />&nbsp; {diypage:state /} &lt;!-- ״̬�� 0 - ���أ� 1 - ��ʾ --&gt;<br />&nbsp; {diypage:issystem /} &lt;!-- �Ƿ���ϵͳ����ҳ�棺0 - �� 1 - �� --&gt;<br />
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
 'If��ǩ��ǩ����
 Sub Hif()
%>

<div id="view">
    <fieldset>
    <legend>IF��ǩ�﷨</legend>
    <h3>1�������﷨��</h3>
<div class="red">
	&nbsp;&nbsp;{if:<u class="blue">�������</u>}<br />
    	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u class="blue">�����ʽ������ֵ</u><br />
    &nbsp;&nbsp;{else}<br />
    	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u class="blue">������ʽ��������ֵ</u><br />
    &nbsp;&nbsp;{/if}
    <div class="green" style="font-weight:bold;">���ߣ�</div>
	&nbsp;&nbsp;{if:<u class="blue">�������</u>}<br />
    	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u class="blue">������ʽ������ֵ</u><br />
    &nbsp;&nbsp;{/if}
</div>


<div class="gray">(�˱�ǩ��Ӧȫ��ģ��ҳ��)</div>
 <div class="green">
 ע�⣺�������ļ����ܳ�Խ��ǰģ��Ŀ¼��Χ�����ǲ��ܰ�����ģ��Ŀ¼������κ��ļ������⣬��������Ƕ�ױ������ֻ�ܰ������㡣
��</div>

    </fieldset>
    <br />
    <fieldset>
    <legend>��ǩ�ο�����</legend>
    	
    <div id="tagCode">
		
        {list:art src="article" row="10" col="1" order="asc" ispage="false"} &lt;!--  ����art�б�ʼ --&gt;<br />
        	&nbsp;&nbsp;&nbsp;&nbsp;{if:[art:i] Mod 2 = 0} &lt;!-- ��[art:i]��2ȡ�� --&gt;<br />
            	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[�����¼i]Ϊż��ʱ������� <br />
            &nbsp;&nbsp;{else} &lt;!-- ������ȡ��Ϊ0���� --&gt;<br />
            	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[�����¼i]Ϊ����ʱ�������<br />
            &nbsp;&nbsp;&nbsp;&nbsp;{/if} &lt;!-- If���� --&gt;<br />
         {/list:art} &lt;!-- �б�art���� --&gt;<br /><br />
         
         
         
          
    </div>
    <div class="gray"> 
    -----------------------<br />
         �����������б�ִ��˳���ǣ�������ǩ �� �Զ����ǩ �� ϵͳ��ǩ �� �б��ǩ �� ��ҳ��ǩ �� �жϱ�ǩ�����Ի���ִ���б��ǩ��Ҳֻ�ǰ�[art:i]�������1��2��3....9��10���������֣�Ȼ��1 Mod 2��1��2ȡ�ࣩ ��ȻΪ1��Ҳ����ż�������������������{else}�����ݡ�</div>
    </fieldset>
</div>
<script type="text/javascript" language="javascript">
<!--

HightLightTag("tagCode");
-->
</script>
<%End Sub%>




<%
 '��ǩִ��˳��
 Sub Hdoorder()
%>

<div id="view">
    <fieldset>
    <legend>��ǩִ��˳��</legend>
    
	&nbsp;&nbsp; ������ǩ �� �Զ����ǩ �� ϵͳ��ǩ �� �б��ǩ �� ��ҳ��ǩ �� �жϱ�ǩ

    </fieldset>
</div>
<%End Sub%>

