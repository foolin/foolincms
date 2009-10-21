<!--#include file="../../inc/config.asp"-->
<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/md5.asp"-->
<!--#include file="../../inc/func_file.asp"-->
<!--#include file="../../inc/func_common.asp"-->
<!--#include file="admin.func_chkadmin.asp"-->
<!--#include file="class_upload.asp"-->
<%
 Call ChkLogin()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>上传焦点图片</title>
<style type="text/css">
TABLE {border:1px green solid;margin-top:5px;}
TD{border-bottom:1px #dddddd solid;height:20px;padding:3px 0 0 5px;}
.head{background-color:#eeeeee;}

</style>
</head>
<body style="font-size:12px;margin:0px;">
<%
'请求上传时
if request.QueryString("act")="upload" then
 Dim Upload, strPath, tempCls, fName, strPic, strSmallPic, Jpeg
 Dim strFolder:  strFolder = "upload/images/" & GetFolderName
'===============================================================================
 set Upload=new AnUpLoad				 				'创建类实例
 Upload.SingleSize=300*1024            					'设置单个文件最大上传限制,按字节计；默认为不限制
 Upload.MaxSize=1024*1024*1024            				'设置最大上传限制,按字节计；默认为不限制
 Upload.Exe="bmp|png|jpg|gif"          					'设置合法扩展名,以|分割,忽略大小写
 Upload.Charset = CHARSET								'设置文本编码，默认为gb2312
 Upload.openProcesser=false								'禁止进度条功能，如果启用，需配合客户端程序
 Upload.GetData()										'获取并保存数据,必须调用本方法
'===============================================================================

 if Upload.ErrorID>0 then								'判断错误号,如果myupload.Err<=0表示正常
 	Response.Write Upload.Description 					'如果出现错误,获取错误描述
	Response.Write "[<a href='upload_picture.asp'>重新上传</a>]"
	Response.End()
 else
 	if Upload.files(-1).count>0 then 					'这里判断你是否选择了文件
			If ExistFolder("../../" & strFolder) = False Then
				CreateFolder("../../" & strFolder)
			End If
    		strPath = server.mappath("../../" & strFolder) 				'文件保存路径(这里是files文件夹)
    		'保存文件(以新文件名保存)
    		set tempCls = Upload.files("file1") 
    		tempCls.SaveToFile strPath,0
    	    fName=tempCls.FileName
    		set tempCls=nothing
			strPic = strFolder & fName	'图片路径
			strSmallPic = strPic		'缩略图路径（初始化，不支持AspJpeg组件时用）
			'判断是否支持Persit.jpeg组件
			If IsObjInstalled("Persits.Jpeg") Then
				strPath = Server.MapPath("../../" & strPic ) '处理图片路径
				strSmallPic = strFolder & "small_" & fName
				Set Jpeg = Server.CreateObject("Persits.Jpeg") '创建实例
				Jpeg.Open strPath '打开图片
				'调整宽度和高度
				Jpeg.Width = GetWidth(Jpeg.OriginalWidth)
				Jpeg.Height = GetHeight(Jpeg.OriginalWidth, Jpeg.OriginalHeight)
				Jpeg.Save Server.MapPath("../../" & strSmallPic) '保存图片到磁盘
				Jpeg.Close:Set Jpeg = Nothing
			End If
%>
 <script type ="text/javascript" language="javascript">
 <!--//
 	var _smlPic = window.parent.document.forms["form1"].elements["fSmallPicPath"];
 	var _pic = window.parent.document.forms["form1"].elements["fPicPath"];
	var _numId = window.parent.document.getElementById("PicNum");
	var _num = parseInt(_numId.innerHTML) + 1;
	_numId.innerHTML = _num;
	if (_pic.value != '')
	{
		_smlPic.value = _smlPic.value + '|<%Response.Write(strSmallPic)%>';
		_pic.value =  _pic.value + '|<%Response.Write(strPic)%>';
	}
	else
	{
		_smlPic.value = '<%Response.Write(strSmallPic)%>';
		_pic.value = '<%Response.Write(strPic)%>';
	}
 //-->
 </script>
 文件上传成功.
<%
    else
		Response.Write "您没有上传任何文件！"
 	end if
end if
set Upload=nothing                   '销毁类实例
%>
[<a href='upload_picture.asp'>继续上传</a>]
 <%
 else
 %>
 <form name="upload" method="post" action="?act=upload" enctype="multipart/form-data" style="margin:0px;padding:0px;">
<input type ="file" name ="file1" /> <input type ="submit" value="上传" /> 
</form>
 <%
 end if
%>
</body>
</html>
<%
 '自动生成Folder函数
 Function GetFolderName
	Dim sYear, sMonth
	sYear = Year(Now())
	sMonth = Month(Now())
	If Cint(sMonth) < 10 Then sMonth = "0" & sMonth
	GetFolderName = sYear & "/" & sMonth & "/"
 End Function
 
 '获取缩略图高度
 Function GetWidth(Byval oWidth)
	Dim tW: tW = oWidth
 	If oWidth > 200 Then
		tW = 200
	End If
	GetWidth = tW
 End Function
 
 '获取缩略图高度
 Function GetHeight(Byval oWidth, ByVal oHeight)
	Dim tH: tH = oHeight
 	If oWidth > 200 Then
		tH = oHeight / (oWidth/200)
	End If
	GetHeight = tH
 End Function
%>

