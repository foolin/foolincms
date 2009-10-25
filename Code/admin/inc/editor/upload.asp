<!--#include file="../../../inc/include.asp"-->
<!--#include file="../class_upload.asp"-->
<!--#include file="../admin.func_chkadmin.asp"-->
<%
 Dim Upload,successful,path,tempCls
'===============================================================================
 set Upload=new AnUpLoad				 				'创建类实例
 Upload.SingleSize=1024*1024*1024            			'设置单个文件最大上传限制,按字节计；默认为不限制
 Upload.MaxSize=1024*1024*1024            				'设置最大上传限制,按字节计；默认为不限制
 Upload.Exe="bmp|png|jpg|gif"          					'设置合法扩展名,以|分割,忽略大小写
 'Upload.Charset= "utf-8"								'设置文本编码，默认为gb2312
 Upload.openProcesser=false								'禁止进度条功能，如果启用，需配合客户端程序
 Upload.GetData()										'获取并保存数据,必须调用本方法
'===============================================================================
 if Upload.ErrorID>0 then						'判断错误号,如果myupload.Err<=0表示正常
 	response.write Upload.Description 			'如果出现错误,获取错误描述
 else
 	if Upload.forms("imgFile")<>"" then 			'这里判断你file1是否选择了文件
    		path=server.mappath("../../../upload/images/") 			'文件保存路径(这里是files文件夹)
    		set tempCls=Upload.files("imgFile") 
    		successful=tempCls.SaveToFile(path,0)		'以时间+随机数字为文件名保存
    		'successful=tempCls.SaveToFile(path,1)		'如果想以原文件名保存,请使用本句
		if successful then
			'//插入图片，关闭层
			Response.Write "<html>"
			Response.Write "<head>"
			Response.Write "<title>Insert Image</title>"
			Response.Write "<meta http-equiv=""content-type"" content=""text/html; charset=gb2312"">"
			Response.Write "</head>"
			Response.Write "<body>"
			Response.Write "<script type=""text/javascript"">parent.KE.plugin[""image""].insert('"&Upload.forms("id")&"', '"&"../upload/images/"&tempCls.FileName&"','"&Upload.forms("imgTitle")&"','"&Upload.forms("imgWidth")&"','"&Upload.forms("imgHeight")&"','"&Upload.forms("imgBorder")&"');</script>"
			Response.Write "</body>"
			Response.Write "</html>"
		else
			response.write "上传失败"
		end if
    		set tempCls=nothing
 	end if
 end if



'//提示，关闭层
Sub alert(msg)
    Response.Write "<html>"
    Response.Write "<head>"
    Response.Write "<title>error</title>"
    Response.Write "<meta http-equiv=""content-type"" content=""text/html; charset=gb2312"">"
    Response.Write "</head>"
    Response.Write "<body>"
    Response.Write "<script type=""text/javascript"">alert('"&msg&"');history.back();</script>"
    Response.Write "</body>"
    Response.Write "</html>"
End Sub
%>