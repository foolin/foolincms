<!--#include file="../inc/admin.chklogin.asp"-->
<!--#include file="../../inc/class_upload.asp"-->
<%Call ChkLogin()%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn" lang="zh-cn">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ϴ�����ͼƬ</title>
<style type="text/css">
TABLE {border:1px green solid;margin-top:5px;}
TD{border-bottom:1px #dddddd solid;height:20px;padding:3px 0 0 5px;}
.head{background-color:#eeeeee;}

</style>
</head>
<body style="font-size:12px;margin:0px;">
<%
if request.QueryString("act")="upload" then
 Dim Upload,path,tempCls,fName
'===============================================================================
 set Upload=new AnUpLoad				 				'������ʵ��
 Upload.SingleSize=1024*1024*1024            			'���õ����ļ�����ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
 Upload.MaxSize=1024*1024*1024            				'��������ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
 Upload.Exe="bmp|rar|pdf|jpg|gif"          				'���úϷ���չ��,��|�ָ�,���Դ�Сд
 Upload.Charset="gb2312"								'�����ı����룬Ĭ��Ϊgb2312
 Upload.openProcesser=false								'��ֹ���������ܣ�������ã�����Ͽͻ��˳���
 Upload.GetData()										'��ȡ����������,������ñ�����
'===============================================================================
 if Upload.ErrorID>0 then								'�жϴ����,���myupload.Err<=0��ʾ����
 	response.write Upload.Description 					'������ִ���,��ȡ��������
 else
 	if Upload.files(-1).count>0 then 					'�����ж����Ƿ�ѡ�����ļ�
    		path=server.mappath("../../upload/images/focuspic/") 				'�ļ�����·��(������files�ļ���)
    		'�����ļ�(�����ļ�������)
    		set tempCls=Upload.files("file1") 
    		tempCls.SaveToFile path,0
    	    fName=tempCls.FileName
    		set tempCls=nothing
    else
		response.Write "��û���ϴ��κ��ļ���"
 	end if
 end if
 set Upload=nothing                   '������ʵ��
 %>
 <script type ="text/javascript" language="javascript">
 <!--//
	window.parent.document.forms["form1"].elements["FocusPic"].value='upload/images/focuspic/<%=fName%>';
 //-->
 </script>
 �ļ��ϴ��ɹ�.[<a href='upload_focuspic.asp'>�����ϴ�</a>]
 <%
 else
 %>
 <form name="upload" method="post" action="?act=upload" enctype="multipart/form-data" style="margin:0px;padding:0px;">
<input type ="file" name ="file1" /> <input type ="submit" value="�ϴ�" /> 
</form>
 <%
 end if
%>
</body>
</html>
