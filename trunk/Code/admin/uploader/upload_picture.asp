<!--#include file="../inc/admin.chklogin.asp"-->
<!--#include file="../../inc/class_upload.asp"-->
<!--#include file="../../inc/func_file.asp"-->
<%
 Call ChkLogin()

 '�Զ�����Folder����
 Function GetFolderName
	Dim sYear, sMonth
	sYear = Year(Now())
	sMonth = Month(Now())
	If Cint(sMonth) < 10 Then sMonth = "0" & sMonth
	GetFolderName = sYear & "/" & sMonth & "/"
 End Function
%>
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
'�����ϴ�ʱ
if request.QueryString("act")="upload" then
 Dim Upload,path,tempCls,fName
 Dim strFolder:  strFolder = "upload/images/" & GetFolderName
'===============================================================================
 set Upload=new AnUpLoad				 				'������ʵ��
 Upload.SingleSize=200*1024            			'���õ����ļ�����ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
 Upload.MaxSize=1024*1024*1024            				'��������ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
 Upload.Exe="bmp|rar|pdf|jpg|gif"          				'���úϷ���չ��,��|�ָ�,���Դ�Сд
 Upload.Charset="gb2312"								'�����ı����룬Ĭ��Ϊgb2312
 Upload.openProcesser=false								'��ֹ���������ܣ�������ã�����Ͽͻ��˳���
 Upload.GetData()										'��ȡ����������,������ñ�����
'===============================================================================

 if Upload.ErrorID>0 then								'�жϴ����,���myupload.Err<=0��ʾ����
 	Response.Write Upload.Description 					'������ִ���,��ȡ��������
	Response.Write "[<a href='upload_picture.asp'>�����ϴ�</a>]"
	Response.End()
 else
 	if Upload.files(-1).count>0 then 					'�����ж����Ƿ�ѡ�����ļ�
			If ExistFolder("../../" & strFolder) = False Then
				CreateFolder("../../" & strFolder)
			End If
    		path=server.mappath("../../" & strFolder) 				'�ļ�����·��(������files�ļ���)
    		'�����ļ�(�����ļ�������)
    		set tempCls = Upload.files("file1") 
    		tempCls.SaveToFile path,0
    	    fName=tempCls.FileName
    		set tempCls=nothing
    else
		Response.Write "��û���ϴ��κ��ļ���"
 	end if
 end if
 set Upload=nothing                   '������ʵ��
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
		_smlPic.value = _smlPic.value + '|<%Response.Write(strFolder & fName)%>';
		_pic.value =  _pic.value + '|<%Response.Write(strFolder & fName)%>';
	}
	else
	{
		_smlPic.value = '<%Response.Write(strFolder & fName)%>';
		_pic.value = '<%Response.Write(strFolder & fName)%>';
	}
 //-->
 </script>
 �ļ��ϴ��ɹ�.[<a href='upload_picture.asp'>�����ϴ�</a>]
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
