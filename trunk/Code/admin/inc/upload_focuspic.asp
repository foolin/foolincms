<!--#include file="../../inc/config.asp"-->
<!--#include file="../../inc/conn.asp"-->
<!--#include file="../../inc/md5.asp"-->
<!--#include file="../../inc/func_file.asp"-->
<!--#include file="../../inc/func_common.asp"-->
<!--#include file="admin.func_chkadmin.asp"-->
<!--#include file="class_upload.asp"-->
<%Call ChkLogin()%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
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
 '�Զ�����Folder����
 Function GetFolderName
	Dim sYear, sMonth
	sYear = Year(Now())
	sMonth = Month(Now())
	If Cint(sMonth) < 10 Then sMonth = "0" & sMonth
	GetFolderName = sYear & "/" & sMonth & "/"
 End Function
 
if request.QueryString("act")="upload" then
 Dim Upload,path,tempCls,fName
 Dim strFolder:  strFolder = "upload/images/focuspic/" & GetFolderName
'===============================================================================
 set Upload=new AnUpLoad				 				'������ʵ��
 Upload.SingleSize=1024*1024*1024            			'���õ����ļ�����ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
 Upload.MaxSize=1024*1024*1024            				'��������ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
 Upload.Exe="bmp|png|jpg|gif"          					'���úϷ���չ��,��|�ָ�,���Դ�Сд
 Upload.Charset=CHARSET									'�����ı����룬Ĭ��Ϊgb2312
 Upload.openProcesser=false								'��ֹ���������ܣ�������ã�����Ͽͻ��˳���
 Upload.GetData()										'��ȡ����������,������ñ�����
'===============================================================================
 if Upload.ErrorID>0 then								'�жϴ����,���myupload.Err<=0��ʾ����
 	response.write Upload.Description 					'������ִ���,��ȡ��������
	Response.Write "[<a href='?'>�����ϴ�</a>]"
	Response.End()
 else
 	if Upload.files(-1).count>0 then 					'�����ж����Ƿ�ѡ�����ļ�
			If ExistFolder("../../" & strFolder) = False Then
				CreateFolder("../../" & strFolder)
			End If
    		path=server.mappath("../../" & strFolder) 
    		'�����ļ�(�����ļ�������)
    		set tempCls=Upload.files("file1") 
    		tempCls.SaveToFile path,0
    	    fName=tempCls.FileName
    		set tempCls=nothing
%>
 �ļ��ϴ��ɹ�.
 <script type ="text/javascript" language="javascript">
 <!--//
	window.parent.document.forms["form1"].elements["FocusPic"].value='<%=strFolder & fName%>';
	//���µ��༭��
 	parent.KE.util.focus("content1");
	parent.KE.util.selection("content1"); 
 	parent.KE.util.insertHtml("content1", "<img src=\"../<%=strFolder & fName%>\" border=\"0\" \/>");
 //-->
 </script>
<%
    else
		response.Write "��û���ϴ��κ��ļ���"
 	end if
 end if
 set Upload=nothing                   '������ʵ��
 %>
[<a href='upload_focuspic.asp'>�����ϴ�</a>]
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

