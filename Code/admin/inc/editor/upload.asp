<!--#include file="../../../inc/include.asp"-->
<!--#include file="../class_upload.asp"-->
<!--#include file="../admin.func_chkadmin.asp"-->
<%
 Dim Upload,successful,path,tempCls
'===============================================================================
 set Upload=new AnUpLoad				 				'������ʵ��
 Upload.SingleSize=1024*1024*1024            			'���õ����ļ�����ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
 Upload.MaxSize=1024*1024*1024            				'��������ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
 Upload.Exe="bmp|png|jpg|gif"          					'���úϷ���չ��,��|�ָ�,���Դ�Сд
 'Upload.Charset= "utf-8"								'�����ı����룬Ĭ��Ϊgb2312
 Upload.openProcesser=false								'��ֹ���������ܣ�������ã�����Ͽͻ��˳���
 Upload.GetData()										'��ȡ����������,������ñ�����
'===============================================================================
 if Upload.ErrorID>0 then						'�жϴ����,���myupload.Err<=0��ʾ����
 	response.write Upload.Description 			'������ִ���,��ȡ��������
 else
 	if Upload.forms("imgFile")<>"" then 			'�����ж���file1�Ƿ�ѡ�����ļ�
    		path=server.mappath("../../../upload/images/") 			'�ļ�����·��(������files�ļ���)
    		set tempCls=Upload.files("imgFile") 
    		successful=tempCls.SaveToFile(path,0)		'��ʱ��+�������Ϊ�ļ�������
    		'successful=tempCls.SaveToFile(path,1)		'�������ԭ�ļ�������,��ʹ�ñ���
		if successful then
			'//����ͼƬ���رղ�
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
			response.write "�ϴ�ʧ��"
		end if
    		set tempCls=nothing
 	end if
 end if



'//��ʾ���رղ�
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