<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit		'ǿ������
'On Error Resume Next	'�ݴ���

Dim DBPATH			'Access���ݿ�·��
	DBPATH = "database/data.mdb"
	
Dim SITENAME		'��վ����
	SITENAME = "E����"
	
	
Dim SITEKEYWORDS	'��վ�ؼ��ʣ����ö��ŷָ���
	SITEKEYWORDS = "E����,E�Ṥ����,CMS,eekku.com"
	
Dim HTTPURL			'��վ��ַǰ׺��ǰ��Ҫ��http
	HTTPURL	 = "http://localhost/"
	
Dim INSTALLDIR		'��װĿ¼�����治�ü�/
	INSTALLDIR = "/eekku"
		
Dim TEMPLATEPATH	'ģ��·�������治�ü�/
	TEMPLATEPATH = "/template/default"

	
Dim ISHIDETEMPPATH	'�Ƿ�����ģ��·����������·���������ҳ�������ٶȻ���һЩ
	ISHIDETEMPPATH = 0

Dim ISCACHE	'�Ƿ񻺴�ģ��, 1��ʾ���棬0��ʾ������
	ISCACHE = 0

Dim CACHEFLAG		'�����־
	CACHEFLAG = "EEKKU"

Dim CACHETIME		'����ʱ��,��λΪ��
	CACHETIME = 0
	
Dim ISWEBLOG		'�Ƿ��¼��̨����
	ISWEBLOG = 0
	
Dim LIMITIP			'���Ʒ���IP����IP��|�ָ�
	LIMITIP = "172.168.168.20|"

Dim DIRTYWORDS		'�Ƿ����໰�����ˣ������|�ָ�
	DIRTYWORDS = "������|������|�¼ұ�|�����|������|������|���|���ֹ�|���־|�Ҳ�|�Ҳ�"
%>