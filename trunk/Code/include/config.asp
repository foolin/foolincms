<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit		'ǿ������
'On Error Resume Next	'�ݴ���

Dim DBPath			'Access���ݿ�·��
	DBpath = "database/data.mdb"
	
Dim SiteName		'��վ����
	SiteName = "E����"
	
	
Dim SiteKeywords	'��վ�ؼ��ʣ����ö��ŷָ���
	SiteKeywords = "E����,E�Ṥ����,CMS,eekku.com"
	
Dim HttpUrl			'��վ��ַǰ׺��ǰ��Ҫ��http
	HttpUrl = "http://localhost/"
	
Dim InstallDir		'��װĿ¼�����治�ü�/
	InstallDir = "/eekku"
		
Dim TemplatePath	'ģ��·�������治�ü�/
	TemplatePath = "/template/default"

	
Dim IsHideTempPath	'�Ƿ�����ģ��·����������·���������ҳ�������ٶȻ���һЩ
	IsHideTempPath = 0

Dim IsCache	'�Ƿ񻺴�ģ��, 1��ʾ���棬0��ʾ������
	IsCache = 0

Dim CacheFlag		'�����־
	CacheFlag = "EEKKU"

Dim CacheTime		'����ʱ��,��λΪ��
	CacheTime = 0
	
Dim LimitIP			'���Ʒ���IP����IP��|�ָ�
	LimitIP = "172.168.168.20|"

Dim DirtyWords		'�Ƿ����໰�����ˣ������|�ָ�
	DirtyWords = "������|������|�¼ұ�|�����|������|������|���|���ֹ�|���־|�Ҳ�|�Ҳ�"
%>