<!--#include file="inc/admin.include.asp"-->
<%
	'��¼�˳�
	Call WebLog("�û�[User:"& Session("AdminName") &"]�˳��ɹ�", "SESSION")
	'���Session
	Session("AdminName")=""
	Session("AdminLevel")=""
	Call ConnClose()	'�ر�����
%>
<script type='text/javascript'>alert('�ɹ��˳�');this.top.location.href='login.asp';</script>
