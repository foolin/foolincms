<!--#include file="inc/admin.include.asp"-->
<%
	'��¼�˳�
	Call WebLog("�û�[User:"& GetCookies("AdminName") &"]�˳��ɹ�", "SESSION")
	'���Cookies
	Call SetCookies("AdminName","")
	Call SetCookies("AdminNickname","")
	Call SetCookies("AdminPassword","")
	Call SetCookies("AdminLevel","")
	Call ConnClose()	'�ر�����
%>
<script type='text/javascript'>alert('�ɹ��˳�');this.top.location.href='login.asp';</script>
