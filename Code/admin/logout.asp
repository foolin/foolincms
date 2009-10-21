<!--#include file="inc/admin.include.asp"-->
<%
	'记录退出
	Call WebLog("用户[User:"& GetCookies("AdminName") &"]退出成功", "SESSION")
	'清除Cookies
	Call SetCookies("AdminName","")
	Call SetCookies("AdminNickname","")
	Call SetCookies("AdminPassword","")
	Call SetCookies("AdminLevel","")
	Call ConnClose()	'关闭连接
%>
<script type='text/javascript'>alert('成功退出');this.top.location.href='login.asp';</script>
