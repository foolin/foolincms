<%Sub Header%>
    <div id="header">
         <img src="images/logo.gif" height="35"  border="0"/> <span style="font-size:32px; font-weight:bold;"><%=SiteName%>管理</span> <span style="font-size:18px; color:#099;">欢迎使用E酷CMS，一个小型站点的内容管理系统~</span>
    </div>
<%End Sub%>

<%Sub TopNav(ByVal act)
	act = LCase(act)
%>
    <ul>
     <li<%If act = "index" Then Echo(" class=""on""")%>><a href="index.asp">首页</a></li>
     <li<%If act = "article" Then Echo(" class=""on""")%>><a href="admin_article.asp">文章</a></li>
     <li<%If act = "picture" Then Echo(" class=""on""")%>><a href="admin_picture.asp">图片</a></li>
     <li<%If act = "guestbook" Then Echo(" class=""on""")%>><a href="admin_guestbook.asp">留言</a></li>
     <!--
     <li<%If act = "comment" Then Echo(" class=""on""")%>><a href="#admin_comment.asp">评论</a></li>
     -->
     <%If Session("AdminLevel") > 1 Then%>
     	<li<%If act = "mytag" Then Echo(" class=""on""")%>><a href="admin_mytag.asp">标签</a></li>
     	<li<%If act = "diypage" Then Echo(" class=""on""")%>><a href="admin_diypage.asp">DIY页面</a></li>
     <%End If%>
     <%If Session("AdminLevel") > 2 Then%>
	 	<li<%If act = "template" Then Echo(" class=""on""")%>><a href="admin_template.asp">模板</a></li>
     	<li<%If act = "user" Then Echo(" class=""on""")%>><a href="admin_user.asp">团队</a></li>
     <%End If%>
     <!--
     <li<%If act = "file" Then Echo(" class=""on""")%>><a href="#admin_uploadfile.asp">文件</a></li>
      -->
      <%If Session("AdminLevel") > 0 Then%>
	 	<li<%If act = "weblog" Then Echo(" class=""on""")%>><a href="admin_weblog.asp">管理日志</a></li>
     <%End If%>
     <%If Session("AdminLevel") > 2 Then%>
     	<li<%If act = "config" Then Echo(" class=""on""")%>><a href="admin_config.asp">系统配置</a></li>
     <%End If%>
     <li<%If act = "password" Then Echo(" class=""on""")%>><a href="modify_password.asp">修改密码</a></li>
     <li<%If act = "help" Then Echo(" class=""on""")%>><a href="help.asp">标签帮助</a></li>
     <li><a href="logout.asp">退出</a></li>
    </ul>
<%End Sub%>

<%Sub Footer()%>
    <div id="footer">
    	<a href="../index.asp" target="_blank">网站首页</a> | <a href="index.asp?action=clearcache">更新缓存</a>  | <a href="help.asp">用户帮助</a>  | <a href="http://www.liufu.org/ling" target="_blank">最新版本</a><br />
       <%=Session("AdminName")%>，欢迎您进入[<%=sitename%>]后台管理。　<br />
     &copy; 2009  Power by <%=studio%> ，System kernel：<%=syslink%><br />
    </div>
	<iframe height="0" width="0" frameborder="0" src="keeponline.asp" scrolling="no"></iframe>
<%End Sub%>

<%Sub MyInfo()%>
        <ul class="menu">
         <li class="mTitle">--== 控制面板 ==--</li>
         <li><a href="index.asp">欢迎[<%=Session("AdminName")%>]！</a></li>
         <li><a href="modify_password.asp">修改资料</a></li>
         <li><a href="logout.asp">退出[<%=Session("AdminName")%>]</a></li>
        </ul>
<%End Sub%>

<%Sub SysInfo()%>
        <dl class="menu">
            <dt>--== 网站信息 ==--</dt>
            <dd>网站:<a href="<%=siteurl%>" target="_blank"><%=sitename%></a></dd>
            <dd>网址:<a href="<%=siteurl%>" target="_blank"><%=siteurl%></a></dd>
            <dd><a href="http://www.eekku.com"><%=syslink%></a></dd>
        </dl>
<%End Sub%>