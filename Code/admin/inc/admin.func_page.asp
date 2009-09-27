<%Sub Header%>
    <div id="header">
        E酷CMS 一个简捷而不简单的内容管理系统。努力而不后悔，↖(^ω^)↗加油！
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
     <li<%If act = "comment" Then Echo(" class=""on""")%>><a href="#admin_comment.asp">评论</a></li>
     <li<%If act = "mytag" Then Echo(" class=""on""")%>><a href="admin_mytag.asp">标签</a></li>
     <li<%If act = "diypage" Then Echo(" class=""on""")%>><a href="admin_diypage.asp">DIY页面</a></li>
<li<%If act = "webftp" Then Echo(" class=""on""")%>><a href="admin_webftp.asp">模板</a></li>
     <li<%If act = "user" Then Echo(" class=""on""")%>><a href="admin_user.asp">团队</a></li>
     <li<%If act = "file" Then Echo(" class=""on""")%>><a href="#admin_uploadfile.asp">文件</a></li>
	 <li<%If act = "weblog" Then Echo(" class=""on""")%>><a href="admin_weblog.asp">管理日志</a></li>
     <li<%If act = "config" Then Echo(" class=""on""")%>><a href="admin_config.asp">系统配置</a></li>
     <li<%If act = "password" Then Echo(" class=""on""")%>><a href="modify_password.asp">修改密码</a></li>
     <li<%If act = "help" Then Echo(" class=""on""")%>><a href="help.asp">帮助</a></li>
     <li><a href="logout.asp">退出</a></li>
    </ul>
<%End Sub%>

<%Sub Footer()%>
    <div id="footer">
    	<a href="../index.asp" target="_blank">网站首页</a> | <a href="http://www.eekku.com">官方网站</a> | <a href="help.asp">用户帮助</a>   | <a href="index.asp?action=clearcache">更新缓存</a> <br />
        Author:Foolin  E-mail:Foolin@126.com HomePage: http://www.eekku.com<br />
        Copyright &copy; 2009　www.eekku.com　版权所有　<br />
    </div>
	<iframe height="30" width="100%" frameborder="0" src="keeponline.asp" scrolling="no" style="display:none;"></iframe>
<%End Sub%>

<%Sub MyInfo()%>
        <ul class="menu">
         <li class="mTitle">--== 控制面板 ==--</li>
         <li><a href="index.asp">欢迎[Foolin]！</a></li>
         <li><a href="modify_password.asp">修改资料</a></li>
         <li><a href="logout.asp">退出[Foolin]</a></li>
        </ul>
<%End Sub%>

<%Sub SysInfo()%>
        <dl class="menu">
            <dt>--== 版权信息 ==--</dt>
            <dd>制作：E酷工作室(Foolin)</dd>
            <dd>E-mail：Foolin@126.com</dd>
            <dd><a href="http://www.eekku.com">官网：http://www.eekku.com</a></dd>
            <dd><a href="http://www.eekku.com">E酷CMS V1.0.0 build0901</a></dd>
        </dl>
<%End Sub%>