<%Sub Header%>
    <div id="header">
        E��CMS һ����ݶ����򵥵����ݹ���ϵͳ��Ŭ��������ڣ��I(^��^)�J���ͣ�
    </div>
<%End Sub%>

<%Sub TopNav(ByVal act)
	act = LCase(act)
%>
    <ul>
     <li<%If act = "index" Then Response.Write(" class=""on""")%>><a href="index.asp">��ҳ</a></li>
     <li<%If act = "article" Then Response.Write(" class=""on""")%>><a href="admin_article.asp">����</a></li>
     <li<%If act = "picture" Then Response.Write(" class=""on""")%>><a href="admin_picture.asp">ͼƬ</a></li>
     <li<%If act = "guestbook" Then Response.Write(" class=""on""")%>><a href="admin_guestbook.asp">����</a></li>
     <li<%If act = "comment" Then Response.Write(" class=""on""")%>><a href="admin_comment.asp">����</a></li>
     <li<%If act = "mytag" Then Response.Write(" class=""on""")%>><a href="admin_mytag.asp">��ǩ</a></li>
     <li<%If act = "diypage" Then Response.Write(" class=""on""")%>><a href="admin_diypage.asp">DIYҳ��</a></li>
     <li<%If act = "user" Then Response.Write(" class=""on""")%>><a href="admin_user.asp">�Ŷӹ���</a></li>
     <li<%If act = "config" Then Response.Write(" class=""on""")%>><a href="admin_config.asp">ϵͳ����</a></li>
     <li<%If act = "file" Then Response.Write(" class=""on""")%>><a href="admin_uploadfile.asp">�ϴ��ļ�</a></li>
     <li<%If act = "weblog" Then Response.Write(" class=""on""")%>><a href="admin_weblog.asp">������¼</a></li>
     <li<%If act = "password" Then Response.Write(" class=""on""")%>><a href="modify_password.asp">�޸�����</a></li>
     <li<%If act = "help" Then Response.Write(" class=""on""")%>><a href="help.asp">����</a></li>
     <li><a href="logout.asp">�˳�</a></li>
    </ul>
<%End Sub%>

<%Sub Footer()%>
    <div id="footer">
    	<a href="../index.asp">��վ��ҳ</a> | <a href="http://www.eekku.com">�ٷ���վ</a> | <a href="help.asp">�û�����</a> <br />
        Author:Foolin  E-mail:Foolin@126.com HomePage: http://www.eekku.com<br />
        Copyright &copy; 2009��www.eekku.com����Ȩ���С�<br />
    </div>
	<iframe height="30" width="100%" frameborder="0" src="keeponline.asp" scrolling="no" style="display:none;"></iframe>
<%End Sub%>

<%Sub MyInfo()%>
        <ul class="menu">
         <li class="mTitle">--== ������� ==--</li>
         <li><a href="index.asp">��ӭ[Foolin]��</a></li>
         <li><a href="modify_password.asp">�޸�����</a></li>
         <li><a href="logout.asp">�˳�[Foolin]</a></li>
        </ul>
<%End Sub%>

<%Sub SysInfo()%>
        <dl class="menu">
            <dt>--== ��Ȩ��Ϣ ==--</dt>
            <dd>������E�Ṥ����(Foolin)</dd>
            <dd>E-mail��Foolin@126.com</dd>
            <dd><a href="http://www.eekku.com">������http://www.eekku.com</a></dd>
            <dd><a href="http://www.eekku.com">E��CMS V1.0.0 build0901</a></dd>
        </dl>
<%End Sub%>