<%Sub Header%>
    <div id="header">
         <img src="images/logo.gif" height="35"  border="0"/> <span style="font-size:32px; font-weight:bold;"><%=SiteName%>����</span> <span style="font-size:18px; color:#099;">��ӭʹ��E��CMS��һ��С��վ������ݹ���ϵͳ~</span>
    </div>
<%End Sub%>

<%Sub TopNav(ByVal act)
	act = LCase(act)
%>
    <ul>
     <li<%If act = "index" Then Echo(" class=""on""")%>><a href="index.asp">��ҳ</a></li>
     <li<%If act = "article" Then Echo(" class=""on""")%>><a href="admin_article.asp">����</a></li>
     <li<%If act = "picture" Then Echo(" class=""on""")%>><a href="admin_picture.asp">ͼƬ</a></li>
     <li<%If act = "guestbook" Then Echo(" class=""on""")%>><a href="admin_guestbook.asp">����</a></li>
     <!--
     <li<%If act = "comment" Then Echo(" class=""on""")%>><a href="#admin_comment.asp">����</a></li>
     -->
     <%If Session("AdminLevel") > 1 Then%>
     	<li<%If act = "mytag" Then Echo(" class=""on""")%>><a href="admin_mytag.asp">��ǩ</a></li>
     	<li<%If act = "diypage" Then Echo(" class=""on""")%>><a href="admin_diypage.asp">DIYҳ��</a></li>
     <%End If%>
     <%If Session("AdminLevel") > 2 Then%>
	 	<li<%If act = "template" Then Echo(" class=""on""")%>><a href="admin_template.asp">ģ��</a></li>
     	<li<%If act = "user" Then Echo(" class=""on""")%>><a href="admin_user.asp">�Ŷ�</a></li>
     <%End If%>
     <!--
     <li<%If act = "file" Then Echo(" class=""on""")%>><a href="#admin_uploadfile.asp">�ļ�</a></li>
      -->
      <%If Session("AdminLevel") > 0 Then%>
	 	<li<%If act = "weblog" Then Echo(" class=""on""")%>><a href="admin_weblog.asp">������־</a></li>
     <%End If%>
     <%If Session("AdminLevel") > 2 Then%>
     	<li<%If act = "config" Then Echo(" class=""on""")%>><a href="admin_config.asp">ϵͳ����</a></li>
     <%End If%>
     <li<%If act = "password" Then Echo(" class=""on""")%>><a href="modify_password.asp">�޸�����</a></li>
     <li<%If act = "help" Then Echo(" class=""on""")%>><a href="help.asp">��ǩ����</a></li>
     <li><a href="logout.asp">�˳�</a></li>
    </ul>
<%End Sub%>

<%Sub Footer()%>
    <div id="footer">
    	<a href="../index.asp" target="_blank">��վ��ҳ</a> | <a href="index.asp?action=clearcache">���»���</a>  | <a href="help.asp">�û�����</a>  | <a href="http://www.liufu.org/ling" target="_blank">���°汾</a><br />
       <%=Session("AdminName")%>����ӭ������[<%=sitename%>]��̨������<br />
     &copy; 2009  Power by <%=studio%> ��System kernel��<%=syslink%><br />
    </div>
	<iframe height="0" width="0" frameborder="0" src="keeponline.asp" scrolling="no"></iframe>
<%End Sub%>

<%Sub MyInfo()%>
        <ul class="menu">
         <li class="mTitle">--== ������� ==--</li>
         <li><a href="index.asp">��ӭ[<%=Session("AdminName")%>]��</a></li>
         <li><a href="modify_password.asp">�޸�����</a></li>
         <li><a href="logout.asp">�˳�[<%=Session("AdminName")%>]</a></li>
        </ul>
<%End Sub%>

<%Sub SysInfo()%>
        <dl class="menu">
            <dt>--== ��վ��Ϣ ==--</dt>
            <dd>��վ:<a href="<%=siteurl%>" target="_blank"><%=sitename%></a></dd>
            <dd>��ַ:<a href="<%=siteurl%>" target="_blank"><%=siteurl%></a></dd>
            <dd><a href="http://www.eekku.com"><%=syslink%></a></dd>
        </dl>
<%End Sub%>