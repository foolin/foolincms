------------------------------------------------------------
 Plugin Name:	留言本(guestbook)
 Purpose:     	用户留言功能
 Author:		Foolin
 E-mail:		Foolin@126.com
 Create on:		2009-9-27
 Notice:    	本插件是系统自带插件。
 Copyright (c) 2009 E酷工作室（Foolin）All Rights Reserved。
------------------------------------------------------------

一、使用说明：

	1、本插件为系统自带插件，您可以在后台自行配置是否启用该插件。详细操作：登录后台管理，点击【系统配置】→ 其中有【是否开放留言】和【是否需要审核留言】两个选项，自行配置即可。
	
	2、前台显示留言内容方法：
	
	 ① 你需要在您的模板目录中增加一个guestbook.html的模板（最简单的模板方法就是把文章列表的模板复制一份，然后修改相应的列表标签即可）。
	 
	 ② 修改列表标签，需要理解程序系统的基本标签以及使用，你在需要显示留言列表的地方，加入类似以下代码即可：
	 
		标签说明（标签名gbook可以自定义）：
		
            <!--留言列表开始-->
            {list:gbook mode="sql" sql="SELECT * FROM GuestBook WHERE State = 1 ORDER BY ID DESC" col="1" row="20" ispage="true"}
			
				<!-- 内层标签 -->
				
				[gbook:id] 		<!-- 留言ID -->
				
				[gbook:title] 	<!-- 留言标题 -->
				
				[gbook:content clearhtml="true"]	<!-- 留言内容 -->
				
				[gbook:user]	<!-- 留言者名字 -->
				
				[gbook:ip]	<!-- 留言者内容 -->
				
				[gbook:homepage]	<!-- 留言者的个人主页 -->
				
				[gbook:createtime format="yyyy-mm-dd"]	<!-- 留言时间 -->
				
				[gbook:recomment]	<!-- 回复留言内容 -->
				
				[gbook:reuser]	<!-- 回复留言者名字 -->
				
				[gbook:retime]	<!-- 回复留言时间 -->
				
				<!-- 内层标签 -->
				
          {/list:gbook}
          <!--留言列表结束-->
		  
		  
		  {tag:page /} <!--留言列表分页导航-->
		  
		  范例请看：Example_GbookList.txt文件

	 ③ 填写表单
		  留言表单代码请看：Example_PostForm.txt文件，把里面的代码复制过去你的模板页即可。
		  

二、本插件所涉及的文件：

	根目录中的guestbook.asp		--------- 显示文件
	
	模板目录中guestbook.html	--------- 留言簿模板（即是template/您的模板目录名/guestbook.html）
	
	插件目录plugins/save.asp	--------- 保存留言文件
	
	inc/config.asp				--------- 里面有IsOpenGbook（是否开放留言）和IsAuditGbook（是否需要审核留言）两个变量。
	
	admin/admin_config.asp		--------- 里面有设置上面两个变量的操作。
	
	数据表名为：GuestBook		--------- 数据库中表名[GuestBook]，字段请自行查看。
	
	
三、版权声明
	这是系统自带插件，详细协议以及版权声明请看本系统相关文件。


