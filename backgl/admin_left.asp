<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link type="text/css" href="css.css" rel="stylesheet">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0"  bgcolor="#799AE1" cellspacing="0" cellpadding="0" width="100%" height=100% style="border-collapse: collapse">
  <tr>
    <td valign="top"> 
      <table width="158" border="0" align="center" cellpadding="2" cellspacing="1">
        <tr> 
          <td height="38" background="images/title.jpg" class="bigclass"></td>
        </tr>
        <tr>
          <td height="23" bgcolor="#E5ECFE"><div align="center"><strong>欢迎&nbsp;<%=session("admin")%> ！</strong></div></td>
        </tr>
        <tr> 
          <td height="25" bgcolor="#E5ECFE"><div align="center"><a href='quit.asp'target='_top'>退出后台管理</a></div></td>
        </tr>
      </table>
      <table width="158" border="0" align="center" cellpadding="2" cellspacing="1" style="margin:10px; 0px;">
        <tr> 
          <td height="25" background="images/admin_left_1.gif"><div align="left" class="whitetext">功能模块设置</div></td>
        </tr>
        <%If InStr(session("flag"),"prod")> 0 Then %>
        <tr> 
          <td height="25" bgcolor="#E5ECFE">&nbsp;<img src="images/bullet.gif" width="15" height="20" border="0" align="absmiddle">&nbsp;<a href="admin_prodclass.asp" target="mainFrame">产品类别管理</a></td>
        </tr>

        <tr> 
          <td height="25" bgcolor="#E5ECFE">&nbsp;<img src="images/bullet.gif" width="15" height="20" border="0" align="absmiddle">&nbsp;<a href="admin_prod.asp" target="mainFrame">产品列表管理</a></td>
        </tr>
        <%end if%>
		
        <%If InStr(session("flag"),"link")> 0 Then %>
        <tr> 
          <td height="25" bgcolor="#E5ECFE">&nbsp;<img src="images/bullet.gif" width="15" height="20" border="0" align="absmiddle">&nbsp;<a href="admin_link.asp" target="mainFrame">友情链接管理</a></td>
        </tr>
        <%end if%>
        <%If InStr(session("flag"),"banner")> 0 Then %>
        <tr> 
          <td height="25" bgcolor="#E5ECFE">&nbsp;<img src="images/bullet.gif" width="15" height="20" border="0" align="absmiddle">&nbsp;<a href="banneradd.asp" target="mainFrame">动画图片管理</a></td>
        </tr>
        <%end if%>
      </table>
      <table width="158" border="0" align="center" cellpadding="2" cellspacing="1" style="margin:10px; 0px;">
        <tr> 
          <td height="25" background="images/admin_left_2.gif"><div align="left" class="whitetext">系统资源设置区</div></td>
        </tr>
		<%If InStr(session("flag"),"seo")> 0 Then %>
        <tr> 
          <td height="25" bgcolor="#E5ECFE">&nbsp;<img src="images/bullet.gif" width="15" height="20" border="0" align="absmiddle">&nbsp;<a href="baseinfo.asp" target="mainFrame">网站优化管理</a></td>
        </tr>
		<%end if%>
		<%If InStr(session("flag"),"tem")> 0 Then %>
        <tr> 
          <td height="25" bgcolor="#E5ECFE">&nbsp;<img src="images/bullet.gif" width="15" height="20" border="0" align="absmiddle">&nbsp;<a href="admin_backup.asp" target="mainFrame">数据库备份与还原</a></td>
        </tr>
		<%end if%>
		<%If InStr(session("flag"),"password")> 0 Then %>
        <tr> 
          <td height="25" bgcolor="#E5ECFE">&nbsp;<img src="images/bullet.gif" width="15" height="20" border="0" align="absmiddle">&nbsp;<a href="admin_index.asp" target="mainFrame">管理员密码修改</a></td>
        </tr>
		<%end if%>
		<%If InStr(session("flag"),"manager")> 0 Then %>
        <tr> 
          <td height="25" bgcolor="#E5ECFE">&nbsp;<img src="images/bullet.gif" width="15" height="20" border="0" align="absmiddle">&nbsp;<a href="admin_manager.asp" target="mainFrame">管理员权限设置</a></td>
        </tr>
        <%end if %>
      </table>
      <table width="158" border="0" align="center" cellpadding="2" cellspacing="1" style="margin:10px; 0px;">
        <tr> 
          <td width="1195" height="25" background="images/admin_left_3.gif"><div align="left" class="whitetext">版权声明V4.02</div></td>
        </tr>
        <tr> 
          <td height="25" bgcolor="#E5ECFE"><div align="left"><a href="http://www.szwebcn.com" target="_blank">本网站后台管理系统由灵瑞开发，未经允许不得复制或盗用。<br>
              </a>技术服务电话0755-:86379013<br>
              灵瑞网络:<a href="http://www.szwebcn.com">www.szwebcn.com</a></div></td>
        </tr>
      </table></td>
</tr>
</table>
