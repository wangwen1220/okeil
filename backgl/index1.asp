<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link type="text/css" href="css.css" rel="stylesheet">
<title>网站后台管理系统</title><body onmouseover="self.status='欢迎使用网站后台管理系系统';return true" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="document.login.Username.focus();">
<table width="75%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
  <tr> 
    <td valign="middle"> <center>
        <table width=70% height="191" border="0" align="center" cellpadding="5" cellspacing="1" class="sft">
          <tr> 
            <td height="35" background="images/backtop_2.jpg"> <div align="center" class="toptext">管理员登陆成功</div></td>
          </tr>
          <tr> 
            <td width="100%" height="30" bgcolor="#E5ECFE"><strong> 
              <%=session("admin")%> :欢迎您使用网站管理系统V4.02版！</strong></td>
          </tr>
          <tr> 
            <td height="30" bgcolor="#E5ECFE">*当前时间:<%=now()%></td>
          </tr>
          <tr> 
            <td height="30" bgcolor="#E5ECFE">*您的IP:<%=Request.ServerVariables("REMOTE_ADDR")%></td>
          </tr>
          <tr> 
            <td height="30" bgcolor="#E5ECFE">*服务器名称:<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
          </tr>
          <tr> 
            <td height="30" bgcolor="#E5ECFE">现在您可以从左边列表中选择链接来管理网站了</td>
          </tr>
        </table>
        <br>
      </center>
      </td>
  </tr>
</table>
</body>
</html>
