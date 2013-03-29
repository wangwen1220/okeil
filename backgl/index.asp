<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="admin_conn.asp"-->
<!--#include file="md5.asp"-->
<html>
<head>
<meta content="" name="keywords">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>管理中心</title>
<link type="text/css" href="css.css" rel="stylesheet">
</head>
<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true"bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="if(document.login.Username!=null) document.login.Username.focus();">
<%if session("admin")="" then%>
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#222222" style="border-collapse: collapse">
<tr><td width="100%">
<center>
<b><font color="#ff0000"> 
<%
select case request("err")
case 1
strmsg="登陆成功！！！"
case 2
strmsg="2"
case 3
strmsg="用户名和密码错误！！！"
case 4
strmsg="请输入用户名和密码！！！"
case 5
strmsg="权限不够！！！"
case 6
strmsg="验证码错误！"
case 7
strmsg="验证码没有填写！"
case else
strmsg=""
end select
response.write strmsg
%></font></b></center>
      <table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td background="images/backtop_2.jpg"><div align="center"><img src="images/backtop_3.jpg" width="450" height="60"></div></td>
        </tr>
      </table>
      <table width="500" border="0" align="center" cellpadding="4" cellspacing="1" >
        <form action="checkadmin.asp" method="post" name="login" id="login">
          <tr> 
            <td bgcolor="#E5ECFE"><div align="center">用户名称&nbsp;<input name="Username" type="text" class="box" size="15"></div></td>
          </tr>
          <tr> 
            <td bgcolor="#E5ECFE"><div align="center">用户密码&nbsp;<input name="Password" type="password" class="box" size="15"></div></td>
          </tr>
          <tr> 
            <td bgcolor="#E5ECFE"><div align="center">&nbsp;&nbsp;&nbsp;验&nbsp;证&nbsp;码&nbsp; 
                <input name="yz" type="text"  size="8"><img src="wx_pSN.asp" align="absmiddle"></div></td>
          </tr>
          <tr> 
            <td align="center" bgcolor="#E5ECFE"> <input type="submit" name="Submit3" value="确 认"> 
              &nbsp; <input name="Submit22" type="reset" value="清 除"> </td>
          </tr>
        </form>
      </table>
      <table width="500" border="0" align="center" cellpadding="4" cellspacing="0">
        <tr>
          <td><div align="center">如果验证码看不清楚，请刷新页面!</div></td>
        </tr>
      </table>
    </td>
  </tr></table>
<%else%>
<% response.Redirect("main.asp")%>
<%end if %>
</td></tr></table>
</body>
</html>
