<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="admin_conn.asp"-->
<!--#include file="md5.asp"-->
<%
dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
else
	if session("flag")<>"0" then
	
		cls = Instr(session("flag"), "password")
		if cls <= 0 then
	%>
		<script language="javascript">
			if (confirm("您的操作权限不够,系统拒绝你的访问,请点确定返回,或者点取消退出重新登录"))
			  location.href="index.asp?err=5";
			else
			  location.href="quit.asp";
		</script>
	<%
		end if
	end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>管理中心——>管理员密码修改</title>
<link type="text/css" href="css.css" rel="stylesheet">
<script language="JavaScript" type="text/JavaScript">
function checksend(){
if (document.repl.password.value==""){
	alert("您没有填写密码。");
	document.repl.password.focus();
	document.repl.password.value="";
	return false;
	}
if (document.repl.password2.value==""){
	alert("请再次输入您的密码！");
	document.repl.password2.focus();
	document.repl.password2.value="";
	return false;
	}
if (document.repl.password2.value != document.repl.password.value){
	alert("两次输入的密码不一致！");
	document.repl.password2.focus();
	document.repl.password2.value="";
	return false;
	}
	return true;
}
</script>
</head>
<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true"bgcolor=#ffffff leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#cccccc" style="border-collapse: collapse">
        <tr> 
          
    <td width="50%" height="28">当前位置:网站管理中心--<span class="blod">管理员密码修改</span></td>
        </tr>
      </table>
      <table width="98%" height="80%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#cccccc" style="border-collapse: collapse">
        <tr> 
          <td valign="middle">
		 <%
		id=request.form("id")
		action=request("action")
		if action="modifypost" then
		
			Set rs=Server.CreateObject("ADODB.Recordset")
			sql="select * from admin where id="&id
			rs.open sql,conn,1,3
			if rs.eof and rs.bof then
			response.write "<script language='javascript'>"
			response.write "alert('用户不存在，或者被删除了！');"
			response.write "location.href='index.asp?action=password';"			
			response.write "</script>"
		
			else
			'----------------------产生一个随机数 -------------------------
				dim num1
				dim rndnum
				Randomize
				for i=0 to 6
				num1=CStr(Chr((57-48)*rnd+48))
				rndnum=rndnum&num1
				next
				'日期，格式：YYYYMMDD
				yy=year(date)
				mm=right("00"&month(date),2)
				dd=right("00"&day(date),2)
				riqi=yy & mm & dd
				
				'生成元素,格式为：小时，分钟，秒
				xiaoshi=right("00"&hour(time),2)
				fenzhong=right("00"&minute(time),2)
				miao=right("00"&second(time),2)
				shijian=xiaoshi & fenzhong & miao
				'生成随机数
				inBillNo=riqi & shijian & rndnum
			'-------------------------end----------------------------------
		
			rs("username")=Server.Htmlencode(Request("username"))
			rs("password")=md5(Server.Htmlencode(Request("password")))
			rs("note")=riqi+request("password")+inbillNo
			rs.update
			end if
			
			rs.close
			conn.close
			set rs=nothing
			set conn=nothing
			response.write "<script language='javascript'>"
			response.write "alert('密码更新成功！\n\n请记住："&Request("password")&"');"
			response.write "location.href='admin_index.asp?action=password';"			
			response.write "</script>"
		else
				username=session("admin")
				Set rs = conn.Execute("select * from admin where username='"&username&"'")
			%> 
      <table width="650" border="0" align="center" cellpadding="4" cellspacing="1">
              <form  name="repl" method="post" action="admin_index.asp?action=modifypost" onsubmit="return checksend();">
                <tr> 
                  <td height="35" colspan="2" background="images/backtop_2.jpg"> <div align="center" class="toptext">管理员密码修改</div></td>
                </tr>
                <tr> 
                  <td width="243" height="25" bgcolor="#E5ECFE"> <div align="right">管理员名称</div></td>
                  <td width="391" bgcolor="#E5ECFE"><%=rs("username")%></td>
                </tr>
                <tr> 
                  <td height="25" bgcolor="#E5ECFE"> <div align="right">请输入新密码</div></td>
                  <td bgcolor="#E5ECFE"> <input name="password" type="password" size="20" maxlength="12" >
                    * <font color="#FF0000">6-12位的数字或英文字母</font></td>
                </tr>
                <tr> 
                  <td height="25" bgcolor="#E5ECFE"> <div align="right">请再输入一次新密码</div></td>
                  <td bgcolor="#E5ECFE"> <input name="password2" type="password" size="20" maxlength="12"> 
                    <input name="id" type="hidden" value=<%=rs("id")%>> 
                    <input name="username" type="hidden" value="<%=rs("username")%>"> 
                  </td>
                </tr>
                <tr> 
                  <td height="25" colspan="2" bgcolor="#E5ECFE"> <div align="center"> 
                      <input type="submit" name="Submit" value="确 认">
                      &nbsp;&nbsp; 
                      <input type="reset" name="Submit2" value="清 除">
                      &nbsp;&nbsp;&nbsp;<a href="javascript:history.go(-1);">返回上一页</a></div></td>
                </tr>
                <tr> 
                  
            <td height="25" colspan="2" bgcolor="#E5ECFE" class="bigclass">说明：</td>
                </tr>
                <tr> 
                  
            <td height="25" colspan="2" bgcolor="#E5ECFE">1.输入的密码会经过加密处理，所以请记住您在此输入的密码！</td>
                </tr>
                <tr> 
                  
            <td height="25" colspan="2" bgcolor="#E5ECFE"> 2.如果忘记管理员的密码，请与超级管理员联系。</td>
                </tr>
              </form>
            </table>
            <%   
		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
		end if
	    %> </td>
        </tr>
      </table>
</body>
</html>
