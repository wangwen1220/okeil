<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="admin_conn.asp"-->
<!--#include file="md5.asp"-->
<%
username=request("username")
id=request("id")
action=request("action")

dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
else
	if session("flag")<>"0" then
	
		cls = Instr(session("flag"), "manager")
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
<title>管理中心——>管理权限设置</title>
<link type="text/css" href="css.css" rel="stylesheet">
<SCRIPT LANGUAGE="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')  e.checked = form.chkall.checked; 
   }
  }
//-->
</SCRIPT>
</head>
<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true"bgcolor=#ffffff leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#222222" style="border-collapse: collapse">
        <tr> 
          <td width="100%" height="28">当前位置:<a href="index.asp">网站管理中心</a>--管理权限设置</td>
        </tr>
        <tr> 
          <td width="100%"> <%
	 if action="" then
     Set rs = conn.Execute("select * from admin order by id")  
	 %> <table width="95%" border="1"  style="border-collapse: collapse;border:solid 1px" bordercolor="#cccccc"  cellspacing="0" cellpadding="4" align="center">
              <tr> 
                <td width=140 height="28" bgcolor="#E5ECFE"><strong>管理员登陆号</strong></td>
                <td width="328" bgcolor="#E5ECFE"><strong>使用密码(经过加密处理)</strong></td>
                <td width="100" bgcolor="#E5ECFE"> <div align="left"><strong>管理员级别</strong></div></td>
                <td width="72" bgcolor="#E5ECFE"><strong>编辑权限</strong></td>
                <td width=100 bgcolor="#E5ECFE"><strong>删除</strong> </td>
              </tr>
              <%	
	   do while not rs.eof
		if  session("id")=rs("id") then
	%>
              <tr> 
                <td height="25"><%=rs("username")%></td>
                <td><%=rs("password")%><%=left(rs("note"),(cint(len(rs("note"))-21)))%></td>
                <td><% if rs("flagid")="9" then %>
                  超级管理员
                  <%else%>
                  管理员
                  <% end if %></td>
                <td><a href="admin_manager.asp?action=detail&username=<%=rs("username")%>&id=<%=rs("id")%>">修改</a></td>
                <td><font color="#a4a4a4">当前用户</font></td>
              </tr>
              <%else%>
              <tr <%if rs("username")="admin" then %> style="display:none;"<%end if%>> 
                <td height="25"><%=rs("username")%></td>
                <td><%=rs("password")%><%=left(rs("note"),(cint(len(rs("note"))-21)))%></td>
                <td><% if rs("flagid")="9" then %>
                  超级管理员
                  <%else%>
                  管理员
                  <% end if %></td>
                <td><a href="admin_manager.asp?action=detail&username=<%=rs("username")%>&id=<%=rs("id")%>">修改</a></td>
                <td><a href="admin_manager.asp?action=del&username=<%=rs("username")%>&id=<%=rs("id")%>">删除</a></td>
              </tr>
              <%
	end if
	rs.movenext   
	loop    
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	%>
            </table>
            <br> <table width="95%" border="1"  style="border-collapse: collapse;border:solid 1px" bordercolor="#cccccc"  cellspacing="0" cellpadding="4" align="center">
              <form action="admin_manager.asp?action=adduserpost" name="list" method="post">
                <tr bgcolor="#E5ECFE"> 
                  <td height="28" colspan=2 align=center>添加管理员</td>
                </tr>
                <tr> 
                  <td width=100 height="25">管理员名称</td>
                  <td><input type="text" name="username" size="11">
                    级别： 
                    <select size="1" name="D1">
                      <option value="8" selected>管理员</option>
                      <option  value="9">超级管理员</option>
                    </select> </td>
                </tr>
                <tr> 
                  <td width=100 height="25">登陆密码</td>
                  <td><input type="text" name="password" size="11"> </td>
                </tr>
                <tr> 
                  <td height="25" colspan=2> <input type="hidden" name="flag" value=""> 
                    <input type="submit" name="adduser" value="增加管理员"> </td>
                </tr>
              </form>
            </table>
            <br> <table width="95%" border="0" align="center" cellpadding="4" cellspacing="0">
              <tr> 
                <td><strong><font color="#FF0000">特别声明：</font></strong></td>
              </tr>
              <tr> 
                <td><font color="#FF0000">1.登陆密码会经过加密处理，所以请记住您在此填写的密码。</font></td>
              </tr>
              <tr> 
                <td><font color="#FF0000">2.现在所看到的密码是还原了的用户密码。</font></td>
              </tr>
              <tr> 
                <td><font color="#FF0000">3.只能给超级管理员[管理权限设置]功能，才能看到密码和进行权限的修改，所以请不要随便开放[管理权限设置]功能给一般管理员。</font></td>
              </tr>
            </table>
            <%else 
end if%> <%
'显示权限详细信息
if action="detail" then
	set rs=conn.execute("select * from admin where username='"&username&"'") 
	if not rs.eof then
	%> <table border="1" style="border-collapse:collapse" bordercolor="#cccccc" width="95%" align="center">
              <form action="admin_manager.asp" method="post" name="modify">
                <tr> 
                  <td height="25" bgcolor="#E5ECFE">用户名</td>
                  <td bgcolor="#E5ECFE">拥有权限</td>
                </tr>
                <tr> 
                  <td valign="top"><%=rs("username")%></td>
                  <td> <%
			set rs=conn.Execute("select * from admin where username='"&username&"'")
			dim gradeType,sheet		
			gradeType="产品中心管理,友情链接管理,其它栏目管理,动画图片管理,管理权限设置,密码修改功能,数据库管理,网站优化管理"
			gradeCode="prod,link,wz,banner,manager,password,tem,seo"
			code=Split(gradeCode,",")
			sheet=Split(gradeType,",")
			for i=0 to ubound(sheet)
				response.write "<input type=""checkbox"" name=""flag"" value="""&trim(code(i))&"""" 
				if instr(rs("flag"),trim(code(i)))>0 then		'如果有此项权利；
				response.write " checked" 
				end if
				response.write ">"&trim(sheet(i))&""
				if ((i+1) mod 1)=0 then response.write "<br>"	'每行显示1个权限
			next
			response.write character %> </td>
                </tr>
                <tr> 
                  <td> <input type="hidden" name="action" value="modify"> <input type="hidden" name="username" value=<%=rs("username")%>> 
                  </td>
                  <td> <input type="submit" name="ok" value="提交"> </td>
                </tr>
              </form>
            </table>
            <%else
	response.write "没有找到合适的纪录"
	end if
	rs.close
	conn.close
	else
end if%> <%
'修改 后返回参数提交数据库
if action="modify" then
	flag=checkreal(request.form("flag"))	
     Set rs=Server.CreateObject("ADODB.Recordset")
	 sql="select * from admin where username='"&username&"'"
	 rs.open sql,conn,3,3
	 rs("flag")=Server.Htmlencode(flag)
	
	' response.Write Server.Htmlencode(flag)
	' response.End()
	 
	 rs.update
    url="admin_manager.asp?action=detail&username="&rs("username")
    rs.close
	conn.close
	set rs=nothing
	set conn=nothing
	response.write "<script language='javascript'>"
	response.write "alert('权限设置修改成功！');"
	response.write "location.href='admin_manager.asp';"			
	response.write "</script>"
	'Response.Redirect"admin_manager.asp?action=detail&username="&Request("username")
	else
end if%> <%
'删除 后返回参数提交数据库
if action="del" then
	username=request("username")
	id=request("id")
    Dim StrSQL
    StrSQL="delete from admin where username='"&username&"'"
    conn.Execute StrSQL    
    conn.close
    Response.Redirect"admin_manager.asp" 
else
end if%> </td>
        </tr>
      </table>
</body>
</html>


<%if action="adduserpost" then
	if request("username")="" or request("password")="" then
		response.write "<script language='javascript'>"
		response.write "alert('无效输入！');"
		response.write "location.href='admin_manager.asp';"			
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

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select * from admin"
		rs.Open sql,conn,1,3
		rs.Addnew
		rs("username")=Server.Htmlencode(Request("username"))
		rs("password")=md5(Server.Htmlencode(Request("password")))
		rs("flag")=Server.Htmlencode(Request("flag"))
		rs("flagid")=Server.Htmlencode(Request("d1"))
		rs("note")=request("password")+inbillNo
		rs.Update
		rs.Close
		Set rs=Nothing
		Response.Redirect"admin_manager.asp?action=detail&username="&Request("username")
	end if
else
end if%>

<%
'处理数组函数
function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,",","|")
	w=replace(w," ","")
	checkreal=w
end if
end function
%>
