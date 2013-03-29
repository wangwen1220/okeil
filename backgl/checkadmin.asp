<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="admin_conn.asp"-->
<!--#include file="md5.asp"-->
<%
if request("yz")="" then
Response.Redirect "index.asp?err=7"
end if

if session("GetCode")=Trim(request("yz")) then
	dim sql
	dim rs
	dim username
	dim password
	username=replace(trim(request("username")),"'","")
	password=md5(replace(trim(Request("password")),"'",""))

	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where password='"&password&"' and username='"&username&"'"
'	response.write ""&sql&""
'	response.end
	rs.open sql,conn,1,1
 	if not(rs.bof and rs.eof) then
 		if password=rs("password") then
			session.Timeout = 1440
			session("admin")=rs("username")
			session("flag")=rs("flag")
			session("id")=rs("id")
			session("flagid")=rs("flagid")
			'统计系统中需要用到的参数
			Response.Cookies("feiyue")("Username")=rs("username")
			Response.Cookies("feiyue")("Password")=rs("password")
			'-----------------确定网站版本需要-----------------------
			'如果只要中文版，将session("gb1")=""
			'如果两个都要，将session("gb1")="en"
			'如果只要en，将两个都设为空值，然后在conn.asp中将gb的值设为"en"
			'session("gb")="0"
			'session("gb1")="1"
			
            Response.Redirect "index.asp?err=1"
 		else
            Response.Redirect "index.asp?err=2"
 		end if
	else	
            Response.Redirect "index.asp?err=3"
	end if
        rs.close
	conn.close
	set rs=nothing
	set conn=nothing
else
Response.Redirect "index.asp?err=6"
end if
%>

