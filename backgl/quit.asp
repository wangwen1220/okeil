<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	session("admin")=""
	session("flag")=""
	session("id")=""
	session("midclass")=""
	session("keyword")=""
	session("online")=""
	session("gb")=""
	'统计系统中需要用到的参数
	Response.Cookies("feiyue")("Username")=""
	Response.Cookies("feiyue")("Password")=""
	Response.Redirect "index.asp"
%>