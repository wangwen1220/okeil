<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
'---------------------确定是否已经登陆-------------------------------
dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
end if
'---------------------确定是否已经登陆-------------------------------
%>

<title>网站后台管理系统</title>
<frameset rows="90,*" frameborder="no" border="0" framespacing="0"> 
<frame src="share/top.html" name="topFrame" scrolling="No" noresize="noresize" id="topFrame" title="topFrame" /> 
	<frameset cols="200,*" frameborder="no" border="0" framespacing="0"> 
	<frame src="admin_left.asp" name="leftFrame" id="leftFrame" title="leftFrame" /> 
	<frame src="index1.asp" name="mainFrame" id="mainFrame" title="mainFrame" /> 
	</frameset>
</frameset><noframes></noframes>
