<!--#include file="connections/conn_szwebsp.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=kw1%></title>
<meta name="keywords" content="<%=kw2%>">
<meta name="description" content="<%=kw3%>">
<meta name="author" content="<%=kw4%>" >
<meta name="copyright" content="<%=kw5%>">
<meta name="Abstrat" content="<%=kw6%>">
</head>
<body>
<%
if request("p")="1" then 
response.Redirect(gotourl1)
end if

if request("p")="2" then 
response.Redirect(gotourl2)
end if

if request("p")="3" then 
response.Redirect(gotourl3)
end if

if request("p")="4" then 
response.Redirect(gotourl4)
end if

if request("p")="5" then 
response.Redirect(gotourl5)
end if

if request("p")="6" then 
response.Redirect(gotourl6)
end if

if request("p")="7" then 
response.Redirect(gotourl7)
end if

if request("p")="8" then 
response.Redirect(gotourl8)
end if

%>
</body>
</html>
