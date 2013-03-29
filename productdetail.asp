<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="x-ua-compatible" content="ie=7" />
<!--#include file="connections/conn_szweben.asp"-->
<%
if Request("parentid")<>"" then ptitle=conn.execute("select classname from pclass where id="&request("parentid"))(0)
if Request("parentid")<>"" then ParentID=conn.execute("select ParentID from pclass where id="&request("parentid"))(0)
if ParentID <>"" then ptitle1=conn.execute("select classname from pclass where id="&ParentID)(0)
if Request("prodnum")<>"" then pname=conn.execute("select prodid from prodmain where prodnum="&request("prodnum"))(0)
%>
<title><%=pname%>|<%=ptitle1%>|<%=ptitle%>|<%=kw1%></title>
<meta name="keywords" content="<%=kw2%>">
<meta name="description" content="<%=kw3%>">
<meta name="author" content="<%=kw4%>" >
<meta name="copyright" content="<%=kw5%>">
<meta name="Abstrat" content="<%=kw6%>">
<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript" src="js/jquery-1.6.4.min.js"></script>
<script type="text/javascript" src="js/ready.min.js"></script>
<link href="css/base_min.css" rel="stylesheet" type="text/css" />
<link href="css/main.css" rel="stylesheet" type="text/css" />
<link rel="shortcut icon" type="image/x-icon" href="/favicon.ico">
</head>

<body id="home">
  <div id="wrap">
    <!--#include file="top.asp"-->

    <div id="content">
      <div id="main">
        
	   <!--#include file="banner.asp"-->
        <h2 id="position">Welcome to <a href="/index.asp" title="LIEKO Mobile Phone Case">LIEKO</a><% if ptitle1<>"" then%>&nbsp;/&nbsp;<%=ptitle1%>&nbsp;/&nbsp;<%end if%><% if ptitle<>"" then%><%= ptitle%>&nbsp;/&nbsp;<%end if%></h2>

        <div id="pros" class="prolist">
          <h1><%=pname%></h1>
          <p class="notice">How are you, there is some products list, about more products informations, please contact our sales representive thank you.</p>
          </div>  
		  <div class="product_details">
		       <!--#include file="sharenew/proddetail.asp" -->
		  </div>
        
      <!--#include file="hotprod.asp"--> 
      </div>

     <!--#include file="left.asp"-->
    </div>
  </div>
	<!--#include file="end.asp"-->
</body>
</html>
