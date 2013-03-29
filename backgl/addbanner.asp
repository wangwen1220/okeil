<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="conn.asp"-->
<%
'---------------------确定是否已经登陆-------------------------------
dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
end if
'---------------------确定是否已经登陆-------------------------------
%>

<%
set rs=server.createobject("adodb.recordset")
sql="select * from banner where gb='"&request("gb")&"'" 
rs.open sql,conn,1,3
		rs.Addnew
		rs("gb")=trim(request.Form("gb"))
		rs("pic1")=trim(request.Form("photo"))
		rs("pic2")=trim(request.Form("photo2"))
		rs("pic3")=trim(request.Form("photo3"))
		rs("pic4")=trim(request.Form("photo4"))
		rs("pic5")=trim(request.Form("photo5"))
		rs("pic6")=trim(request.Form("photo6"))
		rs("pic7")=trim(request.Form("photo7"))
		rs("pic8")=trim(request.Form("photo8"))
		
		rs("link1")=trim(request.Form("link1"))
		rs("link2")=trim(request.Form("link2"))
		rs("link3")=trim(request.Form("link3"))
		rs("link4")=trim(request.Form("link4"))
		rs("link5")=trim(request.Form("link5"))
		rs("link6")=trim(request.Form("link6"))
		rs("link7")=trim(request.Form("link7"))
		rs("link8")=trim(request.Form("link8"))
		
		'rs("notes1")=trim(request.Form("notes1"))
		'rs("notes2")=trim(request.Form("notes2"))
		'rs("notes3")=trim(request.Form("notes3"))
		rs.update	
		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
'		Response.Redirect"admin_prod.asp?action=detail&id="&request.form("ProdId")
		response.Write "<script language=JavaScript>{window.alert('图片修改成功！');window.location.href='banneradd.asp?gb="&gb&"';}</script>"	
		response.end
%>