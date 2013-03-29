<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<%
'---------------------确定是否具有管理产品的权利-------------------------------
dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
else
	if session("flag")<>"0" then
	
		cls = Instr(session("flag"), "seo")
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
'---------------------确定是否具有管理产品的权利-------------------------------
%>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from base_info" 
rs.open sql,conn,1,3
		rs.Addnew
		rs("A_title")=trim(request.Form("A_title"))
		rs("A_keywords")=trim(request.Form("A_keywords"))
		rs("A_meta")=trim(request.Form("A_meta"))
		rs("A_notes")=trim(request.Form("A_notes"))
		rs("A_company")=trim(request.Form("A_company"))
		rs("A_tel")=trim(request.Form("A_tel"))
		rs("A_fax")=trim(request.Form("A_fax"))
		rs("A_mail")=trim(request.Form("A_mail"))
		rs("A_address")=trim(request.Form("A_address"))
		rs("gb")= trim(request.Form("gb"))
		rs.update	
		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
'		Response.Redirect"admin_prod.asp?action=detail&id="&request.form("ProdId")
		response.Write "<script language=JavaScript>{window.alert('信息修改成功！');window.location.href='baseinfo.asp?gb="&gb&"';}</script>"	
		response.end
%>