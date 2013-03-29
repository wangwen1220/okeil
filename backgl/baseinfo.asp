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
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>管理中心——>网站SEO产品管理</title>
<link type="text/css" href="css.css" rel="stylesheet">
</head>
<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#222222" style="border-collapse: collapse">
  <tr> 
    <td height="28">当前位置: <a href="main.asp">网站管理中心</a>--网站SEO管理</td>
  </tr>
  <tr> 
    <td width="100%" valign="top">
	<!--#include file="choosegb.asp"--> 
     <%
	set rs=server.createobject("adodb.recordset")
	sql="select top 1 * from base_info where gb='"&gb&"' order by id desc " 
	rs.open sql,conn,1,3
		do while not rs.eof 
	 %>
        <table width="98%" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#cccccc" bgcolor="#FFFFFF" style="border-collapse: collapse;border:solid 1px">
        <form name="infoadd" method="post" action="baseinfo_save.asp"  onSubmit="FrmAdd();">
          <tr bgcolor="#CBE9FC"> 
            <td height="25" colspan="2"> <div align="center"><strong>网站SEO参数设置修改</strong></div></td>
          </tr>
          <tr bgcolor="f5f5f5"> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td width="21%"><div align="right">网站标题栏：</div></td>
            <td width="79%"><textarea name="A_title" cols="50" rows="3" id="A_keywords"><%=rs("A_title")%></textarea></td>
          </tr>
          <tr bgcolor="f5f5f5"> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td width="21%"><div align="right">网站关键字：</div></td>
            <td width="79%"><textarea name="A_keywords" cols="50" rows="3" id="A_keywords"><%=rs("A_keywords")%></textarea></td>
          </tr>
          <tr bgcolor="f5f5f5"> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td><div align="right">网站描述：</div></td>
            <td><textarea name="A_meta" cols="50" rows="3" id="A_meta"><%=rs("A_meta")%></textarea></td>
          </tr>
          <tr bgcolor="f5f5f5"> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td><div align="right">网站底部信息描述：</div></td>
            <td><textarea name="A_notes" cols="78" rows="4" id="textarea2"><%=rs("A_notes")%></textarea></td>
          </tr>
          <tr bgcolor="f5f5f5"> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td><div align="right">公司名称：</div></td>
            <td><input name="A_company" type="text" id="A_company" value="<%=rs("A_company")%>" size="60"></td>
          </tr>
          <tr> 
            <td><div align="right">电话：</div></td>
            <td><input name="A_tel" type="text" id="A_tel" value="<%=rs("A_tel")%>" size="30"></td>
          </tr>
          <tr> 
            <td><div align="right">传真：</div></td>
            <td><input name="A_fax" type="text" id="A_fax" value="<%=rs("A_fax")%>" size="30"></td>
          </tr>
          <tr> 
            <td><div align="right">邮箱：</div></td>
            <td><input name="A_mail" type="text" id="A_mail" value="<%=rs("A_mail")%>" size="30"></td>
          </tr>
          <tr> 
            <td> <div align="right">地址：</div></td>
            <td><input name="A_address" type="text" id="A_address" value="<%=rs("A_address")%>" size="60"></td>
          </tr>
          <tr> 
            <td colspan="2"><div align="center"> 
				<input type="hidden" name="gb" value="<%=gb%>">
                <input type="submit" name="Submit" value="确认提交">
              </div></td>
          </tr>
          <tr> 
            <td colspan="2">&nbsp;</td>
          </tr>
        </form>
      </table>
        <br>
	  <%
	  rs.movenext
	  loop
	  rs.close
	  set rs = nothing	 
	  %>
	  </td>
  </tr>
</table>
</body>
</html>