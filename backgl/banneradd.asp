<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<%
dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
else
	if session("flag")<>"0" then
	
		cls = Instr(session("flag"), "banner")
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
<title>管理中心——>产品管理</title>
<link type="text/css" href="css.css" rel="stylesheet">
</head>

<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#222222" style="border-collapse: collapse">
  <tr> 
    <td height="28">当前位置: <a href="main.asp">网站管理中心</a>-动画图片管理</td>
  </tr>
  <tr> 
    <td width="100%" valign="top"> 
	<!--#include file="choosegb.asp"--><br><br>
    <%
	if request("gb")<>"" then 
		gb=request("gb")
	end if 
	set rs=server.createobject("adodb.recordset")
	sql="select top 1 * from banner where gb='"&gb&"' and pic1<>'' order by id desc " 
	rs.open sql,conn,1,3
		do while not rs.eof 
	 %>
      <table width="80%" border="1" align="center" cellpadding="3" cellspacing="0" bordercolor="#cccccc" bgcolor="#FFFFFF" style="border-collapse: collapse;border:solid 1px">
		<form name="prodtable" method="post" action="addbanner.asp"  onSubmit="FrmAdd();">
          <tr bgcolor="#e5ecfe"> 
            <td height="25" colspan="2"> <div align="center" class="whitetext">图片修改</div></td>
          </tr>
          <tr> 
            <td width="16%"> <div align="right"><strong><font color="#FF0000">[1]-</font></strong>图片1.jpg：</div></td>
            <td width="84%"> <input name="photo" type="text" value="<%=rs("pic1")%>" size="50" maxlength="100"> 
              <a href ="Prod_B_upload.asp?gb=<%=gb%>&t=1" OnClick='return openAdminWindow(this.href);'> 
              <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
              <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=450,scrollbars=yes,resizable=yes');	return false;}</script>
              <font color="#003366">780*248象素 72分辩率 JPG格式(首页大动画1)</font></td>
          </tr>
          <tr> 
            <td><div align="right">图片1的链接：</div></td>
            <td><input name="link1" type="text" id="link1" size="70" value="<%=rs("link1")%>"></td>
          </tr>
          <tr bgcolor="f5f5f5"> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td><div align="right"><strong><font color="#FF0000">[2]-</font></strong>图片2.jpg：</div></td>
            <td><input name="photo2" type="text" value="<%=rs("pic2")%>" size="50" maxlength="100"> 
              <a href ="Prod_B_upload.asp?gb=<%=gb%>&t=2" OnClick='return openAdminWindow(this.href);'> 
              <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
              <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=450,scrollbars=yes,resizable=yes');	return false;}</script>
              <font color="#003366">780*248象素 72分辩率 JPG格式(首页大动画2)</font></td>
          </tr>
          <tr> 
            <td><div align="right">图片2的链接：</div></td>
            <td><input name="link2" type="text" id="link22" size="70" value ="<%=rs("link2")%>"></td>
          </tr>
          <tr bgcolor="f5f5f5"> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td><div align="right"><strong><font color="#FF0000">[3]-</font></strong>图片3.jpg：</div></td>
            <td><input name="photo3" type="text" value="<%=rs("pic3")%>" size="50" maxlength="100"> 
              <a href ="Prod_B_upload.asp?gb=<%=gb%>&t=3" OnClick='return openAdminWindow(this.href);'> 
              <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
              <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=450,scrollbars=yes,resizable=yes');	return false;}</script>
              <font color="#003366">780*248象素 72分辩率 JPG格式(首页大动画3)</font></td>
          </tr>
          <tr> 
            <td><div align="right">图片3的链接：</div></td>
            <td><input name="link3" type="text" id="link3" value="<%=rs("link3")%>" size="70"></td>
          </tr>
          <tr bgcolor="f5f5f5"> 
            <td height="10" colspan="2"></td>
          </tr>
		  <tr > 
            <td><div align="right"><strong><font color="#FF0000">[4]-</font></strong>图片4.jpg：</div></td>
            <td><input name="photo4" type="text" value="<%=rs("pic4")%>" size="50" maxlength="100"> 
              <a href ="Prod_B_upload.asp?gb=<%=gb%>&t=4" OnClick='return openAdminWindow(this.href);'> 
              <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
              <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=450,scrollbars=yes,resizable=yes');	return false;}</script>
              <font color="#003366">780*248象素 72分辩率 JPG格式(首页大动画4)</font></td>
          </tr>
          <tr > 
            <td><div align="right">图片4的链接：</div></td>
            <td><input name="link4" type="text" id="link4" value="<%=rs("link4")%>" size="70"></td>
          </tr>
          <tr bgcolor="f5f5f5"> 
            <td height="10" colspan="2"></td>
          </tr>
		  <tr > 
            <td><div align="right"><strong><font color="#FF0000">[5]-</font></strong>图片5.jpg：</div></td>
            <td><input name="photo5" type="text" id="photo5" value="<%=rs("pic5")%>" size="50" maxlength="100"> 
              <a href ="Prod_B_upload.asp?gb=<%=gb%>&t=5" OnClick='return openAdminWindow(this.href);'> 
              <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
              <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=450,scrollbars=yes,resizable=yes');	return false;}</script>
              <font color="#003366">780*248象素 72分辩率 JPG格式(首页大动画5)</font></td>
          </tr>
          <tr > 
            <td><div align="right">图片5的链接：</div></td>
            <td><input name="link5" type="text" id="link5" value="<%=rs("link5")%>" size="70"></td>
          </tr>
          <tr bgcolor="f5f5f5"> 
            <td height="10" colspan="2"></td>
          </tr>

		  <tr > 
            <td colspan="2"><div align="center"> 
                <input  type="hidden" name="gb" value=<%=gb%>>
				<input type="submit" name="Submit" value="确认提交">
              </div></td>
          </tr>
          <tr> 
            <td colspan="2"><table width="98%" border="0" align="center" cellpadding="4" cellspacing="0">
                <tr> 
                  <td width="79%"><strong>说明：</strong></td>
                </tr>
                <tr> 
                  <td> 1.每张图片为RGB的色彩模式 JPG的图片格式。</td>
                </tr>
                <tr> 
                  <td>2.图片的名称不能以中文和特殊字符命名，只能以英文或数字命名。</td>
                </tr>
                <tr> 
                  <td valign="top"><font color="#FF6600"> 3.如：需要更换第三张图片，就在“图片三”的位置上传一张新图片，然后点击“确认提交” 
                    系统将自动更换原来的第三张图片。</font></td>
                </tr>
                <tr> 
                  <td valign="top">4.图片链接指的是&nbsp;在网站的首页上点击这个图片会打开的页面，可以是本网站的任意页面，只要将网站粘贴在输入框中就可以了</td>
                </tr>
              </table></td>
          </tr>
		</form>  
        </table>
	  <%
	  rs.movenext
	  loop
	  rs.close
	  set rs = nothing	 
	  %>
	  </td>
  </tr>
</table>
<br><br>
</body>
</html>


