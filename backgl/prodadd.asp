<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<%
dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
else
	if session("flag")<>"0" then
	
		cls = Instr(session("flag"), "prod")
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
<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true"bgcolor=#ffffff leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#222222" style="border-collapse: collapse">
  <tr> 
    <td height="28">当前位置: 网站管理中心--产品管理</td>
    <td align="right" ><b><a href="admin_prod.asp">产品</a> | <a href="admin_prodclass.asp">类别</a> | </b></td>
  </tr>
  <tr> 
    <td width="100%" colspan=2> <form name="prodtable" method="post" action="addsave.asp"  onSubmit="FrmAdd();">
        <TABLE width="100%" border=1 cellpadding=3 style="border-collapse: collapse" bordercolor="#cccccc" align=center>
          <tr> 
            <td align="right">排序号码</td>
            <%
						  dim sql,rs,cid
						  sql = "select top 1 xuhao from ProdMain where gb='"&gb&"' order by xuhao desc"
						  set rs = conn.execute(sql)
						  if rs.eof then
						  	cid = 1
						  else
						  	cid = cint(rs(0))+1
						  end if
						  rs.close
						  set rs = nothing
						  %>
            <td><input name="xuhao" type="text" value="<%=cid%>" size="4" maxlength="10" onKeyUp="value=value.replace(/\D+/g,'')">
              （<FONT COLOR=#FF0000>产品的排列顺序，只能输入数字</FONT>）</td>
          </tr>
          <tr style="display:none;"> 
            <td align="right">选择语言</td>
            <td> <select type=text name="gb" id="gb" >
                <%for i=0 to 9%>
                <%if len(gbbase(i))>0 then%>
                <option value='<%=i%>' <% if int(request("gb"))=i then %> selected <%end if%> ><%= gbbase(i) %></option>
                <%end if%>
                <%next%>
              </select> </td>
          </tr>
          <tr> 
            <td align="right">选择类别</td>
            <td> <select name="ParentID" id="ParentID" >
                <%
					dim first
					dim i
					set rs=server.CreateObject("adodb.recordset")
					sql="select * from pclass where parentid=0  and  gb='"&gb&"' order by id asc"
					rs.open sql,conn,3,1
					if not rs.eof then
					first=rs.GetRows()
				  for i=0 to ubound(first,2)
				%>
                <option value="<%=first(0,i)%>" <%if request("ParentID")=cstr(first(0,i)) Then Response.Write("selected")%>><%=first(1,i)%></option>
                <%
					response.Write(Downasp.getDownlist(first(0,i),""))
					next
					end if
					rs.close
					set rs=nothing
					%>
              </select> <font color="#FF0000">*必须选择最底层目录分类</font> </td>
          </tr>
          <TR> 
            <TD align="right">产品名称</TD>
            <TD><input name="prodid" type="text" id="prodid" value="<%=session("prodid")%>" size="50"> 
              <font color="#FF0000">*</font> <font color="#FF0000"><strong>必填</strong></font></TD>
          </TR>
          <tr bgcolor="f5f5f5"> 
            <td height="25" align="right">西班牙文名称</td>
            <td height="25" colspan="3" ><input name="prodid_sp" type="text" id="prodid_sp" value="" size="50">
              <font color="#FF0000">&nbsp; </font> <font color="#FF0000">如果不填，表示与英文版相同</font></td>
          </tr>
          <tr> 
            <td height="25" align="right">产品编号</td>
            <td height="25" colspan="3" ><input name="model" type="text" id="model" size="50">
              <font color="#FF0000">&nbsp; 填写产品编号</font></td>
          </tr>
          <%for i=1 to 1 %>
          <tr bgcolor="f5f5f5"> 
            <td height="25" align="right">西班牙文编号</td>
            <td height="25" colspan="3" ><input name="model_sp" type="text" id="model_sp" size="50">
              &nbsp; &nbsp;<font color="#FF0000">如果不填表示与英文版编号相同</font></td>
          </tr>
          <tr> 
            <td width="20%" height="25" align="right">产品缩略图</td>
            <td width="85%" height="25" colspan="3" ><input type="text" name="photo<%=i%>" maxlength="100" size="50"> 
              <a href ="Prod_X_upload.asp?txt=<%=i%>" OnClick='return openAdminWindow(this.href);'> 
              <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
              <script language=""javascript"">function openAdminWindow(url) { popupWin = window.open(url,'new_page','width=600,height=450,scrollbars=yes,resizable=yes');	return false;}</script> 
              &nbsp;&nbsp;产品图片为150*150像素 或 220*220像素&nbsp;&nbsp; <% if i=1 then%> <font color="#FF0000">*图片必填</font> <%end if%></td>
          </tr>
          <%next%>
          <TR style="display:none;"> 
            <TD align="right">简要概述</TD>
            <TD><textarea name="ImgFull" cols="75" rows="5" id="ImgFull"><%=session("ImgFull")%></textarea></TD>
          </TR>
          <TR bgcolor="f5f5f5"> 
            <TD colspan="2" align="right"><div align="center"><strong><font color="#FF6600">产品特点</font></strong></div></TD>
          </TR>
          <TR> 
            <TD colspan="2" align="center"><textarea name="content1" style="display:none"><%=session("content1")%></textarea> 
              <iframe id="eWebEditor1" src="../webedit/ewebeditor.asp?id=content1&style=standard650&skin=office2003" frameborder="0" scrolling="no" width="100%" height="400"></iframe> 
            </TD>
          </TR>
          <TR> 
            <TD colspan="2" align="center" height="25"> <div align="left"><font color="#FF6600">1.在以上编辑框要换行的话请按 
                Shift+Enter 组合键</font></div></TD>
          </TR>
          <TR> 
            <TD colspan="2" align="center" height="25"> <div align="left"><font color="#FF6600">2.如果您的文字内容是从别的地方复制过来，最好先将其复制到windows的记事本中，这样可以清除别人网站上的格式。</font></div></TD>
          </TR>
          <TR> 
            <TD colspan="2" align="center" height="25"> <div align="left"><font color="#FF6600">3.将鼠标移到编辑框的工具按钮上时，会显示这个按钮的功能，如：第三行的第一个按钮就是添加图片到文本中的功能。</font></div></TD>
          </TR>
          <TR> 
            <TD colspan="2" align="center" height="30"><INPUT TYPE="hidden" name=add value=ok> 
              <input type="submit" name="action" value="加商品"> <input type="reset" name="Submit2" value="重设"></TD>
          </TR>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>