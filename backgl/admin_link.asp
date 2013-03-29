<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<%
'---------------------确定是否具有权利-------------------------------
dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
else
	if session("flag")<>"0" then
	
		cls = Instr(session("flag"), "link")
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
'---------------------确定是否具有权利-------------------------------
%>
<html>
<head>
<META content="" name=keywords>
<META content="" name=description>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>管理中心--友情链接管理</title>
<link type="text/css" href="css.css" rel="stylesheet">
</head>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true"bgcolor=#ffffff leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#222222" style="border-collapse: collapse">
  <tr> 
    <td width="50%" height="28">当前位置:<a href="index.asp">网站管理中心</a>--<strong><span class="blod">友情链接管理</span></strong></td>
  </tr>
  <tr> 
    <td colspan="2">
	<!--#include file="choosegb.asp"--> 
      <%
			if request("gb")<>"" then 
			gb=request("gb")
			end if 
			
			action=request("action")
			if action="" then
	%>
  <tr> 
    <td colspan="2">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="45">
<input type="button" name="action2" onClick="javascript:location.href='admin_link.asp?action=newadd&gb=<%=gb%>';" value="添加网址">&nbsp;&nbsp;
          </td>
        </tr>
      </table>
      <table border="1" style="border-collapse: collapse;border:solid 1px" bordercolor="#cccccc" cellpadding=3 width=100%>
        <form name=Newslist action=admin_link.asp method=post>
          <tr bgcolor="#003366"> 
            <td width="2%" height="30"  align="center" bgcolor="#e5ecfe">选</td>
            <td width="6%" height="27"  align="center" bgcolor="#e5ecfe">序号</td>
			<td width="20%" height="27"  align="center" bgcolor="#e5ecfe">名称(点击可修改)</td>
            <td width="50%" height="27" align="center" bgcolor="#e5ecfe"><div align="left">网址</div></td>
			<td width="17%" height="27"  align="center" bgcolor="#e5ecfe">添加日期</td>
            <td width="5%"  align="center" bgcolor="#e5ecfe">推荐</td>
          </tr>
          <%
				dim rs,msg_per_page
				dim sql
				msg_per_page = 20 '定义每页显示记录条数
				sql = "select * from link where gb='"&gb&"'"
				sql= sql+" order by xuhao desc"
				set rs=Server.CreateObject("ADODB.RecordSet")
				rs.cursorlocation = 3 '使用客户端游标，可以使效率提高
				rs.pagesize = msg_per_page '定义分页记录集每页显示记录数
				rs.open sql,conn,1,1 
				if err.number<>0 then '错误处理
				response.write "数据库操作失败：" & err.description
				err.clear
				else
				if not (rs.eof and rs.bof) then '检测记录集是否为空
				totalrec = RS.RecordCount 'totalrec：总记录条数
				if rs.recordcount mod msg_per_page = 0 then '计算总页数,recordcount:数据的总记录数
				n = rs.recordcount\msg_per_page 'n:总页数
				else 
				n = rs.recordcount\msg_per_page+1 
				end if 
				
				currentpage = request("page") 'currentpage:当前页
				If currentpage <> "" then
				currentpage = cint(currentpage)
				if currentpage < 1 then 
				currentpage = 1
				end if 
				if err.number <> 0 then 
				err.clear
				currentpage = 1
				end if
				else
				currentpage = 1
				End if 
				if currentpage*msg_per_page > totalrec and not((currentpage-1)*msg_per_page < totalrec)then 
				currentPage=1
				end if
				rs.absolutepage = currentpage 'absolutepage：设置指针指向某页开头
				rowcount = rs.pagesize 'pagesize：设置每一页的数据记录数
				do while not rs.eof and rowcount>0
%>
          <tr onmouseover="this.style.backgroundColor='#f5f5f5'" onmouseout="this.style.backgroundColor=''">	
            <td height="30" class="b"> <input type="checkbox" value="<%=Cstr(rs("NewsId"))%>" name="id"></td>
            <td><div align="center"><%=rs("xuhao")%></div></td>
			<td><div align="center"><a href="admin_link.asp?action=modify&id=<%=Cstr(rs("NewsId"))%>&page=<%=currentpage%>&newsclassL=<%=request("newsclass")%>"><%=rs("NewsTitle")%></a></div></td>
            <td><div align="left"><%=rs("Source")%> </div></td>
			<td><div align="center"><%=rs("date")%></div></td>
            <td><div align="center">
				<%if rs("online1")=true then %>
                <a href="admin_link.asp?action=close&type=1&newsid=<%=rs("newsid")%>&page=<%=request("page")%>"><font color="#FF0000">是</font></a> 
                <%else%>
                <a href="admin_link.asp?action=open&type=1&newsid=<%=rs("newsid")%>&page=<%=request("page")%>"><font color="#009999">否</font></a> 
                <%end if%>
</div></td>
          </tr>
          <%
					rowcount=rowcount-1
					rs.movenext   
					loop
					end if
					end if     
					rs.close
					conn.close
					set rs=nothing
					set coon=nothing
				%>
          <tr> 
            <td colspan=7 class=b> <input type='checkbox' name=chkall onclick='CheckAll(this.form)'>
              全选 
              <input name="action" type="submit" id="action" onclick="{if(confirm('该操作不可恢复！\n\n确定删除选定的网址？')){this.document.Prodlist.submit();return true;}return false;}" value="删除"> 
              <input type="hidden" name="gb" value="<%=gb%>"> 
              <input type="hidden" name="page" value="<%=request("page")%>"> 
            </td>
          </tr>
        </form>
      </table>
	  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="2">
        <tr> 
          <td height="20"><font color="#FF6600"><strong>[提示1]:</strong></font><font color="#FF6600"><strong>置顶</strong>和<strong>状态</strong>都是开关按钮，点击<strong>在线</strong>可以设置为<strong>离线</strong>，同样点击<strong>置顶</strong>即可取消置顶功能</font></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="4">
        <tr>
    <td align="center"><%call listPages()%></td>
  </tr>
</table>
<%end if

if action="删除" then
delid=replace(request("id")," ","")
call proddel()
end if

if action="newadd" then
call newsadd()
end if

if action="modify" then
id=replace(request("id")," ","")
call modify()
end if

'单个关闭产品
if action="close" then 
newsid=replace(request("newsid")," ","")
call prodclose()
end if

'单个打开产品
if action="open" then
newsid=replace(request("newsid")," ","")
call prodopen()
end if

%>
</table>
</body>
</html>

<%
'-----------------------分页显示代码开始----------------------------------
sub listPages() 
	if n <= 1 then exit sub 
	if currentpage = 1 then%>
		<font color=darkgray>首页</font>
	<%else%> 
		<font color="black">
		<a href="<%=request.ServerVariables("script_name")%>?page=1&parentid=<%=request("parentid")%>">首页</a>
		<a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage-1%>&parentid=<%=request("parentid")%>">前页</a>
		</font>
	<%end if
	 if currentpage = n then%> 
		&nbsp;&nbsp;<font color=darkgray >后页</font>
	<%else%> 
		<font color=black >
		&nbsp;&nbsp;<a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage+1%>&parentid=<%=request("parentid")%>">下页</a>
		&nbsp;&nbsp;<a href="<%=request.ServerVariables("script_name")%>?page=<%=n%>&parentid=<%=request("parentid")%>">末页</a>
		 </font>
	<%end if%>
	&nbsp;&nbsp;<font color="black">总:<%=currentpage%>/<%=n%>页&nbsp;&nbsp;总共:<%=totalrec%>个产品 [<%=msg_per_page%>产品/页]</font>
  <select name="menu1" onChange="MM_jumpMenu('self',this,0)">
  <option value="?page=<%=currentpage%>&parentid=<%=request("parentid")%>">第<%=currentpage%>页</option>
  <%for i=1 to n
  	if currentpage <> i then%>
    <option value="?page=<%=i%>&parentid=<%=request("parentid")%>">第<%=i%>页</option>
  <%end if
  	next%>
  </select>
<%end sub
'---------------------------分页显示代码结束---------------------------------
%>

<%
'--------------------删除网址开始----------------------------
sub proddel()
	if delid="" or isnull(delid) then
		Response.write "<BLOCKQUOTE><br><br>操作失败，没有选择合适参数！<A HREF='admin_link.asp'><b>点击返回</b></A><BR><br><meta http-equiv=refresh content=""2;URL=admin_link.asp""></BLOCKQUOTE>"
	else
		conn.execute("delete from link where NewsId in ("&delid&")")
		if err.number<>0 then '错误处理
			response.write "数据库操作失败：" & err.description
			err.clear
		else		
			conn.close
			set conn=nothing
			Response.write "<tr><td colspan=8><BLOCKQUOTE><br><br>网址删除成功！<A HREF='admin_link.asp'><b>点击返回</b></A><BR><br><meta http-equiv=refresh content=""2;URL=admin_link.asp?page="&request("page")&"""></BLOCKQUOTE></td></tr>"
		end if
	end if
end sub
'--------------------删除网址结束----------------------------


'-------------------增加网址开始-----------------------------
sub newsadd()
	if request("add") ="ok" then
	if request.form("NewsTitle")="" then
	response.Write "<script language=JavaScript>{window.alert('网站名称不得为空!');window.history.go(-1);}</script>"
	response.end
	end if
	if request.form("source")="" then
	response.Write "<script language=JavaScript>{window.alert('网址不得为空!');window.history.go(-1);}</script>"
	response.end
	end if

		Set rs=Server.CreateObject("ADODB.Recordset")
		sql = "select * from link"
		rs.open sql,conn,1,3
		if err.number<>0 then '错误处理
			response.write "数据库操作失败：" & err.description
			err.clear
		else		
		rs.Addnew
		rs("xuhao")=request.Form("xuhao")
		rs("NewsTitle")=request.form("NewsTitle")
		rs("Source")=request.Form("Source")
		rs("images")=request.Form("images")
		rs("notes")=request.Form("notes")
		rs("Online1")=request.form("Online1")
		rs("gb")=request.Form("gb")
		rs.update	

		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
		response.Write "<script language=JavaScript>{window.alert('提交成功！');window.location.href='admin_link.asp?gb="&request("gb")&"';}</script>"
		end if	
	else
%>
	<br>
	
<table width="100%" border="1" cellpadding="3" style="border-collapse: collapse" bordercolor="#cccccc" align="center">
  <form name="myform" method="post" action="admin_link.asp?action=newadd">
    <tr> 
      <td align="right">排列序号</td>
      <td>
		<%
		sql = "select top 1 xuhao from link where gb='"&gb&"' order by xuhao desc"
	    set rs = conn.execute(sql)
		if rs.eof then
			cid = 1
		else
			cid = cint(rs(0))+1
		end if
		rs.close
		set rs = nothing
	   %>
	  <input name="xuhao" size="4" type="text" id="xuhao" onKeyUp="value=value.replace(/\D+/g,'')" value="<%=cid%>"><font color="#FF0000">*</font>必须是数字,同级目录序号不能重复</td>
    </tr>
    <tr style="display:none;"> 
      <td align="right">语言选择</td>
      <td> 
		<select type=text name="gb" id="gb" >
               <%for i=0 to 9%>
               <%if len(gbbase(i))>0 then%>
               <option value='<%=i%>' <% if int(request("gb"))=i then %> selected <%end if%> ><%= gbbase(i) %></option>
               <%end if%>
               <%next%>
         </select>
		</td>
    </tr>
    <tr> 
      <td align="right">网站标题</td>
      <td><input name="NewsTitle" type="text" value="" size="40" maxlength="100"> (必填)</td>
    </tr>
    <tr> 
      <td align="right">网站地址</td>
      <td><input name="Source" type="text" size="80" maxlength="100" value="Http://"> (必填)</td>
    </tr>
		<tr style="display:none;"> 
            <td align="right">缩略图</td>
            <td><input type="text" name="images"  id="images" maxlength="100" size="50"> 
              <a href ="Prod_pic_upload.asp" OnClick='return openAdminWindow(this.href);'> 
              <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
              <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=400,scrollbars=yes,resizable=yes');	return false;}</script>   
            </td>
    </tr>
    <tr style="display:none;"> 
      <td align="right">网站说明</td>
      <td><textarea name="notes" cols="60" rows="5"></textarea></td>
    </tr>
    <tr> 
      <td align="right">是否推荐</td>
      <td>
	  <input type="radio" name="Online1" value=false checked>否 
      <input type="radio" name="Online1" value=true >是 
      </td>
    </tr>
    <tr> 
      <td align="right"> <input type="hidden" name="add" value="ok"> </td>
      <td><input type="Submit" name="Submit" value="提交"> <input type="reset" name="Submit2" value="重新填写"></td>
    </tr>
  </form>
</table>
<br><br>
<%		
end if
end sub
'-------------------增加网址结束-----------------------------
%>

<%
'-------------------修改网址开始-----------------------------
sub modify()

if id="" then
	response.write "非法网址编号"
	response.write "<meta http-equiv=refresh content=""1;URL=admin_link.asp"">"
else
	'修改网址资料
	if request("modify")="ok" then
	set rs=server.createobject("adodb.recordset")
	sql = "select * from link where NewsId="&id
	rs.open sql,conn,1,3
	if err.number<>0 then
		response.write "数据库操作失败：" & err.description '错误描述
		err.clear
	else
		if not (rs.eof and rs.bof) then
		rs("xuhao")=request.Form("xuhao")
		rs("NewsTitle")=request.form("NewsTitle")
		rs("Source")=request.Form("Source")
		rs("images")=request.Form("images")
		rs("notes")=request.Form("notes")
		rs("Online1")=request.form("Online1")
		rs.update
		end if
	end if	
	rs.close

	response.write "网址资料已经修改"
	response.write "<meta http-equiv=refresh content=""1;URL=admin_link.asp?page="&request("page")&""">"
	response.end
	end if

	'显示详细资料
	set rs = server.createobject("adodb.recordset")
	sql = "select * from link where NewsId="&id
	rs.open sql,conn,1,1
	if err.number<>0 then '错误处理
		response.write "数据库操作失败：" & err.description '错误描述
		err.clear
	else
		if not (rs.eof and rs.bof) then 
%>
	<br>
	
<table width="100%" border="1" cellpadding="3" style="border-collapse: collapse" bordercolor="#cccccc" align="center">
  <form name="myform" method="post" action='admin_link.asp?action=modify&id=<%=id%>'>
    <tr> 
      <td align="right">排列序号</td>
      <td><input name="xuhao" size="8" type="text" id="xuhao" onKeyUp="value=value.replace(/\D+/g,'')" value="<%=rs("xuhao")%>"><font color="#FF0000">*</font>必须是数字,同级目录序号不能重复</td>
    </tr>
    <tr> 
     <td align="right">网站标题</td>
      <td><input name="NewsTitle" type="text" value="<%=rs("NewsTitle")%>" size="40" maxlength="100">  (必填)</td>
    </tr>
    <tr> 
      <td align="right">网站地址</td>
      <td><input name="Source" type="text" value="<%=rs("Source")%>" size="80" maxlength="100"> (必填)</td>
    </tr>
		<tr style="display:none;"> 
            <td align="right">缩略图</td>
            <td><input type="text" name="images" maxlength="100" size="50" value="<%=rs("images")%>"> 
              <a href ="Prod_pic_upload.asp" OnClick='return openAdminWindow(this.href);'> 
              <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
              <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=400,scrollbars=yes,resizable=yes');	return false;}</script> 
            </td>
    </tr>
    <tr style="display:none;"> 
      <td align="right">网站说明</td>
      <td><textarea name="notes" cols="60" rows="5"><%=rs("notes")%></textarea></td>
    </tr>
    <tr> 
      <td align="right">是否推荐</td>
      <td>
	  <input type="radio" name="Online1" value=false <% if rs("online1")=false then %>checked<%end if%>>否 
      <input type="radio" name="Online1" value=true  <% if rs("online1")=true then %>checked<%end if%>>是 
      </td>
    </tr>
	<tr> 
      <td align="right"><input type="hidden" name="modify" value="ok">
        <input type="hidden" name="page" value="<%=request("page")%>"> 
        <input type="hidden" name="newsclassL" value="<%=request("newsclassL")%>"> 
      </td>
      <td><input type="Submit" name="Submit" value="提交修改"> <input type="reset" name="Submit2" value="重新修改"> 
      </td>
    </tr>
  </form>
</table>
<br><br>
<%
else
	response.write "<br><blockquote>无此序号商品！<br><br><font color=ff0000><b>请检查！</b></font><br></blockquote>"
	end if	
	end if
rs.close
set rs=nothing
end if
end sub
'-------------------修改网址结束-----------------------------
%>


<%
'-------------------------------单个关闭产品开始----------------------------------
sub prodclose()
if newsid="" or isnull(newsid) then
		Response.write "<blockquote><br><br>操作失败，没有选择合适参数！<a href='admin_link.asp'><b>点击返回</b></a><br><br><meta http-equiv=refresh content=""2;URL=admin_link.asp""></blockquote>"
	else
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select * from link where newsid in ("&newsid&")"
		rs.open sql,conn,1,3
		if err.number<>0 then '错误处理
			response.write "数据库操作失败：" & err.description
			err.clear
		else
			if rs.eof and rs.bof then
				response.write "<script language='javascript'>"
				response.write "alert('该网址不存在，或者被删除了！');"		
				response.write "</script>"			
			else
				Do while not rs.eof
				rs("online1")=false
				rs.update
				rs.movenext
				loop
			end if  
			
			response.write "<script language='javascript'>"
			response.write "alert('设置成功！');"
			response.write "location.href='admin_link.asp?page="&request("page")&"';"			
			response.write "</script>"
			rs.close
			conn.close
			set rs=nothing
			set conn=nothing
		end if
	end if
end sub
'-------------------------------单个关闭产品结束----------------------------------
%>


<%
'-------------------------------单个打开产品开始----------------------------------
sub prodopen()
	if newsid="" or isnull(newsid) then
		Response.write "<blockquote><br><br>操作失败，没有选择合适参数！<a href='admin_link.asp'><b>点击返回</b></a><br><br><meta http-equiv=refresh content=""2;URL=admin_link.asp""></blockquote>"
	else
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select * from link where newsid in ("&newsid&")"
		rs.open sql,conn,1,3
		if err.number<>0 then '错误处理
			response.write "数据库操作失败：" & err.description
			err.clear
		else
			if rs.eof and rs.bof then
				response.write "<script language='javascript'>"
				response.write "alert('该网址不存在，或者被删除了！');"		
				response.write "</script>"			
			else
				Do while not rs.eof
				rs("online1")=true
				rs.update
				rs.movenext
				loop			
			end if  
			response.write "<script language='javascript'>"
			response.write "alert('设置成功！');"
			response.write "location.href='admin_link.asp?page="&request("page")&"';"			
			response.write "</script>"
			rs.close
			conn.close
			set rs=nothing
			set conn=nothing
		end if
	end if
end sub
'-------------------------------单个打开产品结束----------------------------------
%>
