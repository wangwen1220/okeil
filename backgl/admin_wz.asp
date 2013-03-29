<!--#include file="conn.asp"-->
<%
'---------------------确定是否具有权利-------------------------------
dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
else
	if session("flag")<>"0" then
	
		cls = Instr(session("flag"), "wz")
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
<title>管理中心--站点栏目类别管理</title>
<link href="css.css" rel="stylesheet" type="text/css">
</head>
<SCRIPT LANGUAGE="JavaScript">
	<!--
	function checkdel(delid,gb){	
	if(confirm('请选择页面左边相应的功能区为该栏目添加内容'))
	{location.href="#;"}
	}
	//-->
</SCRIPT>

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>

<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="4" cellspacing="4">
  <tr> 
    <td>当前位置:<a href="index.asp">网站管理中心</a>--其它栏目管理<b></b></td>
  </tr>
</table>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><!--#include file="choosegb.asp"--></td>
  </tr>
</table>
<% if session("admin")="admin" then %>
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <form name="form1" method="post" action="admin_wz.asp?action=chooseobname">  
    <tr> 
      <td><table width="100%" border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#cccccc" bgcolor="#FFFFFF" style="border-collapse: collapse;border:solid 1px">
          <tr bgcolor="#E2F3FE"> 
            <td colspan="8"> <div align="center" class="bigclass"><strong>网站栏目添加</strong></div></td>
          </tr>
          <tr> 
            <td width="6%"> 
              <div align="right">序号</div></td>
            <td width="5%"> 
              <input name="xuhao" type="text" id="xuhao2" size="4" maxlength="4"></td>
            <td width="8%">
<div align="right">栏目名称</div></td>
            <td width="16%"> 
              <input name="wzbig" type="text" id="wzbig2" size="15" maxlength="30"></td>
            <td width="12%">
<div align="right">是否在线编辑</div></td>
            <td width="17%"> 
              <input type="radio" name="istext" value="true">
              是 &nbsp;&nbsp; <input name="istext" type="radio" value="false" checked>
              否 </td>
            <td width="7%">
<div align="right">文件名</div></td>
            <td width="29%"> 
              <div align="left"> 
                <input name="oburl" type="text" id="oburl2" size="20" maxlength="40">
                .asp </div></td>
          </tr>
          <tr> 
            <td colspan="8"><div align="center"> 
                <input type="hidden" name="gb" value="<% if request("gb")<>"" then %><%=request("gb")%><%else%><%=gb%><%end if%>">
                <input type="submit" name="Submit" value="提交确认">
              </div></td>
          </tr>
        </table></td>
    </tr>
</form>	
</table>
<%end if%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<%
if request("action")="chooseobname" then
'===============删除原来的栏目====================
'conn.execute("delete from wzclass where gb='ch'")
'=================================================
	set rsadd=Server.CreateObject("ADODB.Recordset")
	sql="select * from wzclass "
	rsadd.open sql,conn,3,3
		application.lock
		rsadd.addnew
		rsadd("xuhao")=request("xuhao")
		rsadd("wzbig")=request("wzbig")
		rsadd("oburl")=trim(request("oburl"))
		rsadd("istext")=request("istext")
		rsadd("gb")=request("gb")
		rsadd.update
		application.unlock
	rsadd.close
	set rsadd=nothing
	response.write "<script language=JavaScript>{window.alert('网站栏目添加成功');window.location.href='admin_wz.asp?gb="&request("gb")&"';}</script>"
end if
%>

<%
action=request("action")
if action="newadd" then 
call newadd()
end if

if action="delLarclass" then 
call delLarclass()
end if

if action="" then 
call list()
end if

%>
</body>
</html>

<%
sub list()
set rs=server.createobject("adodb.recordset")
sql="select * from wzclass where gb='"&gb&"' order by xuhao asc"
rs.open sql,conn,1,3
if err.number<>0 then '错误处理
	response.write "数据库操作失败：" & err.description '错误描述
	err.clear
else
%>
<table width='98%' border="1" cellpadding='4' bordercolor='#cccccc' style='border-collapse: collapse;border:solid 1px' align="center" bgcolor="#FFFFFF">
  <%
	Dim a,b,c
	if not (rs.eof and rs.bof) then
	a=rs("wzbig")
		Do While Not rs.eof			
			if (a<>b) then
	%>
  <tr onmouseover="this.style.backgroundColor='#f5f5f5'" onmouseout="this.style.backgroundColor=''">
    <td width='6%' height="30"  class="b">
<div align="center"><%=rs("xuhao")%></div></td>
    <td width='21%'><%=rs("wzbig")%>&nbsp;-&nbsp;<font color="#FF0000"><%=gbbase(rs("gb"))%></font></td>
    <td width="73%"><% if session("admin")="admin" then %><a href="admin_wz.asp?action=delLarclass&gb=<%=rs("gb")%>&Reid=<%=rs("wzbig")%>">删除栏目</a>&nbsp;|&nbsp;<%end if%><% if rs("istext")=true then %><a href="?action=newadd&reid=<%=rs("id")%>&gb=<%=rs("gb")%>"><font color="#ff6600">[编辑栏目内容]</font></a><%else%><a href=javascript:checkdel('"&rs("wzbig")&"','"&rs("gb")&"')>编辑内容<font color="#FF9900">[功能模块]</font></a><%end if%>&nbsp;|&nbsp;更新日期:<%=rs("wztime")%></td>
  </tr>
  <%
			else 		
			end if
			rs.movenext
			if rs.eof then
			exit do
			end if
		b=rs("wzbig")
		Loop
	end if
	%>
</table>
<%
end if
rs.close
set rs=nothing
end sub
%>

<%
sub newadd()
	if request("add") ="ok" then
		if request("wzcontain")="" then 
		response.Write "<script language=JavaScript>{window.alert('栏目内容不得为空!');window.history.go(-1);}</script>"
		response.end
		end if

	 	Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select * from wzclass where gb='"&request("gb")&"' and wzbig='"&request("wzbig")&"'"
		rs.open sql,conn,1,3
		if err.number<>0 then '错误处理
			response.write "数据库操作失败：" & err.description
			err.clear
		else
		rs("wzcontain")=request("wzcontain")
		rs("wztime")=now()
		rs.update	
		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
'		Response.Redirect"admin_prod.asp?action=detail&id="&request.form("ProdId")
		response.Write "<script language=Javascript>{window.alert('添加成功！');window.location.href='?gb="&request("gb")&"';}</script>"
		end if	

	else
%>
<table border="1"cellpadding="4" style="border-collapse: collapse" bordercolor="#cccccc" width="98%" id="AutoNumber1"  align="center" bgcolor="#FFFFFF">
  <form name="prodtable" onsubmit="return check()" method="post" action="admin_wz.asp">
    <tr style="display:none;"> 
      <td  height="30" align="right"> <div align="left">选择语言： 
		<select type=text name="gb" id="gb" >
               <%for i=0 to 9%>
               <%if len(gbbase(i))>0 then%>
               <option value='<%=i%>' <% if int(request("gb"))=i then %> selected <%end if%> ><%= gbbase(i) %></option>
               <%end if%>
               <%next%>
         </select>
        </div></td>
    </tr>
    <tr> 
      <td>所属栏目： 
        <select name="wzbig" id="wzbig">
          <%
			sqlp="select * from wzClass where gb='"&request("gb")& "' and id = "&request("reid")&""
			set rsp=server.createobject("ADODB.Recordset")
			rsp.open sqlp,conn,1,1
			while not rsp.eof
		  %>
          <option value="<%=rsp("wzbig")%>" selected><%=rsp("wzbig")%></option>
        </select> </td>
    </tr>
    <tr bgcolor="f5f5f5"> 
      <td height="25" align="center" bgcolor="f5f5f5"><div align="left">特别说明：请在以下的编辑框中填写详细内容 
          <font color="#FF6600"></font></div></td>
    </tr>
    <tr bgcolor="f5f5f5"> 
      <td height="25" align="center" bgcolor="#FFFFFF"> 
        <div align="left"><font color="#FF6600">1.在以下编辑框要换行的话请按 <b>Shift+Enter</b> 组合键</font></div></td>
    </tr>
    <tr bgcolor="f5f5f5">
      <td height="25" align="center" bgcolor="#FFFFFF"> 
        <div align="left"><font color="#FF6600">2.如果您的文字内容是从别的地方复制过来，最好先将其复制到windows的记事本中，这样可以清除别人网站上的格式。</font></div></td>
    </tr>
    <tr bgcolor="f5f5f5"> 
      <td height="25" align="center" bgcolor="#FFFFFF"> 
        <div align="left"><font color="#FF6600"> 3.将鼠标移到编辑框的工具按钮上时，会显示这个按钮的功能，如：第三行的第一个按钮就是添加图片到文本中的功能。</font></div></td>
    </tr>
    <TR> 
      <TD colspan="2" align="center"><textarea name="wzcontain" style="display:none"><%=rsp("wzcontain")%></textarea> 
        <iframe id="eWebEditor1" src="../webedit/ewebeditor.asp?id=wzcontain&style=standard650&skin=office2003" frameborder="0" scrolling="no" width="100%" height="400"></iframe> 
      </TD>
    </TR>
    <tr> 
      <td align="center" height="30"> <input type="submit" value="提 交" > &nbsp;&nbsp;&nbsp; 
        <input type="reset" value="清 除" class="form1"> 
		<input type="hidden" name="add" value="ok"> 
        <input type="hidden" name="action" value="newadd">
        [<a href="javascript:history.go(-1);">返回上一页</a>] </td>
    </tr>
    <%
	rsp.movenext
	wend
	rsp.close
	set rsp=nothing
	%>
  </form>
</table>
<%end if 
end sub%>


<%
sub delLarclass()
sql= "delete from wzclass where wzbig='"&request("Reid")&"'" 
conn.execute sql
conn.close
response.write "<br>&nbsp;&nbsp;删除栏目："&request("Reid")&"<br><br>"
response.write "<meta http-equiv=refresh content=""1;URL=admin_wz.asp?gb="&request("gb")&""">"
end sub
%>