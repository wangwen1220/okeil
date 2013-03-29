<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<%
'---------------------确定是否具有管理产品的权利-------------------------------
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
'---------------------确定是否具有管理产品的权利-------------------------------
%><html>
<style type="text/css">
<!--
.box2 {

	height: 20px;
	width: 200px;
}
-->
</style>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>管理中心——&gt;产品发布管理</title>
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

<SCRIPT LANGUAGE="JavaScript">
	<!--
	function checkdel(delid,gb,page){	
	if(confirm('删除选定的作品'))
	{location.href="admin_prod.asp?action=del&id="+delid+"&gb="+gb+"&page="+page+";"}
	}
	//-->
</SCRIPT>

<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true"bgcolor=#ffffff leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#222222" style="border-collapse: collapse">
  <tr> 
    <td height="28">当前位置: 网站管理中心--产品管理</td>
    <td align="right" ><b><a href="admin_prod.asp">产品</a> | <a href="admin_prodclass.asp">类别</a> 
      </b></td>
  </tr>
  <tr> 
    <td width="100%" colspan="2">
<%
action=request("action")
if action = "" then 
%>
      <table width="100%" border="1" cellpadding="3" bordercolor="cccccc" style="border-collapse: collapse;border:solid 1px">
        <form action="admin_prod.asp" name="search" method="get">
          <tr bgcolor="#EDF7FE"> 
            <td width="21%" bgcolor="#E5ECFE"> 
                <select name="type" id="type">
				 <option value="" selected>请选择类别</option>
                <%
					dim first
					dim j
					set rs=server.CreateObject("adodb.recordset")
					sql="select * from pclass where parentid=0  and  gb='"&gb&"' order by id asc"
					rs.open sql,conn,3,1
					if not rs.eof then
					first=rs.GetRows()
					for j=0 to ubound(first,2)
					%>
                <option value="<%=first(0,j)%>" <%if request("ParentID")=cstr(first(0,j)) Then Response.Write("selected")%>><%=first(1,j)%></option>
                <%
					response.Write(Downasp.getDownlist(first(0,j),""))
					next
					end if
					rs.close
					set rs=nothing
					%>
              </select>
              </td>
            <td width="24%" bgcolor="#E5ECFE">关键字 
              <input name="keyword" type="text" class="box2" id="keyword" size="30" maxlength="20">
            </td>
            <td width="20%" bgcolor="#E5ECFE"> 
              <input type="radio" name="online" value="false">
              离线 
              <input type="radio" name="online" value="true">
              在线 
              <input type="radio" name="online" value="" checked>
              所有 
              <input type="hidden" name="search" value="ok"> <input name="gb" type="hidden" id="gb" value="<%=gb%>"> 
            </td>
            <td width="35%" bgcolor="#E5ECFE"> 
              <input name="submit" type="submit" value="搜 索"></td>
          </tr>
        </form>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="45"> 
      <input type="button" name="action2" onClick="javascript:location.href='prodadd.asp?gb=<%=gb%>';" value="加产品"> 
	  <input type="button" name="action2" onClick="javascript:location.href='admin_prod.asp?remark=1&gb=<%=gb%>';" value="查看推荐产品">
	  <input type="button" name="action2" onClick="javascript:location.href='admin_prod.asp?remarka=1&gb=<%=gb%>';" value="查看最新产品">
      <input type="button" name="action2" onClick="javascript:location.href='admin_prod.asp?gb=<%=gb%>';" value="显示全部产品"> 
	  <input type="button" name="action2" onClick="javascript:location.href='admin_prodclass.asp?gb=<%=gb%>';" value="编辑产品类别">
          </td>
        </tr>
      </table>
<%
'----------------------------用于点击产品标题栏进行排序的 start-----------------
	if Request("Sortname")<>"" then
		session("sortname")=Request("Sortname")
	else
		session("sortname")=session("sortname")
	end if
	
	if Request("Sorttype")<>"" then
		if Request("Sorttype")="desc" then
		   session("sorttype")="asc"
		else
		   session("sorttype")="desc"
		end if
	else
		if session("sorttype")<>"" then
		   session("sorttype")=session("sorttype")
		else
		   session("sorttype")="desc"
	    end if
	end if 
'------------------------------------end -----------------------------------

'默认显示产品列
	dim rs,msg_per_page
	dim sql
	msg_per_page = 20 '定义每页显示记录条数
	set rs = server.createobject("adodb.recordset")
	sql = "select * from ProdMain where gb='"&gb&"' and ProdId is not null"
	
		'----------------接收产品搜索栏内的值----------------------------
		if request("type")<>"" then sql = sql + " and  gbseq ='"&trim(request("type"))&"' or larseq ='"&trim(request("type"))&"' or midseq ='"&trim(request("type"))&"' "
		if request("online")<>"" then sql = sql + " and online="&request("online")
		if request("keyword")<>"" then sql= sql+ " and ProdID like '%"&trim(request("keyword"))&"%' or Model like '%"&trim(request("keyword"))&"%'"
		if request("remark")="1" then sql = sql+" and remark='1' "
		if request("remarka")="1" then sql = sql+" and remarka='1' "
		if request("remarkb")="1" then sql = sql+" and remarkb='1' "
		
		if session("sortname")<>"" then
			sql= sql+" order by "&session("sortname")&" "&session("sorttype")
		else
			sql= sql+" order by xuhao desc" 
		end if

	rs.cursorlocation = 3 '使用客户端游标，可以使效率提高
	rs.pagesize = msg_per_page '定义分页记录集每页显示记录?
	rs.open sql,conn,1,1 
	
%>
      <table width="100%" border="1" align="center" cellpadding="3" bordercolor="#cccccc" style="border-collapse: collapse;border:solid 1px">
        <form name="Prodlist" action="admin_prod.asp" method="post">
          <tr> 
            <td width="5%" height="26" align="center" bgcolor="#E5ECFE"><a href="?Sortname=prodnum&Sorttype=<%=session("sorttype")%>&gb=<%=gb%>">ID</a></td>
            <td width="5%" height="26" align="center" bgcolor="#E5ECFE"><u><a href="?Sortname=xuhao&Sorttype=<%=session("sorttype")%>&gb=<%=gb%>">序号</a></u></td>
            <td width="10%" align="center" bgcolor="#E5ECFE" >缩略图</td>
            <td width="20%" height="26" align="center" bgcolor="#E5ECFE" ><div align="left"><a href="?Sortname=prodid&Sorttype=<%=session("sorttype")%>&gb=<%=gb%>"> 
                产品名称</a>(点击产品名称可编辑产品信息)</div></td>
            <td width="17%" height="26"  align="center" bgcolor="#E5ECFE" ><div align="left"><u><a href="?Sortname=larcode&Sorttype=<%=session("sorttype")%>&gb=<%=gb%>">类别</a></u></div></td>
            <td width="8%" height="26"  align="center" bgcolor="#E5ECFE" ><u><a href="?Sortname=adddate&Sorttype=<%=session("sorttype")%>&gb=<%=gb%>">上线日期</a></u></td>
            <td width="6%" align="center" bgcolor="#E5ECFE" >推荐</td>
			<td width="6%" align="center" bgcolor="#E5ECFE" >新品</td>
            <td width="7%" align="center" bgcolor="#E5ECFE" >浏览次数</td>
            <td width="5%" height="26" align="center" bgcolor="#E5ECFE" >状态</td>
            <td width="5%" height="26" align="center" bgcolor="#E5ECFE">删除</td>
          </tr>
          <%   
	if err.number<>0 then '错误处理
		response.write "<br><br><br>数据库操作失败：" & err.description
		err.clear
	else
		if not (rs.eof and rs.bof) then '检测记录集是否为空
			
			'----------------------------------------部分分页代码开始-----------------------------------------
			totalrec = rs.recordCount 'totalrec：总记录条数
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
	 	'----------------------------------------部分分页代码结束-----------------------------------------

			dim i
			dim k

	do while not rs.eof and rowcount>0
	%>
          <tr onMouseOver="this.style.backgroundColor='#f5f5f5'" onMouseOut="this.style.backgroundColor=''">	
            <%
	response.write "<td align=center><input type='checkbox' value='"&rs("ProdNum")&"' title='"&rs("ProdNum")&"' name=id></td>"
	response.write "<td align=center><FONT COLOR='#ff0000'>"&rs("xuhao")&"</font></td>"
	response.Write "<td align=center><img src='"&rs("photo1")&"' height='50'></td>"
	response.write "<td><a href=admin_prod.asp?action=prodadd&ParentID="&rs("gbSeq")&"&page="&currentpage&"&reid="&rs("ProdNum")&"&type="&request("type")&"&gb="&gb&" >"&rs("ProdId")&"</a></td>" 
	if rs("larcode")<> rs("midcode") then
	response.write "<td align=left>"&rs("larCode")&" - "&rs("midCode")&" - "&rs("gbCode")&"</td>"
	else
	response.write "<td align=left>"&rs("midCode")&" - "&rs("gbCode")&"</td>"
	end if
	response.write "<td align=center>"&rs("AddDate")&"</td>"
	if rs("remark")="1" then 	
	response.write "<td align=center><a href='admin_prod.asp?action=close&typeno=remark&id="&rs("ProdNum")&"&page="&request("page")&"&type="&request("type")&"&keyword="&trim(request("keyword"))&"&gb="&rs("gb")&"'><font color='#009999'>是</font></a></td>"
	else
	response.write "<td align=center><a href='admin_prod.asp?action=open&typeno=remark&id="&rs("ProdNum")&"&page="&request("page")&"&type="&request("type")&"&keyword="&trim(request("keyword"))&"&gb="&rs("gb")&"'>否</a></td>"
	end if

	if rs("remarka")="1" then 	
	response.write "<td align=center><a href='admin_prod.asp?action=close&typeno=remarka&id="&rs("ProdNum")&"&page="&request("page")&"&type="&request("type")&"&keyword="&trim(request("keyword"))&"&gb="&rs("gb")&"'><font color='#009999'>是</font></a></td>"
	else
	response.write "<td align=center><a href='admin_prod.asp?action=open&typeno=remarka&id="&rs("ProdNum")&"&page="&request("page")&"&type="&request("type")&"&keyword="&trim(request("keyword"))&"&gb="&rs("gb")&"'>否</a></td>"
	end if

	response.write "<td align=center>"&rs("ClickTimes")&"</td>"
	if rs("online")=true then 
		response.write "<td align=center><a href='admin_prod.asp?action=close&id="&rs("ProdNum")&"&page="&request("page")&"&type="&request("type")&"&keyword="&trim(request("keyword"))&"&gb="&rs("gb")&"'>在线</a></td>"
	else
		response.write  "<td align=center><a href='admin_prod.asp?action=open&id="&rs("ProdNum")&"&page="&request("page")&"&type="&request("type")&"&keyword="&trim(request("keyword"))&"&gb="&rs("gb")&"'><font color=ff0000>离线</font></a></td>"
	end if
	response.write "<td align=center><a href=javascript:checkdel('"&rs("ProdNum")&"','"&request("gb")&"','"&request("page")&"')>删除</a></td></tr>"
	rowcount=rowcount-1
	rs.movenext   
	loop
		else
		response.write "<tr><td colspan=10 align=center><br>无满足条件产品<br><br></td></tr>"
		end if
	end if     
	rs.close
	conn.close
	set rs=nothing
	set coon=nothing
	%>
          <tr> 
            <td colspan="13"> <input type="checkbox" name="chkall" onclick='CheckAll(this.form)'>
              全选 
              <input type="submit" name="action" value="删除" onClick="{if(confirm('该操作不可恢复！\n\n确定删除选定的产品？')){this.document.Prodlist.submit();return true;}return false;}"> 
              <input type="hidden" name="gb" value="<%=request("gb")%>"> 
            </td>
          </tr>
        </form>
      </table>
	  <!--#include file="pagelist.asp"-->
<%
end if

'删除产品
if action="del" then
delid=replace(request("id"),"'","")
call proddel()
end if

'单个关闭产品
if action="close" then 
ProdNum=replace(request("id")," ","")
call prodclose()
end if

'单个打开产品
if action="open" then
ProdNum=replace(request("id")," ","")
call prodopen()
end if

'修改资料
if action="prodadd" then
call prodadd()
end if

if action="删除" then
delid=replace(request("id")," ","")
call proddel()
end if
%>
</td></tr></table>
</body>
</html>

<%
'-------------------------------单个关闭产品开始----------------------------------
sub prodclose()
if ProdNum="" or isnull(ProdNum) then
		Response.write "<blockquote><br><br>操作失败，没有选择合适参数！<a href='admin_prod.asp'><b>点击返回</b></a><br><br><meta http-equiv=refresh content=""2;URL=admin_prod.asp""></blockquote>"
	else
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select * from ProdMain where ProdNum in ("&ProdNum&")"
		rs.open sql,conn,1,3
		if err.number<>0 then '错误处理
			response.write "数据库操作失败：" & err.description
			err.clear
		else
			if rs.eof and rs.bof then
				response.write "<script language='javascript'>"
				response.write "alert('该产品不存在，或者被删除了！');"		
				response.write "</script>"			
			else
				Do while not rs.eof
				if request("typeno")="remark" then 
				rs("remark")="0"
				elseif request("typeno")="remarka" then 
				rs("remarka")="0"
				elseif request("typeno")="remarkb" then 
				rs("remarkb")="0"
				else
				rs("online")=false
				end if
				rs.update
				rs.movenext
				loop
			end if  
			
			response.write "<script language='javascript'>"
			response.write "alert('设置成功！');"
			response.write "location.href='admin_prod.asp?page="&request("page")&"&gb="&request("gb")&"&type="&request("type")&"&keyword="&request("keyword")&"';"					
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
	if ProdNum="" or isnull(ProdNum) then
		Response.write "<blockquote><br><br>操作失败，没有选择合适参数！<a href='admin_prod.asp'><b>点击返回</b></a><br><br><meta http-equiv=refresh content=""2;URL=admin_prod.asp""></blockquote>"
	else
		Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select * from ProdMain where ProdNum in ("&ProdNum&")"
		rs.open sql,conn,1,3
		if err.number<>0 then '错误处理
			response.write "数据库操作失败：" & err.description
			err.clear
		else
			if rs.eof and rs.bof then
				response.write "<script language='javascript'>"
				response.write "alert('该产品不存在，或者被删除了！');"		
				response.write "</script>"			
			else
				Do while not rs.eof
				if request("typeno")="remark" then 
				rs("remark")="1"
				elseif request("typeno")="remarka" then 
				rs("remarka")="1"
				elseif request("typeno")="remarkb" then 
				rs("remarkb")="1"
				else
				rs("online")=true
				end if
				rs.update
				rs.movenext
				loop			
			end if  
			response.write "<script language='javascript'>"
			response.write "alert('产品 在线显示 设置成功！');"
			response.write "location.href='admin_prod.asp?page="&request("page")&"&gb="&request("gb")&"&type="&request("type")&"&keyword="&request("keyword")&"';"			
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
	
<%function invert(str) 
    invert=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;") 
    invert=replace(replace(replace(replace(invert,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=red>"),"[blue]","<font color=blue>") 
    invert=replace(replace(replace(replace(invert,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/blue]","</font>") 
	end function
%>

<%
'-------------------------------修改产品资料开始----------------------------------
sub prodadd()
'--------------------这些参数是为了修改成功后返回到指定的分类页面-------------------------
 page=request("page")
 reid=request("reid")
 gb=request("gb")
 Ptype=request("type")
 keyword=request("keyword")
 
	if request("add") ="ok" then
		if request("xuhao")="" then
		response.write "<script language=JavaScript>{window.alert('排序号码不得为空!');window.history.go(-1);}</script>"
		response.end
		end if
	
		set rs = conn.execute("select * from pclass where ParentID="&Request("ParentID_new")&"")
		If not rs.eof then
			response.Write "<script language=JavaScript>{window.alert('所属类别必须选择最底层目录分类!');window.history.go(-1);}</script>"
			response.end
		End If
		rs.close
	
		'查找该分类ID的类别名称
		set rs = conn.execute("select * from pclass where id="&Request("ParentID_new")&"")
		classname = rs("classname")
		classname_sp = rs("classname_sp")
		rs.close
	
	 	Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select * from ProdMain where ProdNum="&reid&""
		rs.open sql,conn,1,3
		if err.number<>0 then '错误处理
			response.write "数据库操作失败：" & err.description
			err.clear
		else
		rs("gbCode")=classname
		rs("gbCode_sp")=classname_sp
		rs("gbSeq")=request("ParentID_new")
		
'--------------将此产品的大类和中类和小类也查找出来一并写入数据库只能用于三级分类开始-----------
set rsclass1=server.createobject("adodb.recordset")
sqlclass1="select * from PClass where id="&request("parentID_new")&" and gb='"&gb&"' "
rsclass1.open sqlclass1,conn,1,3
	set rsclass2=server.CreateObject("adodb.recordset")
	sqlclass2="select * from PClass where id="&rsclass1("parentID")&" and gb='"&gb&"' "
	rsclass2.open sqlclass2,conn,1,3
		if rsclass2.eof and rsclass2.bof then 
			midcode=rsclass1("classname")
			midcode_sp=rsclass1("classname_sp")
			midseq=rsclass1("id")
			larcode=rsclass1("classname")
			larcode_sp=rsclass1("classname_sp")
			larseq=rsclass1("id")
		else
			set rsclass3=server.CreateObject("adodb.recordset")
			sqlclass3="select * from PClass where id="&rsclass2("parentID")&" and gb='"&gb&"' "
			rsclass3.open sqlclass3,conn,1,3
				if rsclass3.eof and rsclass3.bof then 
					midcode=rsclass2("classname")
					midcode_sp=rsclass2("classname_sp")
					midseq=rsclass2("id")
					larcode=rsclass2("classname")
					larcode_sp=rsclass2("classname_sp")
					larseq=rsclass2("id")
				else
					midcode=rsclass2("classname")
					midcode_sp=rsclass2("classname_sp")
					midseq=rsclass2("id")
					larcode=rsclass3("classname")
					larcode_sp=rsclass3("classname_sp")
					larseq=rsclass3("id")
				end if 
			rsclass3.close
			set rsclass3 = nothing
		end if 
		rsclass2.close
		set rsclass2 = nothing
	rsclass1.close
	set rsclass1 = nothing
	
'-----------------查找大类和中类结束------------------------
		rs("ProdId")=Trim(request.form("ProdId"))
		rs("ProdId_sp")=Trim(request.form("ProdId_sp"))
		rs("ProdName")=Trim(request.form("ProdName"))
		rs("Model")=Trim(request.form("Model"))
		rs("Model_sp")=Trim(request.form("Model_sp"))
		rs("LarCode")=larcode	
		rs("LarCode_sp")=larcode_sp
		rs("MidCode")=midcode
		rs("MidCode_sp")=midcode_sp
		rs("larseq")=LarSeq
		rs("midseq")=MidSeq
		rs("MemoSpec")=request.form("content1")
		rs("xuhao")=request.form("xuhao")
		rs("gb")=request.Form("gb")
		rs("SearchType")=request.Form("SearchType")
		rs("ImgFull")=request.Form("ImgFull")
		rs("photo1")=trim(request.Form("photo1"))
		rs.update	

		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
'		Response.Redirect"admin_prod.asp?action=detail&id="&request.form("ProdId")
		response.write "<script language=JavaScript>{window.alert('修改成功！');window.location.href='admin_prod.asp?gb="&gb&"&page="&page&"&type="&Ptype&"&keyword="&keyword&"';}</script>"
		end if	

	else
	yuyan=request.form("yuyan")
    reid=request("reid")
	 	Set rs=Server.CreateObject("ADODB.Recordset")
		sql="select * from ProdMain where ProdNum="&reid&" order by ProdNum asc"
		rs.open sql,conn,1,3
%>
<table width="100%" border="1" cellpadding="3" style="border-collapse: collapse" bordercolor="#cccccc" align="center">
  <form name="prodtable" method="post" action="">
    <tr>
      <td align="right">产品图片</td>
      <td><div align="left"><img src="<%=rs("photo1")%>" height="80" border="0"></div></td>
    </tr>
    <tr> 
      <td align="right">排序号码</td>
      <td><input name="xuhao" type="text" value="<%=rs("xuhao")%>" size="3" maxlength="10" onKeyUp="value=value.replace(/\D+/g,'')">&nbsp;&nbsp;产品的排列顺序,只能输入数字.</td>
    </tr>
    <tr style="display:none;"> 
            <td align="right">选择语言</td>
            <td> <select type=text name="gb" id="gb" >
                <%for i=0 to 9%>
                <%if len(gbbase(i))>0 then%>
                <option value='<%=i%>' <% if int(rs("gb"))=i then %> selected <%end if%> ><%= gbbase(i) %></option>
                <%end if%>
                <%next%>
              </select> </td>
          </tr>
    <tr> 
      <td width="15%" align="right">所属类别</td>
      <td width="85%"> <select name="ParentID_new" id="ParentID_new">
          <%
			dim first
			dim i
			set rsC=server.CreateObject("adodb.recordset")
			sql="select * from pclass where parentid=0 and gb='"&rs("gb")&"' order by id asc"
			rsC.open sql,conn,3,1
			first=rsC.GetRows()
			rsC.close
			set rsC=nothing
			
			if(ubound(first,1)<0) then
			%>
          <option value="0" selected>暂时没有分类</option>
          <%
			end if
			
			for i=0 to ubound(first,2)
			%>
          <option value="<%=first(0,i)%>" <%if request("ParentID")=cstr(first(0,i)) Then Response.Write("selected")%>><%=first(1,i)%></option>
          <%
			  response.Write(Downasp.getDownlist(first(0,i),""))
			next
			%>
        </select>
        <font color="#FF0000">*必须选择最底层目录分类</font></td>
    </tr>
    <tr> 
      <td align="right">产品名称</td>
      <td><input name="prodid" type="text" value="<%=rs("prodid")%>" size="50" maxlength="100"></td>
    </tr>

    <tr bgcolor="f5f5f5"> 
      <td align="right">西班牙文名称</td>
      <td><input name="prodid_sp" type="text" value="<%=rs("prodid_sp")%>" size="50" maxlength="100"></td>
    </tr>
	
    <tr> 
      <td align="right">产品编号</td>
      <td><input name="model" type="text" value="<%=rs("model")%>" size="50" maxlength="100"></td>
    </tr>

    <tr bgcolor="f5f5f5"> 
      <td align="right">西班牙文编号</td>
      <td><input name="model_sp" type="text" value="<%=rs("model_sp")%>" size="50" maxlength="100"></td>
    </tr>
	
	   <% for i=1 to 1 %>
          <tr> 
            <td width="20%" height="25" align="right">产品缩略图</td>
            <td width="80%" height="25" colspan="3" ><input type="text" name="photo<%=i%>" maxlength="100" size="50" value="<%=rs("photo"&i&"")%>"> 
              <a href ="Prod_X_upload.asp?txt=<%=i%>" OnClick='return openAdminWindow(this.href);'> 
              <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
              <script language=""javascript"">function openAdminWindow(url) { popupWin = window.open(url,'new_page','width=600,height=450,scrollbars=yes,resizable=yes');	return false;}</script>
              &nbsp; &nbsp;产品图片为150*150像素 或 2200*220像素</td>
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
            <TD colspan="2" align="center"><textarea name="content1" style="display:none"><%=rs("MemoSpec")%></textarea> 
              <iframe id="eWebEditor1" src="../webedit/ewebeditor.asp?id=content1&style=standard650&skin=office2003" frameborder="0" scrolling="no" width="100%" height="400"></iframe> 
            </TD>
          </TR>
		  
		  
   <tr style="display:none;"> 
      <td align="right">是否设定为推荐产品</td>
      <td> <input type="radio" name="Remark" value="0" <%if rs("Remark")="0"then%>checked<%end if%>>否 
        <input type="radio" name="Remark" value="1" <%if rs("Remark")="1" then%>checked<%end if%>>是</td>
    </tr>
    <tr> 
      <td colspan="2" align="right"> <div align="center"> 
          <input name="page" type="hidden" value="<%=request("page")%>"> 
		  <input name="type" type="hidden" value="<%=request("type")%>">
		  <input name="keyword" type="hidden" value="<%=request("keyword")%>">
          <input type="hidden" name="add" value="ok">
          <input type="submit" name="action" value="修改">
          &nbsp;&nbsp; 
          <input type="reset" name="Submit2" value="重设">
          &nbsp;</div></td>
    </tr>
  </form>
</table>
<%		
end if
end sub
'-----------------------修改产品资料结束-----------------------
%>



<%
'-------------------------------删除产品资料开始----------------------------------
sub proddel()
	if delid="" or isnull(delid) then
		Response.write "<blockquote><br><br>操作失败，没有选择合适参数！<a href='admin_prod.asp?gb="&gb&"'><b>点击返回</b></a><br><br><meta http-equiv=refresh content=""2;URL=admin_prod.asp?gb="&gb&"""></blockquote>"
	else
		conn.execute("delete from ProdMain where ProdNum in ("&delid&")")
		if err.number<>0 then '错误处理
			response.write "数据库操作失败：" & err.description
			err.clear
		else		
			conn.close
			set conn=nothing
			Response.write "<blockquote><br><br>产品删除成功！正在返回...<br><br><meta http-equiv=refresh content=""2;URL=admin_prod.asp?gb="&gb&"&page="&request("page")&"""></blockquote>"
		end if
	end if
end sub
'-------------------------------删除产品资料结束----------------------------------
%>
