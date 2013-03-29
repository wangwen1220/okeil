<!--#include file="conn.asp"-->
<%
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
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>管理中心——&gt;栏目管理</title>
<link type="text/css" href="css.css" rel="stylesheet">
</head>
<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true"bgcolor=#ffffff leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" height="150" align="center">
  <tr>
	<td width="98%" height="400" valign="top"> 
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#222222" style="border-collapse: collapse">
        <tr> 
          <td width="85%" height="28">当前位置: 网站管理中心--网站栏目管理</td>
          <td width="15%" align="right"><b> </b></td>
        </tr>
        <tr> 
          <td colspan="2"> 
            <!----------------------无限级分类----------------------->
            <%

if request("gb")<>"" then 
gb=request("gb")
end if 

Select Case request("action")
	Case "AddClass"
		Call AddClass() 
	Case "EditClass"
		Call EditClass()
	Case "DelClass"
		Call DelClass()
	Case "SaveClass"
		Call SaveClass()
	Case Else
		Call wclassList() 
End Select
Call CloseConn()
%> <%
'--------------------------------栏目列表sub--------------------
Sub wclassList()
%> <table width="99%" border="0" cellpadding="2" cellspacing="1" class="tableline">
			<tr bgcolor="#FFFFFF"> 
			  <td height="35" colspan="3"><!--#include file="choosegb.asp"--></td>
			</tr>
              <tr bgcolor="efefef" onMouseOver="this.style.backgroundColor='#C1DEFB'" onMouseOut="this.style.backgroundColor=''"> 
                <td align=center class="underline">编号</td>
                <td height="30" class="underline">栏目列表</td>
                <td width="30%" class="underline"><% if session("admin")="admin" then %>
                  <a href="admin_wzclass.asp?action=AddClass&ParentID=0&gb=<%=gb%>">添加大栏目</a> 
                  <%else%>&nbsp;<%end if%></td>
              </tr>
              <%
'-------------------------------主要列表代码开始-----------------------------------------------
dim first
dim i
set rs=server.CreateObject("adodb.recordset")
sql="select * from wclass where parentid=0 and gb='"&gb&"' order by orders asc"
rs.open sql,conn,3,1
if rs.eof then
	response.Write("暂时没有分类")
	response.End
end if
first=rs.GetRows()
'GetRows的作用是将数据集输出到一数组中
rs.close
set rs=nothing

if(ubound(first,1)<0) then
	response.Write("暂时没有分类")
end if

for i=0 to ubound(first,2)
  if session("admin")="admin" then 
  response.Write("<tr bgcolor=f6f6f6 onmouseover=""this.style.backgroundColor='#C1DEFB'"" onmouseout=""this.style.backgroundColor=''""><td align=center width=""5%"" height=""25"" class=""underline""><b><font color=red>"&first(4,i)&"</font></b></td><td height=""25"" class=""underline""><img src='images/e.png' align='absmiddle'><b>"&first(1,i)&"</b>&nbsp;&nbsp;("&first(0,i)&")</td><td height=""25"" class=""underline""><a href=""admin_wzclass.asp?action=AddClass&ParentID="&first(0,i)&"&gb="&gb&""">添加</a> <a href=""admin_wzclass.asp?action=EditClass&ID="&first(0,i)&"&ParentID="&first(2,i)&"&gb="&gb&""">编辑内容</a> <a href=""admin_wzclass.asp?action=DelClass&ParentID="&first(0,i)&""" onclick=""{if(confirm('删除将包括该版面的所有资源，确定删除吗?')){return true;}return false;}"">删除栏目</a></td></tr><tr><td colspan=3><table name=div id=lar"&first(0,i)&" width=""100%"" border=""0"" cellpadding=0 cellspacing=0>")
  else
  response.Write("<tr bgcolor=f6f6f6 onmouseover=""this.style.backgroundColor='#C1DEFB'"" onmouseout=""this.style.backgroundColor=''""><td align=center width=""5%"" height=""25"" class=""underline""><b><font color=red>"&first(4,i)&"</font></b></td><td height=""25"" class=""underline""><img src='images/e.png' align='absmiddle'><b>"&first(1,i)&"</b></td><td height=""25"" class=""underline""><a href=""admin_wzclass.asp?action=EditClass&ID="&first(0,i)&"&ParentID="&first(2,i)&"&gb="&gb&""">编辑内容</a></td></tr><tr><td colspan=3><table name=div id=lar"&first(0,i)&" width=""100%"" border=""0"" cellpadding=0 cellspacing=0>")
  end if
  response.Write (Downasp.getcatalogs2(first(0,i),"")&"</table></td></tr>")
  
next
'------------------------------主要列表代码结束----------------------------------------
%>
</table>
<%
End Sub
'------------------------------栏目列表sub结束----------------------------------------
%> <%
'----------------------------添加栏目sub开始-----------------------------------------
Sub AddClass()
%> <table width="99%" border="0" cellpadding="0" cellspacing="5" class="tableline">
              <tr> 
                <td colspan="2"> 
                  <form name="myform" action="admin_wzclass.asp?action=SaveClass" method="post">
                    <table width="100%"  border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#cccccc" style="border-collapse: collapse;border:solid 1px">
                      <tr> 
                        <td height="30">栏目序号： 
                          <input name="Orders" size="2" type="text" id="Order" onKeyUp="value=value.replace(/\D+/g,'')"> 
                          <font color="#FF0000">*</font>必须是数字,同级目录序号不能重复</td>
                      </tr>
                      <tr style="display:none;"> 
                        <td height="30">语言种类： 
                          <select type=text name="gb" id="gb" >
                            <%for i=0 to 9%>
                            <%if len(gbbase(i))>0 then%>
                            <option value='<%=i%>' <% if int(request("gb"))=i then %> selected <%end if%> > 
                            <%= gbbase(i) %> </option>
                            <%end if%>
                            <%next%>
                          </select></td>
                      </tr>
                      <tr> 
                        <td height="30">栏目名称： 
                          <input name="ClassName" type="text" id="ClassName" size="40"></td>
                      </tr>
                      <tr style="display:none;"> 
                        <td height="30">设置热门： 
                          <input type="radio" name="ishot" value="true">
                          是 &nbsp;&nbsp; <input type="radio" name="ishot" value="false" checked>
                          否 </td>
                      </tr>
                      <tr style="display:none;"> 
                        <td height="30">所属栏目： 
                          <select name="ParentID" id="ParentID">
                            <%if Request("ParentID")<>0 Then%>
                            <%
							dim first
							dim i
							set rs=server.CreateObject("adodb.recordset")
							sql="select * from wclass where parentid=0  and gb='"&gb&"' order by id asc"							
							rs.open sql,conn,3,1
							first=rs.GetRows()
							rs.close
							set rs=nothing
							
							if(ubound(first,1)<0) then
							%>
                            <option value="0" selected>添加大类</option>
                            <%
							end if
							for i=0 to ubound(first,2)
							%>
                            <option value="<%=first(0,i)%>" <%if request("ParentID")=cstr(first(0,i)) Then Response.Write("selected")%>><%=first(1,i)%></option>
                            <%
							response.Write(Downasp.getDownlist(first(0,i),""))
							next
							%>
                            <%Else%>
                            <option value="0" selected>添加大类</option>
                            <%End IF%>
                          </select></td>
                      </tr>
                      <tr style="display:none;"> 
                        <td height="30" bgcolor="f5f5f5"> <div align="left">栏目图片： 
                            <input name="pic" type="text" id="pic" size="50" maxlength="100">
                            <a href ="Prod_D_upload.asp" OnClick='return openAdminWindow(this.href);'> 
                            <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
                            <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=400,scrollbars=yes,resizable=yes');	return false;}</script>
                          </div></td>
                      </tr>
                      <tr>
                        <td align="center" bgcolor="f5f5f5"><div align="left">我们提供了针对每个栏目的SEO优化设置，可以在以下的文本框内输入此栏目的优化词组或语句。</div></td>
                      </tr>
                      <tr> 
                        <td align="center" bgcolor="f5f5f5"><div align="left">栏目SEO优化标题栏： 
                            <input name="kw1" type="text" id="kw1" size="80">
                          </div></td>
                      </tr>
                      <tr> 
                        <td align="center" bgcolor="f5f5f5"><div align="left"> 
                            栏目SEO优化关键字： 
                            <input name="kw2" type="text" id="kw2" size="80">
                          </div></td>
                      </tr>
                      <tr> 
                        <td align="center" bgcolor="f5f5f5"><div align="left"> 
                            栏目SEO优化描述栏： 
                            <input name="kw3" type="text" id="kw3" size="80">
                          </div></td>
                      </tr>
                      <tr> 
                        <td align="center" bgcolor="f5f5f5"><textarea name="details" style="display:none"></textarea> 
                          <iframe id="eWebEditor1" src="../webedit/ewebeditor.asp?id=details&style=standard650&skin=office2003" frameborder="0" scrolling="no" width="100%" height="400"></iframe></td>
                      </tr>
                      <tr> 
                        <td height="40" align="center"> <div align="left"> &nbsp;&nbsp; 
                            <input type="submit" name="Submit" value=" 添 加 ">
                            <a href="javascript:history.go(-2);">返回上一页</a> &nbsp; 
                            &nbsp; &nbsp; </div></td>
                      </tr>
                      <tr> 
                        <td><font color="#FF6600">特别说明：请在以上的编辑框中填写详细内容 <br>
                          1.在以下编辑框要换行的话请按<strong>Shift+Enter</strong> 组合键<br>
                          2.如果您的文字内容是从别的地方复制过来，最好先将其复制到windows的记事本中，这样可以清除别人网站上的格式。<br>
                          3.将鼠标移到编辑框的工具按钮上时，会显示这个按钮的功能，如：第三行的第一个按钮就是添加图片到文本中的功能。</font></td>
                      </tr>
                    </table>
                  </form></td>
              </tr>
            </table>
            <%
End Sub
'-----------------------------添加栏目sub结束----------------------------------------
%> <%
'--------------------------编辑栏目sub代码开始---------------------------------------
Sub EditClass()
If Request("editid")<>"" then
	If Downasp.IsValidStr(Request("ClassName")) = False Then
		Errmsg =  "栏目名称中含有非法字符。"
		Downasp.ShowMsg(Errmsg)
		Downasp.Back(-1)
		Response.End()
	End If
	'修改栏目名称
	conn.execute("update wclass set classname='"&Trim(Request("ClassName"))&"',parentid="&Request("ParentID")&",orders="&Request("Orders")&",ishot="&Request("ishot")&",kw1='"&trim(Request("kw1"))&"',kw2='"&trim(Request("kw2"))&"' ,kw3='"&trim(Request("kw3"))&"',details='"&trim(replace(Request("details"),"'",chr(34)))&"'  where id="&Request("editid")&"")
			
	'response.Redirect("admin_wzclass.asp?gb="&request("gb")&"")
	Downasp.ShowMsg("修改成功。")
	Downasp.GotoUrl("admin_wzclass.asp?gb="&request("gb")&"")
	Response.End()
End If

set rs = conn.execute("select * from wclass where id="&Request("ID")&"")
if not rs.eof then
	classname = rs(1)
	classid = rs(0)
	parentid = rs(2)
	orders =rs(4)
	ishot=rs(6)
	kw1=rs(7)
	kw2=rs(8)
	kw3=rs(9)
	kw4=rs(11)
end if
rs.close
set rs = nothing
%> <table width="99%" border="0" cellpadding="0" cellspacing="5" class="tableline">
              <tr> 
                <td colspan="2">
<form name="myform" action="admin_wzclass.asp?action=EditClass" method="post">
                    <table width="100%"  border="1" align="center" cellpadding="4" cellspacing="0" bordercolor="#cccccc" style="border-collapse: collapse;border:solid 1px">
                      <tr> 
                        <td height="30">栏目序号： 
                          <input name="Orders" size="2" type="text" id="Order" value="<%=orders%>" onKeyUp="value=value.replace(/\D+/g,'')"> 
                          <font color="#FF0000">*</font>必须是数字,同级目录序号不能重复</td>
                      </tr>
                      <tr> 
                        <td height="30">栏目名称： 
                          <input name="ClassName" type="text" id="ClassName" value="<%=classname%>" size="40"></td>
                      </tr>
                      <tr style="display:none;"> 
                        <td height="30">设置热门： 
                          <input type="radio" name="ishot" value="true" <% if ishot=true then %> checked <%end if%>>
                          是 &nbsp;&nbsp; <input type="radio" name="ishot" value="false" <% if ishot=false then %> checked <%end if%>>
                          否 </td>
                      </tr>
                      <tr <% if session("admin")<>"admin" then %> style="display:none;" <% end if %>> 
                        <td height="30">所属栏目： 
                          <select name="ParentID" id="select2">
                            <%
				dim first
				dim i
				set rs=server.CreateObject("adodb.recordset")
				if session("gb")<>"" then 
				sql="select * from wclass where parentid=0 and gb='"&request("gb")&"' order by id asc"
				else
				sql="select * from wclass where parentid=0 order by id asc"				
				end if
				rs.open sql,conn,3,1
				first=rs.GetRows()
				rs.close
				set rs=nothing
				
				if(ubound(first,1)<0) then
				%>
                            <option value="0" selected>add object</option>
                            <%
				end if
				%>
                            <option value="0" selected>根目录</option>
                            <%
				for i=0 to ubound(first,2)
				%>
                            <option value="<%=first(0,i)%>" <%if request("ParentID")=cstr(first(0,i)) Then Response.Write("selected")%>><%=first(1,i)%></option>
                            <%
					response.Write(Downasp.getDownlist(first(0,i),""))
				next
				%>
                          </select></td>
                      </tr>
                      <tr style="display:none;"> 
                        <td height="30" bgcolor="f5f5f5"> <div align="left">栏目图片： 
                            <input name="pic" type="text" id="pic" size="50" maxlength="100" value="<%=pic%>">
                            <a href ="Prod_D_upload.asp" OnClick='return openAdminWindow(this.href);'> 
                            <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
                            <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=400,scrollbars=yes,resizable=yes');	return false;}</script>
                          </div></td>
                      </tr>
                      <tr> 
                        <td align="center" bgcolor="f5f5f5"><div align="left">我们提供了针对每个栏目的SEO优化设置，可以在以下的文本框内输入此栏目的优化词组或语句。</div></td>
                      </tr>
                      <tr> 
                        <td align="center" bgcolor="f5f5f5"><div align="left">栏目SEO优化标题栏： 
                            <input name="kw1" type="text" id="kw1" size="80" value="<%=kw1%>">
                          </div></td>
                      </tr>
                      <tr> 
                        <td align="center" bgcolor="f5f5f5"><div align="left"> 
                            栏目SEO优化关键字： 
                            <input name="kw2" type="text" id="kw2" size="80" value="<%=kw2%>">
                          </div></td>
                      </tr>
                      <tr> 
                        <td align="center" bgcolor="f5f5f5"><div align="left"> 
                            栏目SEO优化描述栏： 
                            <input name="kw3" type="text" id="kw3" size="80" value="<%=kw3%>">
                          </div></td>
                      </tr>
                      <tr> 
                        <td align="center" bgcolor="f5f5f5"><textarea name="details" style="display:none"><%=kw4%></textarea> 
                          <iframe id="eWebEditor1" src="../webedit/ewebeditor.asp?id=details&style=standard650&skin=office2003" frameborder="0" scrolling="no" width="100%" height="400"></iframe></td>
                      </tr>
                      <tr> 
                        <td height="40" align="center"> <div align="left"> &nbsp;&nbsp; 
                            <input type="submit" name="Submit" value=" 修改 ">
                            <input name="editid" type="hidden" value="<%=classid%>">
                            <input name="gb" type="hidden" value="<%=request("gb")%>">
                            <a href="javascript:history.go(-2);">返回上一页</a> </div></td>
                      </tr>
                      <tr> 
                        <td><font color="#FF6600">特别说明：请在以上的编辑框中填写详细内容 <br>
                          1.在以下编辑框要换行的话请按<strong>Shift+Enter</strong> 组合键<br>
                          2.如果您的文字内容是从别的地方复制过来，最好先将其复制到windows的记事本中，这样可以清除别人网站上的格式。<br>
                          3.将鼠标移到编辑框的工具按钮上时，会显示这个按钮的功能，如：第三行的第一个按钮就是添加图片到文本中的功能。</font></td>
                      </tr>
                    </table>
                  </form></td>
              </tr>
            </table>
            <%
End Sub
'--------------------------------编辑栏目sub代码结束---------------------------------
%> <%
'---------------------------------保存修改和编辑栏目的sub开始---------------------------
Sub SaveClass()
	If Downasp.IsValidStr(Request("ClassName")) = False Then
		Errmsg =  "栏目名称中含有非法字符。"
		Downasp.ShowMsg(Errmsg)
		Downasp.Back(-1)
		Response.End()
	End If
	conn.execute("insert into wclass(classname,parentid,orders,gb,ishot,kw1,kw2,kw3,details) values('"+Trim(Request("ClassName"))+"','"+Request("ParentID")+"','"+Request("Orders")+"','"+Request("gb")+"',"+Request("ishot")+",'"+Trim(Request("kw1"))+"','"+Trim(Request("kw2"))+"','"+Trim(Request("kw3"))+"','"+Trim(replace(Request("details"),"'",chr(34)))+"')")
	Response.Redirect("admin_wzclass.asp?gb="&request("gb")&"")
End Sub
'------------------------------保存修改和编辑栏目的sub结束------------------------------
%> <%
'------------------------------------删除栏目sub开始------------------------------------
Sub DelClass()
	Dim classID,ID
	ID = Request("ParentID")
	classID = ID & "," & Downasp.selectclass(ID)
	classID = Replace(classID,",,",",")
	if right(classID,1) = "," Then classID = left(classID,len(classID)-1)
	if left(classID,1) = "," Then classID = right(classID,len(classID)-1)
	'查找是否有下级栏目
	'set rs = conn.execute("select * from wclass where ParentID="&Request("ParentID")&"")
	'If not rs.eof then 
		'Downasp.ShowMsg("请先删除下级栏目。")
		'Downasp.Back(-1)
		'Response.End()
	'End IF
	'rs.close
	'set rs = nothing
	'删除栏目
	conn.execute("Delete from wclass where id in ("&classID&")")
	'删除相关资源
	'待定，未完成
	'------------
	Response.Redirect("admin_wzclass.asp")
End Sub
'-----------------------------删除栏目sub结束--------------------------
%> </td>
        </tr>
      </table>
</td></tr>
</table>
<script>
function $(id) {
 if (document.getElementById != null)
 {
 if(document.getElementById(id)) return document.getElementById(id);
 else{
	 if(document.getElementsByName(id)) return document.getElementsByName(id)[0];}
 }
 else if (document.all != null) {
 return document.all[id];
 }
 else {
 return null;
 } 
}
function AutoShow(str)
{
	var child = $("lar"+str);
	var displaystyle='none';
	
	if(child.style.display=='none') displaystyle='';
	child.style.display=displaystyle;
	
	/*
	if($("lar"+child[1]).style.display=='none') displaystyle='';
	for(var i=1;i<child.length;i++)
	{
		$("tr_"+child[i]).style.display=displaystyle;
	}
	*/
}
</script>
</body>
</html>