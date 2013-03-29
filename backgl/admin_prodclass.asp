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
<title>管理中心——&gt;类别管理</title>
<link type="text/css" href="css.css" rel="stylesheet">
</head>
<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true"bgcolor=#ffffff leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" height="150" align="center">
  <tr>
	<td width="98%" height="400" valign="top"> 
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#222222" style="border-collapse: collapse">
        <tr> 
          <td width="85%" height="28">当前位置: 网站管理中心--类别管理</td>
          <td width="15%" align="right"><b><a href="admin_prod.asp">产品</a> | <a href="admin_prodclass.asp">类别</a> 
            </b></td>
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
		Call pclassList() 
End Select
Call CloseConn()
%> <%
'--------------------------------类别列表sub--------------------
Sub pclassList()
%> <table width="99%" border="0" cellpadding="2" cellspacing="1" class="tableline">
              <tr bgcolor="efefef" onMouseOver="this.style.backgroundColor='#C1DEFB'" onMouseOut="this.style.backgroundColor=''"> 
                <td align=center class="underline">编号</td>
                <td height="30" class="underline">类别列表<font color="#FF0000">(括号中是西班牙文的名字)</font></td>
                <td width="30%" class="underline"><a href="admin_prodclass.asp?action=AddClass&ParentID=0&gb=<%=gb%>">添加大类</a></td>
              </tr>
              <%
'-------------------------------主要列表代码开始-----------------------------------------------
dim first
dim i
set rs=server.CreateObject("adodb.recordset")
sql="select * from pclass where parentid=0 and gb='"&gb&"' order by orders asc"
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
  
  response.Write("<tr bgcolor=f6f6f6 onmouseover=""this.style.backgroundColor='#C1DEFB'"" onmouseout=""this.style.backgroundColor=''""><td align=center width=""5%"" height=""25"" class=""underline""><b><font color=red>"&first(4,i)&"</font></b></td><td height=""25"" class=""underline""><img src='images/e.png' align='absmiddle'> <b>"&first(1,i)&"&nbsp;&nbsp;(<font color='#ff6600'>"&first(11,i)&"</font>)</b></td><td height=""25"" class=""underline""><a href=""admin_prodclass.asp?action=AddClass&ParentID="&first(0,i)&"&gb="&gb&""">添加</a> <a href=""admin_prodclass.asp?action=EditClass&ID="&first(0,i)&"&ParentID="&first(2,i)&"&gb="&gb&""">编辑</a> <a href=""admin_prodclass.asp?action=DelClass&ParentID="&first(0,i)&"&gb="&gb&""" onclick=""{if(confirm('删除将包括该版面的所有资源，确定删除吗?')){return true;}return false;}"">删除</a></td></tr><tr><td colspan=3><table name=div id=lar"&first(0,i)&" width=""100%"" border=""0"" cellpadding=0 cellspacing=0>")
  response.Write (Downasp.getcatalogs(first(0,i),"")&"</table></td></tr>")
  
next
'------------------------------主要列表代码结束----------------------------------------
%>
</table>
<%
End Sub
'------------------------------类别列表sub结束----------------------------------------
%> <%
'----------------------------添加类别sub开始-----------------------------------------
Sub AddClass()
%> <table width="99%" border="0" cellpadding="0" cellspacing="5" class="tableline">
              <tr> 
                <td colspan="2"> <font color="#FF6633"><b>添加类别</b></font> <form name="myform" action="admin_prodclass.asp?action=SaveClass" method="post">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td height="30">类别序号： 
                          <input name="Orders" size="8" type="text" id="Order" onKeyUp="value=value.replace(/\D+/g,'')"> 
                          <font color="#FF0000">*</font>必须是数字,同级目录序号不能重复</td>
                      </tr>
                      <tr style="display:none;"> 
                        <td height="30">语言种类： 
                          <select type=text name="gb" id="gb" >
                            <%for i=0 to 9%>
                            <%if len(gbbase(i))>0 then%>
                            <option value='<%=i%>' <% if int(request("gb"))=i then %> selected <%end if%> >
                            <%= gbbase(i) %>
                            </option>
                            <%end if%>
                            <%next%>
                          </select></td>
                      </tr>
                      <tr> 
                        <td height="30">英文版类别名称： 
                          <input name="ClassName" type="text" id="ClassName"></td>
                      </tr>
                      <tr> 
                        <td height="30">西班牙文名称： 
                          <input name="ClassName_sp" type="text" id="ClassName_sp">
                          <font color="#FF0000">如果不填表示与英文版相同</font></td>
                      </tr>

                      <tr style="display:none;"> 
                        <td height="30">设置热门： 
                          <input type="radio" name="ishot" value="true">
                          是 &nbsp;&nbsp; <input type="radio" name="ishot" value="false" checked>
                          否 </td>
                      </tr>
                      <tr style="display:none;"> 
                        <td height="30">所属类别： 
                          <select name="ParentID" id="ParentID">
                            <%if Request("ParentID")<>0 Then%>
                            <%
							dim first
							dim i
							set rs=server.CreateObject("adodb.recordset")
							sql="select * from pclass where parentid=0  and gb='"&gb&"' order by id asc"							
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
                        <td height="30" bgcolor="f5f5f5"> <div align="left">类别图片： 
                            <input name="pic" type="text" id="pic" size="50" maxlength="100">
                            <a href ="Prod_D_upload.asp" OnClick='return openAdminWindow(this.href);'> 
                            <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
                            <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=400,scrollbars=yes,resizable=yes');	return false;}</script>
                          </div></td>
                      </tr>
                      <tr style="display:none;">
                        <td align="center" bgcolor="f5f5f5">
						<table width="100%" border="0" cellspacing="2" cellpadding="4">
						  <tr> 
							<td width="10%"><div align="right">网站优化标题栏</div></td>
							<td width="90%"><input name="kw1" type="text" id="kw1" size="80"></td>
						  </tr>
						  <tr> 
							<td><div align="right">网站优化关键字栏</div></td>
							<td><input name="kw2" type="text" id="kw2" size="80"></td>
						  </tr>
						  <tr> 
							<td><div align="right">网站优化描述栏</div></td>
							<td><input name="kw3" type="text" id="kw3" size="80"></td>
						  </tr>
						  <tr> 
							<td><div align="right">类别热门关键词</div></td>
							<td><input name="kw4" type="text" id="kw4" size="80">
                                请用“|”隔开&nbsp; 如： 关键字1|关键字2|关键字3</td>
						  </tr>
						</table>
						</td>
                      </tr>
                      <tr> 
                        <td height="45" align="center"> <div align="left"> &nbsp;&nbsp; 
                            <input type="submit" name="Submit" value=" 添 加 ">
                            <a href="javascript:history.go(-1);">返回上一页</a> &nbsp; 
                            &nbsp; &nbsp; </div></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                      </tr>
                    </table>
                  </form></td>
              </tr>
            </table>
            <%
End Sub
'-----------------------------添加类别sub结束----------------------------------------
%> <%
'--------------------------编辑类别sub代码开始---------------------------------------
Sub EditClass()
If Request("editid")<>"" then

	if request("ClassName")="" then
		response.write "<script language=JavaScript>{window.alert('类别名称不得为空!');window.history.go(-1);}</script>"
		response.end
	end if

	if request("ClassName_sp")="" then
		response.write "<script language=JavaScript>{window.alert('西班牙文类别名称不得为空!');window.history.go(-1);}</script>"
		response.end
	end if


	If Downasp.IsValidStr(Request("ClassName")) = False Then
		Errmsg =  "类别名称中含有非法字符。"
		Downasp.ShowMsg(Errmsg)
		Downasp.Back(-1)
		Response.End()
	End If
	'修改类别名称
	conn.execute("update pclass set classname='"&Trim(Request("ClassName"))&"',parentid="&Request("ParentID")&",orders="&Request("Orders")&",ishot="&Request("ishot")&",kw1='"&trim(Request("kw1"))&"',kw2='"&trim(Request("kw2"))&"' ,kw3='"&trim(Request("kw3"))&"',kw4='"&trim(Request("kw4"))&"',ClassName_sp='"&trim(Request("ClassName_sp"))&"'  where id="&Request("editid")&"")
	'修改产品中记录的对应的类别名称＆编号
	'修改产品中记录的对应的类别名称＆编号 
	'-----------------20090728修改 by thoms---------------------
	set rsprod = server.createobject("adodb.recordset")
	sqlprod = "select * from ProdMain where gb='"&gb&"' and ProdId is not null"
	rsprod.open sqlprod,conn,1,3 
	  do while not rsprod.eof 
	  		if rsprod("gbseq")=Request("editid") then 
			 	rsprod("gbcode")=Trim(Request("ClassName"))
				rsprod("gbcode_sp")=Trim(Request("ClassName_sp"))
					'----------------中类-------------------
					'if  rsprod("midseq")<> request("ParentID") and request("parentid")<>"0" then 
					'	rsprod("midseq")=trim(request("ParentID"))
					'	rsprod("midcode")=conn.execute("select classname from pclass where id="&request("ParentID"))(0)
					'	rsprod("midcode_sp")=conn.execute("select classname_sp from pclass where id="&request("ParentID"))(0)
					'else
				 		'rsprod("midcode")=Trim(Request("ClassName"))
						'rsprod("midcode_sp")=Trim(Request("ClassName_sp"))
						'rsprod("midseq")=Request("editid")
					'end if
					
					'----------------大类--------------------
					'if  rsprod("larseq")<> request("ParentID") and request("parentid")<>"0" then 
					'	rsprod("larseq")=trim(request("ParentID"))
					'	rsprod("larcode")=conn.execute("select classname from pclass where id="&request("ParentID"))(0)
					'	rsprod("larcode_sp")=conn.execute("select classname_sp from pclass where id="&request("ParentID"))(0)
					'else
				 		'rsprod("larcode")=Trim(Request("ClassName"))
						'rsprod("larcode_sp")=Trim(Request("ClassName_sp"))
						'rsprod("larseq")=Request("editid")
					'end if
			 end if
	 
 	  		if rsprod("midseq")=Request("editid") then 
			 	rsprod("midcode")=Trim(Request("ClassName"))
				rsprod("midcode_sp")=Trim(Request("ClassName_sp"))
					'----------------大类--------------------
					'if  rsprod("larseq")<> request("ParentID") and request("parentid")<>"0" then 
					'	rsprod("larseq")=trim(request("ParentID"))
					'	rsprod("larcode")=conn.execute("select classname from pclass where id="&request("ParentID"))(0)
					'	rsprod("larcode_sp")=conn.execute("select classname_sp from pclass where id="&request("ParentID"))(0)
					'else
				 		'rsprod("larcode")=Trim(Request("ClassName"))
						'rsprod("larcode_sp")=Trim(Request("ClassName_sp"))
						'rsprod("larseq")=Request("editid")
					'end if
			 end if
			
			if rsprod("larseq")=Request("editid") then
				 rsprod("larcode")=Trim(Request("ClassName"))
				 rsprod("larcode_sp")=Trim(Request("ClassName_sp"))
			end if
			rsprod.update
			rsprod.movenext
			loop
		rsprod.close
		set rsprod = nothing
			
	'conn.execute("update ProdMain set gbcode='"&Trim(Request("ClassName"))&"'  where gbseq='"&Request("editid")&"'")
	'conn.execute("update ProdMain set midcode='"&Trim(Request("ClassName"))&"' where midseq='"&Request("editid")&"'")
	'conn.execute("update ProdMain set larcode='"&Trim(Request("ClassName"))&"' where larseq='"&Request("editid")&"'")
	response.Redirect("admin_prodclass.asp?gb="&request("gb")&"")
	'Downasp.ShowMsg("修改成功。")
	'Downasp.GotoUrl("admin_prodclass.asp?gb="&request("gb")&"")
	Response.End()
End If

set rs = conn.execute("select * from pclass where id="&Request("ID")&"")
if not rs.eof then
	classname = rs(1)
	classid = rs(0)
	parentid = rs(2)
	orders =rs(4)
	ishot=rs(6)
	kw1=rs(7)
	kw2=rs(8)
	kw3=rs(9)
	kw4=rs(10)
	classname_sp=rs(11)
end if
rs.close
set rs = nothing
%> <table width="99%" border="0" cellpadding="0" cellspacing="5" class="tableline">
              <tr> 
                <td colspan="2">编辑类别： 
                  <form name="myform" action="admin_prodclass.asp?action=EditClass" method="post">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td height="30">类别序号： 
                          <input name="Orders" size="8" type="text" id="Order" value="<%=orders%>" onKeyUp="value=value.replace(/\D+/g,'')"> 
                          <font color="#FF0000">*</font>必须是数字,同级目录序号不能重复</td>
                      </tr>
                      <tr> 
                        <td height="30">英文版类别名称： 
                          <input name="ClassName" type="text" id="ClassName" value="<%=classname%>"></td>
                      </tr>

                      <tr> 
                        <td height="30">西班牙文名称： 
                          <input name="ClassName_sp" type="text" id="ClassName_sp" value="<%=classname_sp%>"></td>
                      </tr>
                      <tr style="display:none;"> 
                        <td height="30">设置热门： 
                          <input type="radio" name="ishot" value="true" <% if ishot=true then %> checked <%end if%>>
                          是 &nbsp;&nbsp; <input type="radio" name="ishot" value="false" <% if ishot=false then %> checked <%end if%>>
                          否 </td>
                      </tr>
                      <tr style="display:none;"> 
                        <td height="30">所属类别： 
                          <select name="ParentID" id="select2">
                            <%
				dim first
				dim i
				set rs=server.CreateObject("adodb.recordset")
				if session("gb")<>"" then 
				sql="select * from pclass where parentid=0 and gb='"&request("gb")&"' order by id asc"
				else
				sql="select * from pclass where parentid=0 order by id asc"				
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
                        <td height="30" bgcolor="f5f5f5"> <div align="left">类别图片： 
                            <input name="pic" type="text" id="pic" size="50" maxlength="100" value="<%=pic%>">
                            <a href ="Prod_D_upload.asp" OnClick='return openAdminWindow(this.href);'> 
                            <img src="images/icon_paperclip.gif" width="23" height="22" align="absmiddle" border="0"></a> 
                            <script language=""javascript"">function openAdminWindow(url) {	popupWin = window.open(url,'new_page','width=600,height=400,scrollbars=yes,resizable=yes');	return false;}</script>
                          </div></td>
                      </tr>
                      <tr style="display:none;">
                        <td align="center" bgcolor="f5f5f5"><table width="100%" border="0" cellspacing="2" cellpadding="4">
                            <tr> 
                              <td width="10%"><div align="right">网站优化标题栏</div></td>
                              <td width="90%"><input name="kw1" type="text" id="kw12" size="80" value="<%=kw1%>"></td>
                            </tr>
                            <tr> 
                              <td><div align="right">网站优化关键字栏</div></td>
                              <td><input name="kw2" type="text" id="kw22" size="80" value="<%=kw2%>"></td>
                            </tr>
                            <tr> 
                              <td><div align="right">网站优化描述栏</div></td>
                              <td><input name="kw3" type="text" id="kw32" size="80" value="<%=kw3%>"></td>
                            </tr>
                            <tr> 
                              <td><div align="right">类别热门关键词</div></td>
                              <td><input name="kw4" type="text" id="kw42" size="80" value="<%=kw4%>">
                                请用“|”隔开&nbsp; 如： 关键字1|关键字2|关键字3</td>
                            </tr>
                          </table></td>
                      </tr>
                      <tr> 
                        <td height="45" align="center"> <div align="left"> &nbsp;&nbsp; 
                            <input type="submit" name="Submit" value=" 修改 ">
                            <input name="editid" type="hidden" value="<%=classid%>">
                            <input name="gb" type="hidden" value="<%=request("gb")%>">
                            <a href="javascript:history.go(-1);">返回上一页</a> </div></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                      </tr>
                    </table>
                  </form></td>
              </tr>
            </table>
            <%
End Sub
'--------------------------------编辑类别sub代码结束---------------------------------
%> <%
'---------------------------------保存修改和编辑类别的sub开始---------------------------
Sub SaveClass()
	If Downasp.IsValidStr(Request("ClassName")) = False Then
		Errmsg =  "类别名称中含有非法字符。"
		Downasp.ShowMsg(Errmsg)
		Downasp.Back(-1)
		Response.End()
	End If
	if request("ClassName_sp")<>"" then 
	ClassName_sp=trim(request("ClassName_sp"))
	else
	ClassName_sp=trim(request("ClassName"))
	end if
	
	conn.execute("insert into pclass(classname,parentid,orders,gb,ishot,kw1,kw2,kw3,kw4,ClassName_sp) values('"+Trim(Request("ClassName"))+"','"+Request("ParentID")+"','"+Request("Orders")+"','"+Request("gb")+"',"+Request("ishot")+",'"+Trim(Request("kw1"))+"','"+Trim(Request("kw2"))+"','"+Trim(Request("kw3"))+"','"+Trim(Request("kw4"))+"','"+ClassName_sp+"')")
	Response.Redirect("admin_prodclass.asp?gb="&request("gb")&"")
End Sub
'------------------------------保存修改和编辑类别的sub结束------------------------------
%> <%
'------------------------------------删除类别sub开始------------------------------------
Sub DelClass()
	Dim classID,ID
	ID = Request("ParentID")
	classID = ID & "," & Downasp.selectclass(ID)
	classID = Replace(classID,",,",",")
	if right(classID,1) = "," Then classID = left(classID,len(classID)-1)
	if left(classID,1) = "," Then classID = right(classID,len(classID)-1)
	'查找是否有下级类别
	'set rs = conn.execute("select * from pclass where ParentID="&Request("ParentID")&"")
	'If not rs.eof then 
		'Downasp.ShowMsg("请先删除下级类别。")
		'Downasp.Back(-1)
		'Response.End()
	'End IF
	'rs.close
	'set rs = nothing
	'删除类别
	'pid=conn.execute("select ParentID from pclass where id="&id&"")(0)
	'if pid=0 then 
	'conn.execute("delete from ProdMain where larseq = '"&ID&"' ")
	'else
	'conn.execute("delete from ProdMain where gbseq = '"&ID&"' ")
	'end if
	
	conn.execute("Delete from pclass where id in ("&classID&")")
	'删除相关资源
	'待定，未完成
	'------------
	Response.Redirect("admin_prodclass.asp?gb="&request("gb")&"")
End Sub
'-----------------------------删除类别sub结束--------------------------
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