<div align="center" style="padding:15px 0px;">
<%
	if currentpage = 1 then%>
		<font color=darkgray>首页</font>
	<%else%> 
		<font color="black">
		<a href="<%=request.ServerVariables("script_name")%>?page=1&remark=<%=request("remark")%>&keyword=<%=request("keyword")%>&starttime=<%=request("starttime")%>&gb=<%=gb%>&userid=<%=request("userid")%>&prodnumlevel=<%=request("prodnumlevel")%>&action2=<%=request("action2")%>">首页</a>
		<a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage-1%>&remark=<%=request("remark")%>&keyword=<%=request("keyword")%>&starttime=<%=request("starttime")%>&gb=<%=gb%>&userid=<%=request("userid")%>&prodnumlevel=<%=request("prodnumlevel")%>&action2=<%=request("action2")%>">前页</a>
		</font>
	<%end if
	 if currentpage = n then%> 
		<font color=darkgray >后页</font>
	<%else%> 
		<font color=black >
		<a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage+1%>&remark=<%=request("remark")%>&keyword=<%=request("keyword")%>&starttime=<%=request("starttime")%>&gb=<%=gb%>&userid=<%=request("userid")%>&prodnumlevel=<%=request("prodnumlevel")%>&action2=<%=request("action2")%>">下页</a>
		<a href="<%=request.ServerVariables("script_name")%>?page=<%=n%>&remark=<%=request("remark")%>&keyword=<%=request("keyword")%>&starttime=<%=request("starttime")%>&gb=<%=gb%>&userid=<%=request("userid")%>&prodnumlevel=<%=request("prodnumlevel")%>&action2=<%=request("action2")%>">末页</a>
		 </font>
<%end if%>
	
<font color="black">总:<%=currentpage%>/<%=n%>页&nbsp;&nbsp; [<%=msg_per_page%>条信息/页]</font> 
  <select name="menu1" onChange="MM_jumpMenu('self',this,0)">
  <option value="?page=<%=currentpage%>&remark=<%=request("remark")%>&keyword=<%=request("keyword")%>&starttime=<%=request("starttime")%>&gb=<%=gb%>&userid=<%=request("userid")%>&prodnumlevel=<%=request("prodnumlevel")%>&action2=<%=request("action2")%>">第<%=currentpage%>页</option>
  <%for i=1 to n
  	if currentpage <> i then%>
    <option value="?page=<%=i%>&remark=<%=request("remark")%>&keyword=<%=request("keyword")%>&starttime=<%=request("starttime")%>&gb=<%=gb%>&userid=<%=request("userid")%>&prodnumlevel=<%=request("prodnumlevel")%>&action2=<%=request("action2")%>">第<%=i%>页</option>
  <%end if
  	next%>
  </select>
</div>