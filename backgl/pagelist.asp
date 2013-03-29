<div align="center" style="padding-top:5px;">
<%
	if currentpage = 1 then%>
		<font color=darkgray>首页</font>
	<%else%> 
		<font color="black">
		<a href="<%=request.ServerVariables("script_name")%>?page=1&type=<%=request("type")%>&keyword=<%=request("keyword")%>&online=<%=request("online")%>&gb=<%=gb%>&remark=<%=request("remark")%>&remarka=<%=request("remarka")%>&remarkb=<%=request("remarkb")%>">首页</a>
		<a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage-1%>&type=<%=request("type")%>&keyword=<%=request("keyword")%>&online=<%=request("online")%>&gb=<%=gb%>&remark=<%=request("remark")%>&remarka=<%=request("remarka")%>&remarkb=<%=request("remarkb")%>">前页</a>
		</font>
	<%end if
	 if currentpage = n then%> 
		<font color=darkgray >后页</font>
	<%else%> 
		<font color=black >
		<a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage+1%>&type=<%=request("type")%>&keyword=<%=request("keyword")%>&online=<%=request("online")%>&gb=<%=gb%>&remark=<%=request("remark")%>&remarka=<%=request("remarka")%>&remarkb=<%=request("remarkb")%>">下页</a>
		<a href="<%=request.ServerVariables("script_name")%>?page=<%=n%>&type=<%=request("type")%>&keyword=<%=request("keyword")%>&online=<%=request("online")%>&gb=<%=gb%>&remark=<%=request("remark")%>&remarka=<%=request("remarka")%>&remarkb=<%=request("remarkb")%>">末页</a>
		 </font>
<%end if%>
	
<font color="black">总:<%=currentpage%>/<%=n%>页&nbsp;&nbsp;总共:<%=totalrec%>条信息 [<%=msg_per_page%>条信息/页]</font> 
  <select name="menu1" onChange="MM_jumpMenu('self',this,0)">
  <option value="?page=<%=currentpage%>&type=<%=request("type")%>&keyword=<%=request("keyword")%>&online=<%=request("online")%>&gb=<%=gb%>&remark=<%=request("remark")%>&remarka=<%=request("remarka")%>&remarkb=<%=request("remarkb")%>">第<%=currentpage%>页</option>
  <%for i=1 to n
  	if currentpage <> i then%>
    <option value="?page=<%=i%>&type=<%=request("type")%>&keyword=<%=request("keyword")%>&online=<%=request("online")%>&gb=<%=gb%>&remark=<%=request("remark")%>&remarka=<%=request("remarka")%>&remarkb=<%=request("remarkb")%>">第<%=i%>页</option>
  <%end if
  	next%>
  </select>
</div>