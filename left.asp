<div class="newbar"> <h2><a href="Newleather.asp"> Leather case</a></h2></div>
<div class="newbar"> <h2><a href="ComboCase.asp">Combo Case</a></h2></div>
<div id="sidebar">
       
	    <%
		'第一级分类
		sqllar="select * from PClass where gb='"&gbbase(0)&"' and ParentID=0 and id<>243 and id<>328 and id<>370 order by Orders asc"
		Set rs=Server.CreateObject("ADODB.RecordSet") 
		rs.Open sqllar,conn,1,1 
		if not rs.eof then
		do while not rs.eof 
		%>
	    
      <h2><a href="#" title="<%=rs("classname")%>"><%=rs("ClassName")%></a></h2>
	  	<%
		'第二级分类
		sqlmid="select * from pclass where ParentID="&rs("ID")&" order by orders asc"
		Set rsmid=Server.CreateObject("ADODB.RecordSet") 
		rsmid.Open sqlmid,conn,1,1 
		if not rsmid.eof then
		do while not rsmid.eof 
		%>
        <dl id="hot_selling" class="menu">
          <dt><a href="products.asp?parentid=<%=rsmid("id")%>" title="<%=rsmid("classname")%>"><%=rsmid("ClassName")%></a></dt>
		  
		  	  <%
			  '第三级分类
			  	sqltemp="select * from pclass where ParentID="&rsmid("ID")&" order by orders asc"
				Set rstemp=Server.CreateObject("ADODB.RecordSet") 
				rstemp.Open sqltemp,conn,1,1 
				if not rstemp.eof then
				do while not rstemp.eof 
			  %>
			  <dd><a href="products.asp?parentid=<%=rstemp("id")%>" title="<%=rstemp("classname")%>"><%=rstemp("ClassName")%></a></dd>
			  <%
			  	rstemp.movenext
				loop
				end if
				rstemp.close
				set rstemp = nothing
				'三级分类结束
			  %>
			  
		  </dl>  
		  <%
		  rsmid.movenext
		  loop
		  end if
		  rsmid.close
		  set rsmid =nothing
		  '二级分类结束
		  %>
		
		<%
		rs.movenext
		loop
		end if
		rs.close
		set rs = nothing
		'一级分类结束
		%>
        

        
      <h2><a href="#" title="LIEKO Mobile Phone Cases Links">Links</a></h2>
        <ul class="nav links">
		 <%
			set rs=Server.CreateObject("ADODB.RecordSet")
			sql= "select * from Link where gb='"&gbbase(0)&"'  order by xuhao desc "
			rs.open sql,conn,1,1
			if not rs.bof then
			do while not rs.eof 
		 %>
		 <li><a class="external_links"  href="<%=rs("Source")%>" target="_blank" title="LIEKO Mobile Phone Case"><%=rs("NewsTitle")%></a></li>
		 <%
			rs.movenext
			loop
			end if
			rs.close
			set rs = nothing
		 %>
        </ul>

        
      <h2><a href="###" title="LIEKO Contact" rel="nofollow">Contact</a></h2>
        <ul class="nav contact">
          <li><em>Email:</em> <a href="mailto:info@lieko.com">info@lieko.com</a></li>
          <li><em>MSN:</em> <a href="msnim:chat?contact=CEO@lieko.com">CEO@lieko.com</a></li>
          <li><em>Tel:</em> +86-769-81768061</li>
        </ul>
      </div>