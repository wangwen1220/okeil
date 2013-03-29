<div id="hot_products"> 
        <h2 title="Hot Products">BEST SELLER</h2>
        <div class="tips"> <span class="current">1</span> <span>2</span> <span>3</span> 
          <span>4</span> </div>
        <div class="btn"> <span class="prev">Prev</span> <span class="next">Next</span> 
        </div>
        <div class="prolist"> 
          <ul>
		  <%
sqlprod="select top 20 * from ProdMain where gb='"&gbbase(0)&"' and remark='1' and Online = true order by xuhao desc "
Set rsprod=Server.CreateObject("ADODB.RecordSet") 
rsprod.open sqlprod,conn,1,1
if rsprod.bof and rsprod.eof then
response.write ""
else
do while not rsprod.eof 
		  %>
		  
            <li><a href="productdetail.asp?ProdNum=<%=rsprod("ProdNum")%>&parentid=<%=rsprod("gbseq")%>&larseq=<%=rsprod("larseq")%>"><img src='<%=rsprod("photo1")%>' alt="<%=rsprod("ProdID")%>|<%=rsprod("larcode")%>|<%=rsprod("searchtype")%>" border=0></a></a> 
              <h3><a href="productdetail.asp?ProdNum=<%=rsprod("ProdNum")%>&parentid=<%=rsprod("gbseq")%>&larseq=<%=rsprod("larseq")%>"><%=rsprod("prodid")%></a></h3>
            </li>
			
<%
rsprod.movenext
loop
end if
rsprod.close
set rsprod = nothing
%>
          </ul>
        </div>
        <div class="more"><a href="/products.asp?parentid=106">More</a></div>
      </div>