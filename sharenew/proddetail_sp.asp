<%
Set rsprod=Server.CreateObject("ADODB.RecordSet") 
sqlprod="select * from ProdMain where gb='"&gbbase(0)&"' and Online=true and prodnum="&request.QueryString("prodnum")&""
rsprod.Open sqlprod,conn,1,3 
if rsprod.bof and rsprod.eof then
		response.write "<table><tr><td width=100% rowspan='2' height=150 align=center>Sorry! No Product Now.</td></tr></table>"
else
%>

<% if rsprod("MemoSpec")<>"" then %>
<%=rsprod("MemoSpec")%>
<%end if %>

<%
newtimes = rsprod("ClickTimes")+1
rsprod("ClickTimes")= newtimes
rsprod.update
rsprod.close
set rsprod=nothing
end if
%>
<div class="pubdate" style="margin-top:15px;"><a href="javascript:history.go(-1);">[volver]</a></div>
