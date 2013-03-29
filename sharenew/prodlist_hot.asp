
<%
if request("keywords")<>"" then 
keywords=trim(request("keywords"))
sqlprod = "select * from ProdMain where gb='"&gbbase(0)&"' and prodid like '%"&keywords&"%' "
		if request("type")<>"" then 
		sqlprod=sqlprod + "and (gbseq ='"&request("type")&"' or midseq ='"&request("type")&"' or larseq ='"&request("type")&"') "
		end if 
else

sqlprod="select * from ProdMain where gb='"&gbbase(0)&"' and remarkb='1'  "
		
		if request("Parentid")<>"" then 
		sqlprod=sqlprod + "and gbseq='"&request("Parentid")&"'"
		end if 
		
		if request("larseq")<>"" then 
		sqlprod=sqlprod + "and larseq='"&request("larseq")&"'"
		end if 
		

		if request("remark")="1" then 
		sqlprod=sqlprod + "and remark ='1' "
		end if 
end if
sqlprod = sqlprod + " and Online = true order by remark desc , xuhao desc "
Set rsprod=Server.CreateObject("ADODB.RecordSet") 
rsprod.open sqlprod,conn,1,1

'if rsprod.recordcount=1 then 
'zid=rsprod("id")
'num=rsprod("prodnum")
'response.Redirect ("product.asp?ProdNum="&num&"&zid="&zid&"")
'end if 

if rsprod.bof and rsprod.eof then
response.write "<center>暂时没有找到相关的产品信息！</center>"
else

'---------------------分页用代码开始-------------------------
'Dim currentpage, page_count, Pcount, MaxPage
'Dim totalrec, endpage
currentPage = request("page")
	If currentpage = "" Then
	currentpage = 1
	Else
	currentpage = CLng(currentpage)
		If Err Then
		currentpage = 1
		Err.Clear
		End If
	End If
rsprod.PageSize = Cint(35)
MaxPage = Cint(35)
rsprod.AbsolutePage = currentpage
page_count = 0
totalrec = rsprod.recordcount
Pcount = rsprod.PageCount
page1=currentpage
if page1 = totalrec then page1=0
'--------------------分页用代码结束---------------------------
%>
<ul>
<%do while irecordsshown<MaxPage and NOT rsprod.EOF%>
<li><a href="product_show.asp?ProdNum=<%=rsprod("ProdNum")%>&parentid=<%=rsprod("gbseq")%>&larseq=<%=rsprod("larseq")%>"><img src='<%=rsprod("ssmallpath1")%>' alt="<%=rsprod("ProdID")%>|<%=rsprod("larcode")%>|<%=rsprod("searchtype")%>" border=0></a><p><a href="product_show.asp?ProdNum=<%=rsprod("ProdNum")%>&parentid=<%=rsprod("gbseq")%>&larseq=<%=rsprod("larseq")%>"><%=rsprod("prodid")%></a></p></li>
<%
irecordsshown = irecordsshown +1
rsprod.movenext
loop
%>
</ul>

<div class="pubdate_new" style="margin-bottom:15px;">
<%
down=""
if request("Parentid")<>"" then
down=down & "&parentid="&request("Parentid")&""
end if
if request("remark")<>"" then
down= down & "&remark="&request("remark")&""
end if
if request("keywords")<>"" and request("Keywords")<>"输入关键字" and request("keywords")<>"please input keywords" then 
down= down & "&keywords="&request("keywords")&""
end if 
if request("larseq")<>"" then 
down= down & "&larseq="&request("larseq")&""
end if 
id=request("id")
urlaaa=request.ServerVariables("script_name")
%>

    Product Page(s):&nbsp;<%
	If currentpage > 10 Then
	'加入特殊搜索的分页代码
	response.Write " <a href="""&urlaaa&".asp?down="&down&"&page=1"" class='pagelist'>1</a> ..."
	End If
	If Pcount>currentpage + 9 Then
	endpage = currentpage + 9
	Else
	endpage = Pcount
	End If
	For i = currentpage -9 To endpage
	If Not i<1 Then
	If i = CLng(currentpage) Then
	response.Write " <font class='red'>"&i&"</font> "
	Else
	response.Write " <a href="""&urlaaa&"?down="&down&"&page="&i&""" class='pagelist'>"&i&"</a> "
	End If
	End If
	Next
	If currentpage + 9 < Pcount Then
	response.Write "... <a href="""&urlaaa&"?down="&down&"&page="&Pcount&"""  class='pagelist'>"&Pcount&"</a> "
	'特殊搜索分页代码加入完毕
	End If
	%>
</div>
<%
end if
rsprod.close
set rsprod=nothing
%>
