<%
dim style, title, keywords, description, wd : wd = Trim(Request.QueryString("wd"))
style = "main"
title="Search for "&wd&" Results - LIEKO Mobile Phone Case"
keywords="Blackberry Case, iPhone Case, iPad Case, Mobile Phone Case"
description="Search Result of Lieko mobile phone cases: Blackberry Case, iPhone Case, iPad Case, HTC Case, Nokia Case and "&wd&" etc."
%>
<!--#include virtual="/inc/head.asp"-->

<body id="search" class="pros">
  <div id="wrap">
    <!--#include virtual="/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/inc/banner.asp"-->

        <h2 id="position"><a title="lieko.com" href="/index.asp">Home</a> <span>/</span> <strong>Search Results</strong></h2>

        <div id="pros" class="search">
          <h1 title="<%=wd%>">Product Information</h1>

          <%
          Easp.db.PageSize = 20
          Easp.db.SetPager "jqpager", "<p id='jqpager'>{prev} <span class='qp_counter'>{pageindex} / {pagecount}</span> {next}</p>", Array("prev: &laquo; Prev", "next: Next &raquo;", "disabledclass: qp_disabled")
          Easp.db.SetPager "pager", "{prev}| <span class='list'>{liststart}{list}{listend}</span> |{next}", Array("prev:&laquo; Prev", "next:Next &raquo;", "listlong:10", "listsidelong:1")
          Dim rs, number, model, brand, name, image, wdn, i
          wdn = wd
          If IsNumeric(Right(wd,1)) and InStrRev(wd,"-") Then wdn = left(wd,InStrRev(wd,"-")-1)
          Set rs = Easp.db.GPR("sql", "SELECT number, model, brand, title, image FROM hot_selling WHERE title LIKE '%"&Replace(wd," ","%")&"%' OR number LIKE '"&wdn&"' ORDER BY number DESC")
          Easp.W Easp.db.GetPager("jqpager")
          %>
          <p class="notice">Search for "<%=wd%>" get <%=rs.RecordCount%> results</p>
          <%
          For i = 1 To rs.PageSize
            If rs.eof or rs.bof Then Exit For
            number = rs("number")
            model = rs("model")
            brand = rs("brand")
            name = rs("title")
            image = rs("image")
          %>
          <div class="box">
            <a class="img" href="/<%=brand%>_case_view.asp?name=<%=Server.URLEncode(name)%>&number=<%=number%>&model=<%=model%>" target="_blank" title='<%=name%>'><img src="<%=image%>" alt='<%=name%>' width="150" height="150" /></a>
            <h3><a href="/<%=brand%>_case_view.asp?name=<%=Server.URLEncode(name)%>&number=<%=number%>&model=<%=model%>" target="_blank" title='<%=name%>'><%=name%></a><span><%=number%></span></h3>
          </div>
          <%
            rs.MoveNext()
          Next
          If rs.eof and rs.bof Then
          %>
          <p class="error">No related products Please choose a different brand.</p>
          <%Else%>
          <div class="pager">
            <%Easp.W Easp.db.GetPager("pager")%>
          </div>
          <%
          End If
          Easp.C(rs)
          %>
        </div>

        <!--#include virtual="/inc/hot_products.asp"-->
      </div>

      <!--#include virtual="/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/inc/footer.asp"-->
</body>
</html>