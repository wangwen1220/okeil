<%
Dim style, title, keywords, description, pro_type, pro_type_rule
style = "pros"
pro_type = "Hybrid Case"
pro_type_rule = Replace(LCase(pro_type), " ", "_")
title = pro_type&" - Mobile Phone "&pro_type&" - Cell Phone Case"
keywords = pro_type&", Cell Phone "&pro_type&", Mobile Phone Case"
description = pro_type&", Cell Phone "&pro_type&", Mobile Phone Case from China Lieko Co., Ltd."
%>
<!--#include virtual="/inc/head.asp"-->

<body id="<%=pro_type_rule%>" class="pros jqbox">
  <div id="wrap">
    <!--#include virtual="/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/inc/banner.asp"-->

        <h2 id="position"><a title="lieko.com" href="/index.asp">Home</a> <span>/</span> <strong><%=pro_type%></strong></h2>

        <div id="pros" class="case">
          <h1 title="<%=pro_type%>"><%=pro_type%></h1>

          <%
          Easp.db.SetPager "pager", "<p id='pager'>{prev} | <span class='list'><em>Page:</em> {list}</span> | {next}<span class='jump'>Page {jump} of {pagecount}</span></p>", Array("prev:&laquo; Previous", "next:Next &raquo;", "listlong:10", "jump:select")
          Easp.db.SetPager "jqpager", "<p id='jqpager'>{prev} <span class='qp_counter'>{pageindex} / {pagecount}</span> {next}</p>", Array("prev: &laquo; Prev", "next: Next &raquo;", "disabledclass: qp_disabled")
          Dim pros_rs, pros_title, pros_image, pros_no, i
          Set pros_rs = Easp.db.GPR(0, Array(pro_type_rule&":title,image,number", "", "id Desc"))
          Easp.W Easp.db.GetPager("jqpager")
          For i = 1 To pros_rs.PageSize
            If pros_rs.Eof Then Exit For
            pros_title = pros_rs("title")
            pros_image = Replace(pros_rs("image"),".","_thumb.")
            pros_no = pros_rs("number")
            Easp.W "<div><h3><a href='"&pros_rs("image")&"' title='"&pros_title&"'>"&pros_title&"</a></h3><a href='/public/view.asp?type="&pro_type&"&title="&pros_title&"' title='"&pros_title&"'><img src='"&pros_image&"' alt='"&pros_title&"' width='138' height='138' /></a><span class='number'>"&pros_no&"</span></div>"
            pros_rs.MoveNext()
          Next
          Easp.W Easp.db.GetPager("pager")
          Easp.C(pros_rs)
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