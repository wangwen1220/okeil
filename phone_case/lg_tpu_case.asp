<%
Dim style, title, keywords, description, pro_type, pro_type_rule, pro_brand, pro_name, pro_name_rule
style = "pros"
pro_type = "TPU Case"
pro_brand = "LG"
pro_type_rule = Replace(LCase(pro_type), " ", "_")
pro_name = pro_brand&" "&pro_type
pro_name_rule = Replace(LCase(pro_name), " ", "_")
title = pro_name&" - "&pro_brand&" Case - Cell Phone Case"
keywords = pro_name&", "&pro_type&", "&pro_brand&" Case, Cell Phone "&pro_type&", Mobile Phone Case"
description = pro_name&", "&pro_type&", Mobile Phone "&pro_type&", Cell Phone Case from China Lieko Co., Ltd."
%>
<!--#include virtual="/inc/head.asp"-->

<body id="<%=pro_name_rule%>" class="pros jqbox">
  <div id="wrap">
    <!--#include virtual="/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/inc/banner.asp"-->

        <h2 id="position"><a title="lieko.com" href="/index.asp">Home</a> <span>/</span> <a title="<%=pro_type%>" href="/phone_case/<%=pro_type_rule%>.asp"><%=pro_type%></a> <span>/</span> <strong><%=pro_name%></strong></h2>

        <div id="pros" class="case">
          <h1 title="<%=pro_name%>"><%=pro_name%></h1>

          <%
          Easp.db.SetPager "pager", "<p id='pager'>{prev} | <span class='list'><em>Page:</em> {list}</span> | {next}<span class='jump'>Page {jump} of {pagecount}</span></p>", Array("prev:&laquo; Previous", "next:Next &raquo;", "listlong:10", "jump:select")
          Easp.db.SetPager "jqpager", "<p id='jqpager'>{prev} <span class='qp_counter'>{pageindex} / {pagecount}</span> {next}</p>", Array("prev: &laquo; Prev", "next: Next &raquo;", "disabledclass: qp_disabled")
          Dim pros_rs, pros_title, pros_image, pros_no, i
          Set pros_rs = Easp.db.GPR(0, Array(pro_type_rule&":number, title, image", "brand = '"&pro_brand&"'", "id Desc"))
          Easp.W Easp.db.GetPager("jqpager")
          For i = 1 To pros_rs.PageSize
            If pros_rs.Eof Then Exit For
            pros_title = pros_rs("title")
            pros_image = Replace(pros_rs("image"),".","_thumb.")
            pros_no = pros_rs("number")
            Easp.W "<div><h3><a href='"&pros_rs("image")&"' title='"&pros_title&"'>"&pros_title&"</a></h3><a href='/public/view.asp?type="&pro_type&"&title="&pros_title&"&brand="&pro_brand&"' title='"&pros_title&"'><img src='"&pros_image&"' alt='"&pros_title&"' width='138' height='138' /></a><span class='number'>"&pros_no&"</span></div>"
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