<%
Dim style, title, keywords, description, pro_type, pro_type_rule
style = "main"
pro_type = "Portable Power"
pro_type_rule = Replace(LCase(pro_type), " ", "_")
title = "iPhone, Blackberry "&pro_type&" - Lieko Mobile Phone Cases"
keywords = pro_type&", Mobile Phone Cases"
description = "Lieko Co., Ltd. is the largest mobile phone cases supplier in Dongguan, the main products are 4D Case, Bamboo Case, Wood Case, Combo Case, Silicone Case, Leather Case, Screen Protector and "&pro_type&" for iPhone 4/3G/3GS, iPad, Blackberry and Nokia."
%>
<!--#include virtual="/inc/head.asp"-->

<body id="type_list" class="pros">
  <div id="wrap">
    <!--#include virtual="/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/inc/banner.asp"-->

        <h2 id="position"><a title='Home' href='/index.asp'>Home</a> <span>/</span> <strong><%=pro_type%></strong></h2>

        <div id="pros">
          <h1 title="<%=pro_type%>"><%=pro_type%></h1>

          <p class="notice">How are you, there is some products list, about more products informations, please contact our sales representive thank you.</p>

          <%
          Easp.db.PageSize = 10
          Easp.db.SetPager "pager", "{prev}| <span class='list'>{liststart}{list}{listend}</span> |{next}", Array("prev:&laquo; Prev", "next:Next &raquo;", "listlong:5", "listsidelong:1")
          Dim rs, brand, model, name, image, number, material, i
          Set rs = Easp.db.GPR("sql", "SELECT * FROM "&pro_type_rule&" ORDER BY id DESC, number ASC")
          For i = 1 To rs.PageSize
            If rs.Eof Then Exit For
            brand = rs("brand")
            model = rs("model")
            name = rs("title")
            image = rs("image")
            number = rs("number")
            capacity = rs("capacity")
            dimension = rs("dimension")
            weight = rs("weight")
            color = rs("color")
          %>
          <div class="list">
            <a class="img" href="<%=pro_type_rule%>_view.asp?name=<%=name&" "&number%>" title='<%=name%>'><img src="<%=image%>" alt='<%=name%>' width="150" height="150" /></a>
            <h3><a href="<%=pro_type_rule%>_view.asp?name=<%=name&" "&number%>" title='<%=name%>'><%=name%></a></h3>
            <ul>
              <li><span>Product number:</span> <%=number%></li>
              <li><span>Compatible brand&Model:</span> <%=brand&" "&model%></li>
              <li><span>Capacity:</span> <%=capacity%></li>
              <li><span>Dimension:</span> <%=dimension%></li>
              <li><span>Weight:</span> <%=weight%></li>
              <li><span>Color:</span> <%=color%></li>
            </ul>
            <a class="more" href="<%=pro_type_rule%>_view.asp?name=<%=name&" "&number%>" title='<%=name%>'>See More</a>
          </div>
          <%
            rs.MoveNext()
          Next
          Easp.C(rs)
          %>
          <div class="pager"><%Easp.W Easp.db.GetPager("pager")%></div>
        </div>

        <!--#include virtual="/inc/hot_products.asp"-->
      </div>

      <!--#include virtual="/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/inc/footer.asp"-->
</body>
</html>