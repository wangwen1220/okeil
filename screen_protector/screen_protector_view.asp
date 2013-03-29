<%
Dim style, title, keywords, description, pro_brand, pro_type, pro_type_rule, brand_type, brand_type_rule, pro_name
style = "main"
pro_brand = Request.QueryString("brand")
pro_name = Request.QueryString("name")
pro_type = "Screen Protector"
brand_type = pro_brand&" "&pro_type
pro_type_rule = Replace(LCase(pro_type), " ", "_")
brand_type_rule = Replace(LCase(brand_type), " ", "_")
title = pro_name&" - Lieko Mobile Phone "&pro_type
keywords = pro_name&", "&pro_type&", Cell Phone "&pro_type&", Mobile Phone Cases"
description = "Lieko Co., Ltd. is the largest mobile phone cases supplier in Dongguan, the main products are Mobile Phone 4D Case, Bamboo Case, Wood Case, Combo Case, Silicone Case, Leather Case, Screen Protector and "&pro_name&"."
%>
<!--#include virtual="/inc/head.asp"-->

<body id="<%=brand_type_rule%>" class="pros">
  <div id="wrap">
    <!--#include virtual="/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/inc/banner.asp"-->

        <h2 id="position"><a title='Home' href='/index.asp'>Home</a> <span>/</span> <a  title='<%=brand_type%>' href='<%=brand_type_rule%>.asp'><%=brand_type%></a> <span>/</span> <strong><%=pro_name%></strong></h2>

        <div id="pros">
          <h1 title="<%=brand_type%>"><%=pro_name%></h1>

          <div class="poster">
            <img src='/img/products/<%=pro_type_rule%>/<%=pro_brand%>/<%=Easp.RegReplace(pro_name, "[ /]", "_")%>_poster_1.jpg' alt='<%=pro_name%>' width="780" />
          </div>
        </div>

        <!--#include virtual="/inc/hot_products.asp"-->
      </div>

      <!--#include virtual="/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/inc/footer.asp"-->
</body>
</html>