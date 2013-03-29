<%
Dim style, title, keywords, description,pro_type, pro_type_rule, pro_name
style = "main"
pro_type = "Portable Power"
pro_type_rule = Replace(LCase(pro_type), " ", "_")
pro_name = Request.QueryString("name")
title = pro_name&" - Lieko Mobile Phone Cases"
keywords = pro_name&", "&pro_type&", Cell Phone "&pro_type&", Mobile Phone Cases"
description = "Lieko Co., Ltd. is the largest mobile phone cases supplier in Dongguan, the main products are 4D Case, Bamboo Case, Wood Case, Combo Case, Silicone Case, Leather Case, Screen Protector and "&pro_name&"."
%>
<!--#include virtual="/inc/head.asp"-->

<body id="poster" class="pros">
  <div id="wrap">
    <!--#include virtual="/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/inc/banner.asp"-->

        <h2 id="position"><a title='Home' href='/index.asp'>Home</a> <span>/</span> <a  title='<%=pro_type%>' href='<%=pro_type_rule%>.asp'><%=pro_type%></a> <span>/</span> <strong><%=pro_name%></strong></h2>

        <div id="pros">
          <h1 title="<%=pro_name%>"><%=pro_name%></h1>

          <div class="poster">
            <img src='/img/products/<%=pro_type_rule%>/<%=Easp.RegReplace(pro_name, "[ /]", "_")%>_poster_1.jpg' alt='<%=pro_name%>' width="780" />
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