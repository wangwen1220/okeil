<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/lib/myasp.asp"-->
<%
Dim style, title, keywords, description, pro_type, pro_type_rule, img_path
style = "pros"
pro_type = "PU Case"
pro_type_rule = Replace(LCase(pro_type), " ", "_")
img_path = "/img/products/"&pro_type_rule&"/"
title = pro_type&" - Mobile Phone Case"
keywords = pro_type&", Mobile Phone Case"
description = "Lieko Co., Ltd. is the largest mobile phone cases supplier in Dongguan, the main products are "&pro_type&" and Mobile Phone "&pro_type&"."
%>
<!--#include virtual="/inc/head.asp"-->

<body id="<%=pro_type_rule%>" class="pros jqbox jqpager">
  <div id="wrap">
    <!--#include virtual="/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/inc/banner.asp"-->

        <h2 id="position"><a title="lieko.com" href="/index.asp">Home</a> <span>/</span> <strong><%=pro_type%></strong></h2>

        <div id="pros" class="case jqbox">
          <h1 title="<%=pro_type%>"><%=pro_type%></h1>

          <p id="jqpager"> </p>
          <%
          dim files : files = file_list(img_path)
          dim img, pros_model, i : i = 0
          for each img in split(files(1), "|")
            pros_model = replace(img, ".jpg", "")
            i = i + 1
          %>
          <div<%if i > 20 then echo " class='hide'"%>>
            <h3><%=pro_type&" "&pros_model%></h3>
            <a href="<%=img_path&img%>" title="<%=pro_type&" "&pros_model%>"><img src="<%=img_path&img%>" alt='<%=pros_model%>' /></a>
            <span class="number"><%=pros_model%></span>
          </div>
          <%next%>
        </div>

        <!--#include virtual="/inc/hot_products.asp"-->
      </div>

      <!--#include virtual="/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/inc/footer.asp"-->
</body>
</html>