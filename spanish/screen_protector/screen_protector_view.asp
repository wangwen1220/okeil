<%
Dim style, title, keywords, description, pro_brand, pro_type, pro_type_rule, brand_type, brand_type_rule, pro_name
style = "main"
pro_brand = Request.QueryString("brand")
pro_name = Request.QueryString("name")
pro_type = "Pantalla Protector"
brand_type = pro_brand&" "&pro_type
pro_type_rule = "screen_protector"
brand_type_rule = Replace(LCase(pro_brand&" Screen Protector"), " ", "_")
title = pro_name&" - LIEKO "&pro_type
keywords = pro_name&", "&pro_type
description = "Lieko "&pro_name&"."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="<%=brand_type_rule%>" class="pros">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->

        <h2 id="position"><a title="LIEKO Protectors para Telefonía Móvil" href="/spanish/index.asp">Inicio</a> <span>/</span> <a  title='<%=brand_type%>' href='<%=brand_type_rule%>.asp'><%=brand_type%></a> <span>/</span> <strong><%=pro_name%></strong></h2>

        <div id="pros">
          <h1 title="<%=brand_type%>"><%=pro_name%></h1>

          <div class="poster">
            <img src='/img/products/<%=pro_type_rule%>/<%=pro_brand%>/<%=Easp.RegReplace(replace(replace(replace(pro_name, pro_type, pro_type_rule),"Esperjo","mirror"),"Privacidad","pricacy"), "[ /]", "_")%>_poster_1.jpg' alt='<%=pro_name%>' width="780" />
          </div>
        </div>

        <!--#include virtual="/spanish/inc/hot_products.asp"-->
      </div>

      <!--#include virtual="/spanish/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/spanish/inc/footer.asp"-->
</body>
</html>