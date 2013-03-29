<%
Dim style, title, keywords, description, pro_type, pro_type_rule, brand_type, brand_type_rule, pro_name
style = "main"
pro_type = "Energía Portátil"
brand_type = "iPhone "&pro_type
pro_type_rule = "portable_power"
brand_type_rule = "iphone_portable_power"
pro_name = Request.QueryString("name")
title = pro_name&" - Lieko "&pro_type&" Telefonía Móvil"
keywords = pro_name&", "&pro_type
description = "Lieko "&pro_name&"."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="poster" class="pros">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->

        <h2 id="position"><a title="LIEKO Protectors para Telefonía Móvil" href="/spanish/index.asp">Inicio</a> <span>/</span> <a  title='<%=pro_type%>' href='<%=pro_type_rule%>.asp'><%=pro_type%></a> <span>/</span> <a  title='<%=brand_type%>' href='<%=brand_type_rule%>.asp'><%=brand_type%></a> <span>/</span> <strong><%=pro_name%></strong></h2>

        <div id="pros">
          <h1 title="<%=pro_name%>"><%=pro_name%></h1>

          <div class="poster">
            <img src='/img/products/<%=pro_type_rule%>/<%=Easp.RegReplace(replace(pro_name, pro_type, pro_type_rule), "[ /]", "_")%>_poster_1.jpg' alt='<%=pro_name%>' width="780" />
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