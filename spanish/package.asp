<%
Dim style, title, keywords, description, pro_type, pro_type_rule
style = "main"
pro_type = "Paquete"
pro_type_rule = "package"
title = pro_type&" Telefonía Móvil - Clear Polybag, Colored Polybag, Blister Package, Gift Box"
keywords = pro_type&" Telefonía Móvil"
description = "Lieko Clear Polybag, Colored Polybag, Blister Package, Gift Box "&pro_type&" Telefonía Móvil for iPhone 4/3G/3GS, iPad, Blackberry y Nokia."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="type_list" class="pros">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->

        <h2 id="position"><a title="LIEKO Protectors para Telefonía Móvil" href="/spanish/index.asp">Inicio</a> <span>/</span> <strong><%=pro_type%></strong></h2>

        <div id="pros">
          <h1 title="<%=pro_type%>"><%=pro_type%></h1>

          <div class="poster">
            <img src='/img/products/hot_selling/<%=pro_type_rule%>/package_poster_1.jpg' alt='<%=pro_type%> Telefonía Móvil' width="780" />
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