<%
Dim style, title, keywords, description, pro_type, pro_type_rule
style = "main"
pro_type = "Energía Pulsera"
pro_type_rule = "energy_bracelet"
title = "Silicone "&pro_type&" - Lieko "&pro_type
keywords = "Silicone "&pro_type
description = "Lieko "&pro_type&" for iPhone 4/3G/3GS, iPad, Blackberry y Nokia."
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

          <p class="notice">Cómo está, aca tiene lista de unos productos, sobre más productos informaciones, por favor contácte con nuestros representantes de ventas, gracias.</p>

          <%
          Easp.db.PageSize = 10
          Easp.db.SetPager "pager", "{prev}| <span class='list'>{liststart}{list}{listend}</span> |{next}", Array("prev:&laquo; Anterior", "next:Siguiente &raquo;", "listlong:5", "listsidelong:1")
          Dim rs, name, image, number, material, production_time, min_order, color, i
          Set rs = Easp.db.GPR("sql", "SELECT * FROM "&pro_type_rule&" ORDER BY id DESC, number ASC")
          For i = 1 To rs.PageSize
            If rs.Eof Then Exit For
            name = Replace(rs("title"),"Energy Bracelet",pro_type)
            image = rs("image")
            number = rs("number")
            material = rs("material")
            production_time = rs("production_time")
            min_order = rs("min_order")
            color = replace(replace(replace(replace(replace(rs("color"),"black","negro"),"blue","azul"),"silver","plata"),"purple","púrpura"),"red","rojo")
          %>
          <div class="list">
            <a class="img" href="<%=pro_type_rule%>_view.asp?name=<%=name&" "&number%>" title='<%=name%>'><img src="<%=image%>" alt='<%=name%>' width="150" height="150" /></a>
            <h3><a href="<%=pro_type_rule%>_view.asp?name=<%=name&" "&number%>" title='<%=name%>'><%=name%></a></h3>
            <ul>
              <li><span>Producto Número:</span> <%=number%></li>
              <li><span>Tiempo de Producción:</span> <%=production_time%> días</li>
              <li><span>Cantidad Mínima de pedido:</span> <%=min_order%> pcs</li>
              <li><span>Material:</span> <%=material%></li>
              <li><span>Color:</span> <%=color%></li>
              <li><span>Detalle de embalaje:</span> PE paquete y Blister paquete</li>
            </ul>
            <a class="more" href="<%=pro_type_rule%>_view.asp?name=<%=name&" "&number%>" title='<%=name%>'>Conozca más</a>
          </div>
          <%
            rs.MoveNext()
          Next
          Easp.C(rs)
          %>
          <div class="pager"><%Easp.W Easp.db.GetPager("pager")%></div>
        </div>

        <!--#include virtual="/spanish/inc/hot_products.asp"-->
      </div>

      <!--#include virtual="/spanish/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/spanish/inc/footer.asp"-->
</body>
</html>