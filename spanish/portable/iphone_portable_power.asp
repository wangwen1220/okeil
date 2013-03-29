<%
Dim style, title, keywords, description, pro_brand, pro_type, pro_type_rule, brand_type, brand_type_rule
style = "main"
pro_brand = "iPhone"
pro_type = "Energía Portátil"
brand_type = pro_brand&" "&pro_type
pro_type_rule = "portable_power"
brand_type_rule = Replace(LCase(pro_brand&"_"&pro_type_rule), " ", "_")
title = pro_brand&" 4/3G/3GS "&pro_type&" - Lieko "&pro_type&" Telefonía Móvil"
keywords = pro_brand&", "&pro_type&" Telefonía Móvil"
description = "Lieko "&pro_type&" for "&pro_brand&" 4/3G/3GS."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="type_list" class="pros">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->

        <h2 id="position"><a title="LIEKO Protectors para Telefonía Móvil" href="/spanish/index.asp">Inicio</a> <span>/</span> <a  title='<%=pro_type%>' href='<%=pro_type_rule%>.asp'><%=pro_type%></a> <span>/</span> <strong><%=brand_type%></strong></h2>

        <div id="pros">
          <h1 title="<%=brand_type%>"><%=brand_type%></h1>

          <p class="notice">Cómo está, aca tiene lista de unos productos, sobre más productos informaciones, por favor contácte con nuestros representantes de ventas, gracias.</p>

          <%
          Easp.db.PageSize = 10
          Easp.db.SetPager "pager", "{prev}| <span class='list'>{liststart}{list}{listend}</span> |{next}", Array("prev:&laquo; Anterior", "next:Siguiente &raquo;", "listlong:5", "listsidelong:1")
          Dim rs, brand, model, name, image, number, material, i
          Set rs = Easp.db.GPR("sql", "SELECT * FROM "&pro_type_rule&" WHERE brand='"&pro_brand&"' ORDER BY id DESC, number ASC")
          For i = 1 To rs.PageSize
            If rs.Eof Then Exit For
            brand = rs("brand")
            model = rs("model")
            name = replace(rs("title"),"Portable Power","Energía Portátil")
            image = rs("image")
            number = rs("number")
            capacity = rs("capacity")
            dimension = rs("dimension")
            weight = rs("weight")
            color = replace(replace(replace(replace(replace(rs("color"),"black","negro"),"blue","azul"),"silver","plata"),"purple","púrpura"),"red","rojo")
          %>
          <div class="list">
            <a class="img" href="<%=brand_type_rule%>_view.asp?name=<%=name&" "&number%>" title='<%=name%>'><img src="<%=image%>" alt='<%=name%>' width="150" height="150" /></a>
            <h3><a href="<%=brand_type_rule%>_view.asp?name=<%=name&" "&number%>" title='<%=name%>'><%=name%></a></h3>
            <ul>
              <li><span>Producto Número:</span> <%=number%></li>
              <li><span>Compatible marca&modelo:</span> <%=brand&" "&model%></li>
              <li><span>Capacilidad:</span> <%=capacity%></li>
              <li><span>Dimension:</span> <%=dimension%></li>
              <li><span>Weight:</span> <%=weight%></li>
              <li><span>Color:</span> <%=color%></li>
            </ul>
            <a class="more" href="<%=brand_type_rule%>_view.asp?name=<%=name&" "&number%>" title='<%=name%>'>Conozca más</a>
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