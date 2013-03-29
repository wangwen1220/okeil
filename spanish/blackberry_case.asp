<%
Dim style, title, keywords, description, pro_brand, brand_type, pro_model, pro_name, pro_brand_rule, brand_type_rule
style = "main"
pro_brand = "Blackberry"
pro_model = "8520/8530"
brand_type = pro_brand&" Protector"
If Request.QueryString("model") <> "" Then pro_model = Request.QueryString("model")
pro_name = pro_brand&" "&pro_model&" Protector"
pro_brand_rule = Replace(LCase(pro_brand), " ", "_")
brand_type_rule = Replace(LCase(pro_brand&" Case"), " ", "_")
title = pro_name&" - "&brand_type&" - Protectors para Telefonía Móvil"
keywords = pro_name&", "&brand_type&", Protectors para Telefonía Móvil"
description = "Lieko Co., Ltd es el más grande proveedor de protector para celulares en Dongguan, los principales productos son "&pro_name&", "&brand_type&" y Protectors para Telefonía Móvil."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="type_list" class="pros hot">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->

        <h2 id="position"><a title="LIEKO Protectors para Telefonía Móvil" href="/spanish/index.asp">Inicio</a> <span>/</span> <strong><%=pro_name%></strong></h2>

        <div id="pros">
          <h1 title='<%=pro_brand&" "&pro_model%>'><%=pro_name%></h1>

          <%
          Easp.db.PageSize = 30
          Easp.db.SetPager "jqpager", "<p id='jqpager'>{prev} <span class='qp_counter'>{pageindex} / {pagecount}</span> {next}</p>", Array("prev: &laquo; Anterior", "next: Siguiente &raquo;", "disabledclass: qp_disabled")
          Easp.db.SetPager "pager", "{prev}| <span class='list'>{liststart}{list}{listend}</span> |{next}", Array("prev:&laquo; Anterior", "next:Siguiente &raquo;", "listlong:5", "listsidelong:1")
          Dim rs, name, image, number, material, production_time, min_order, i
          Set rs = Easp.db.GPR("sql", "SELECT title, number, material, image, production_time, min_order FROM hot_selling WHERE brand = '"&pro_brand&"' AND model = '"&pro_model&"' ORDER BY id DESC, number ASC")
          Easp.W Easp.db.GetPager("jqpager")
          %>
          <p class="notice">Cómo está, aca tiene lista de unos productos, sobre más productos informaciones, por favor contácte con nuestros representantes de ventas, gracias.</p>
          <%
          For i = 1 To rs.PageSize
            If rs.Eof Then Exit For
            name = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(LCase(rs("title")),"case","Protector"),"bamboo","Bambú"),"wood","Madera"),"silicone","Silicon"),"hybrid","Híbrido"),"leather","Cuero"),"hard","Duro"),"metal","Metálico"),"bumper","Bamper"),"armband","Brazalete"),"diamond","Diamante"),"crystal","Cristal"),"electroplated","Electrochapado"),"sponge","Esponja")
            image = rs("image")
            number = rs("number")
            material = Replace(Replace(Replace(Replace(Replace(LCase(rs("material")),"bamboo","Bambú"),"wood","Madera"),"silicone","Silicon"),"leather","Cuero"),"metal","Metálico")
            production_time = rs("production_time")
            min_order = rs("min_order")
          %>
          <div class="item">
            <a class="img" href="/spanish/<%=brand_type_rule%>_view.asp?name=<%=Server.URLEncode(name)%>&number=<%=number%>&model=<%=pro_model%>" title='<%=name%>'><img src="<%=image%>" alt='<%=name%>' width="150" height="150" /></a>
            <h3><a href="/spanish/<%=brand_type_rule%>_view.asp?name=<%=Server.URLEncode(name)%>&number=<%=number%>&model=<%=pro_model%>" title='<%=name%>'><%=name%></a></h3>
            <ul>
              <li class='number'><span>Producto Número:</span> <%=number%></li>
              <li><span>Compatible marca&modelo:</span> <%=pro_brand&" "&pro_model%></li>
              <li><span>Tiempo de Producción:</span> <%=production_time%> días</li>
              <li><span>Cantidad Mínima de pedido:</span> <%=min_order%> pcs</li>
              <li><span>Material:</span> <%=material%></li>
              <li><span>Detalle de embalaje:</span> PE paquete y Blister paquete</li>
            </ul>
            <a class="more" href="/spanish/<%=brand_type_rule%>_view.asp?name=<%=Server.URLEncode(name)%>&number=<%=number%>&model=<%=pro_model%>" title='<%=name%>'>Conozca más</a>
          </div>
          <%
            rs.MoveNext()
          Next
          Easp.C(rs)
          %>
          <div class="pager">
            <%Easp.W Easp.db.GetPager("pager")%>
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