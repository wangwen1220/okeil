<%
Dim style, title, keywords, description, brand
style = "main"
brand = request.queryString("brand")
If brand = "Screen Protector" Then
  title="Blackberry iPad iPhone Pantalla Protector - Pantalla Protector Telefonía Móvil"
ElseIf brand <> "" Then
  title=brand&" Protector de cuero, Protector de silicona, Protector de bambú - Protectors para Telefonía Móvil"
Else
  title = "Blackberry, iPad 2, iPhone 4 Protectors para Telefonía Móvil - LIEKO"
End If
description = "LIEKO suministro Protectors para Telefonía Móvil: Blackberry Protector, iPad 2 Protector, iPhone 4 Protector, "&brand&" Protector de madera, Protector de silicona, Protector de bambú & Pantalla Protector."
keywords = "iPhone Protector, iPad Protector, Blackberry Protector, Pantalla Protector, "&brand&" protectors para Telefonía Móvil"
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="home">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->

        <h2 id="position">Bienvenido a <a href="/spanish/index.asp" title="LIEKO Protectors para Telefonía Móvil">LIEKO</a></h2>

        <div id="pros" class="prolist">
        <%if brand = "Screen Protector" then%>
          <h1>Pantalla Protector</h1>
        <%elseif brand <> "" then%>
          <h1><%=brand%> Protectors</h1>
        <%else%>
          <h1>Protectors para Telefonía Móvil</h1>
        <%end if%>

          <%
          Easp.db.PageSize = 15
          Easp.db.SetPager "jqpager", "<p id='jqpager'>{prev} <span class='qp_counter'>{pageindex} / {pagecount}</span> {next}</p>", Array("prev: &laquo; Anterior", "next: Siguiente &raquo;", "disabledclass: qp_disabled")
          Easp.db.SetPager "pager", "{prev}| <span class='list'>{liststart}{list}{listend}</span> |{next}", Array("prev:&laquo; Anterior", "next:Siguiente &raquo;", "listlong:10", "listsidelong:1")
          Dim rs, pro_title, image, condition, i
          if brand <> "" then condition = "WHERE brand = '"&brand&"'"
          Set rs = Easp.db.GPR("sql", "SELECT title, image FROM prolist "&condition&" ORDER BY id DESC")
          Easp.W Easp.db.GetPager("jqpager")
          %>
          <p class="notice">Verifique la información de cada producto, por favor visite la <strong>"venta caliente"</strong> sección.</p>
          <p class="notice">Tambien podria buscar producto con el numero de producto en la barra de búsqueda, por ejemplo: IP-4-A</p>
          <ul class="brand_list">
            <li><strong>Categorías:</strong></li>
            <li><a href="/spanish/index.asp?brand=iPad" title="iPad Protector">iPad</a></li>
            <li><a href="/spanish/index.asp?brand=iPhone" title="iPhone Protector">iPhone</a></li>
            <li><a href="/spanish/index.asp?brand=Blackberry" title="Blackberry Protector">Blackberry</a></li>
            <li><a href="/spanish/index.asp?brand=Nokia" title="Nokia Protector">Nokia</a></li>
            <!--<li><a href="/spanish/index.asp?brand=HTC" title="HTC Phone Protector">HTC</a></li>-->
            <li class="last"><a href="/spanish/index.asp?brand=Samsung" title="Samsung Phone Protector">Samsung</a></li>
            <li class="last"><a href="/spanish/index.asp?brand=Screen Protector" title="Pantalla Protector">Pantalla Protector</a></li>
            <li class="view_all"><a href="/spanish/index.asp" title="Ver todos Protectors para Telefonía Móvil">Ver todos</a></li>
          </ul>
          <%
          For i = 1 To rs.PageSize
            If rs.eof Then Exit For
            pro_title = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(LCase(rs("title")),"case","protector"),"list","lista"),"bamboo","bambú"),"silicone","silicon"),"hybrid","híbrido"),"bumper","bamper"),"design","diseño"),"hard","duro"),"leather","cuero"),"skin","piel"),"screen protector","Pantalla Protector")
            image = rs("image")
          %>
          <div class="box">
            <a class="img" href="/spanish/public/poster_view.asp?title=<%=pro_title%>&image=<%=image%>" title='<%=pro_title%>'><img src="<%=image%>" alt='<%=pro_title%>' width="150" height="150" /></a>
            <h3><a href="/spanish/public/poster_view.asp?title=<%=pro_title%>&image=<%=image%>" title='<%=pro_title%>'><%=pro_title%></a></h3>
          </div>
          <%
            rs.MoveNext()
          Next
          If rs.eof and rs.bof Then
          %>
          <p class="error">No hay productos relacionados favor, elija una categoría diferente</p>
          <%Else%>
          <div class="pager">
            <%Easp.W Easp.db.GetPager("pager")%>
          </div>
          <%
          End If
          Easp.C(rs)
          %>
        </div>

        <!--#include virtual="/spanish/inc/hot_products.asp"-->
      </div>

      <!--#include virtual="/spanish/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/spanish/inc/footer.asp"-->
</body>
</html>