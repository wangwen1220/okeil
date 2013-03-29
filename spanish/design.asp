<%
dim style, title, keywords, description, brand
style = "main"
if request.querystring("brand") <> "" then brand = request.queryString("brand")
title="Protectors para Telefonía Móvil Diseño - Lieko"
if brand <> "" then title=brand&" "&title
keywords="Protectors para Telefonía Móvil Diseño"
description="LIEKO Protectors para Telefonía Móvil Diseño."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="design" class="pros">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->
        <%if brand <> "" then%>
       <h2 id="position"><a title="LIEKO Protectors para Telefonía Móvil" href="/spanish/index.asp">Inicio</a> <span>/</span> <a href="design.asp" title="Diseño">Diseño</a> <span>/</span> <strong><%=brand%></strong></h2>
        <%else%>
       <h2 id="position"><a title="LIEKO Protectors para Telefonía Móvil" href="/spanish/index.asp">Inicio</a> <span>/</span> <strong>Diseño</strong></h2>
        <%end if%>

        <div id="pros" class="prolist">
          <h1 title="Diseño"><%if brand <> "" then%><%=brand&" "%><%end if%>Diseño para celular protector</h1>

          <%
          Easp.db.PageSize = 9
          Easp.db.SetPager "jqpager", "<p id='jqpager'>{prev} <span class='qp_counter'>{pageindex} / {pagecount}</span> {next}</p>", Array("prev: &laquo;  Anterior", "next: Siguiente &raquo;", "disabledclass: qp_disabled")
          Easp.db.SetPager "pager", "{prev}| <span class='list'>{liststart}{list}{listend}</span> |{next}", Array("prev:&laquo;  Anterior", "next:Siguiente &raquo;", "listlong:10", "listsidelong:1")
          Dim rs, pro_title, image, condition, i
          if brand <> "" then condition = "WHERE brand = '"&brand&"'"
          Set rs = Easp.db.GPR("sql", "SELECT title, image FROM design "&condition&" ORDER BY upload_time DESC, id ASC")
          Easp.W Easp.db.GetPager("jqpager")
          %>
          <p class="notice">Todos los protectores para celular pueden producir depende de los siguientes Diseño.</p>
          <ul class="brand_list<%if brand <> "" then%><%=" "&brand%><%end if%>">
            <li><strong>Categorías:</strong></li>
            <li<%if brand = "LA" then%> class="current"<%end if%>><a href="design.asp?brand=LA" title="LA Diseño">LA</a></li>
            <li<%if brand = "LB" then%> class="current"<%end if%>><a href="design.asp?brand=LB" title="LB Diseño">LB</a></li>
            <li<%if brand = "LC" then%> class="current"<%end if%>><a href="design.asp?brand=LC" title="LC Diseño">LC</a></li>
            <li<%if brand = "LD" then%> class="current"<%end if%>><a href="design.asp?brand=LD" title="LD Diseño">LD</a></li>
            <li<%if brand = "LE" then%> class="current"<%end if%>><a href="design.asp?brand=LE" title="LE Diseño">LE</a></li>
            <li<%if brand = "LF" then%> class="current"<%end if%>><a href="design.asp?brand=LF" title="LF Diseño">LF</a></li>
            <li<%if brand = "LG" then%> class="current"<%end if%>><a href="design.asp?brand=LG" title="LG Diseño">LG</a></li>
            <li<%if brand = "QJC" then%> class="current"<%end if%>><a href="design.asp?brand=QJC" title="QJC Diseño">QJC</a></li>
            <li<%if brand = "QJL" then%> class="current"<%end if%>><a href="design.asp?brand=QJL" title="QJL Diseño">QJL</a></li>
            <li<%if brand = "QJP" then%> class="current"<%end if%>><a href="design.asp?brand=QJP" title="QJP Diseño">QJP</a></li>
            <li<%if brand = "QJT" then%> class="current"<%end if%>><a href="design.asp?brand=QJT" title="QJT Diseño">QJT</a></li>
            <li<%if brand = "QJW" then%> class="current"<%end if%>><a href="design.asp?brand=QJW" title="QJW Diseño">QJW</a></li>
            <li class="last<%if brand = "QJY" then%> current<%end if%>"><a href="design.asp?brand=QJY" title="QJY Diseño">QJY</a></li>
            <li class="view_all"><a href="design.asp" title="Ver todos Diseño">Ver todos</a></li>
          </ul>
          <%
          For i = 1 To rs.PageSize
            If rs.eof Then Exit For
            pro_title = Replace(rs("title"),"Design","Diseño")
            image = rs("image")
          %>
          <div class="box">
            <a class="img" href="public/poster_view.asp?title=<%=pro_title%>&image=<%=image%>" title='<%=pro_title%>'><img src="<%=image%>" alt='<%=pro_title%>' width="150" height="150" /></a>
            <h3><a href="public/poster_view.asp?title=<%=pro_title%>&image=<%=image%>" title='<%=pro_title%>'><%=pro_title%></a></h3>
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