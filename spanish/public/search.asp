<%
dim style, title, keywords, description, wd, wd_en
wd = Request.QueryString("wd")
wd_en = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(LCase(Trim(Request.QueryString("wd"))),"protector","Case"),"bambú","Bamboo"),"madera","Wood"),"silicon","Silicone"),"híbrido","Hybrid"),"cuero","Leather"),"duro","Hard"),"metálico","Metal"),"bamper","Bumper"),"brazalete","Armband"),"diamante","Diamond"),"cristal","Crystal"),"electrochapado","Electroplated"),"Esponja","sponge")
style = "main"
title="Buscar "&wd&" Resultados - LIEKO Protectors para Telefonía Móvil"
keywords="Blackberry Protectors, iPhone Protectors, iPad Protectors, Mobile Phone Protectors, Protectors para Telefonía Móvil"
description="Resultados de la búsqueda de Lieko Protectors para Telefonía Móvil: Blackberry Protectors, iPhone Protectors, iPad Protectors, HTC Protectors, Nokia Protectors y "&wd&" etc."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="search" class="pros">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->

        <h2 id="position"><a title="LIEKO Protectors para Telefonía Móvil" href="/spanish/index.asp">Inicio</a> <span>/</span> <strong>Resultados de la búsqueda de <%=wd%></strong></h2>

        <div id="pros" class="search">
          <h1 title="<%=wd%>">Información sobre el producto</h1>

          <%
          Easp.db.PageSize = 20
          Easp.db.SetPager "jqpager", "<p id='jqpager'>{prev} <span class='qp_counter'>{pageindex} / {pagecount}</span> {next}</p>", Array("prev: &laquo; Anterior", "next: Siguiente &raquo;", "disabledclass: qp_disabled")
          Easp.db.SetPager "pager", "{prev}| <span class='list'>{liststart}{list}{listend}</span> |{next}", Array("prev:&laquo; Anterior", "next:Siguiente &raquo;", "listlong:10", "listsidelong:1")
          Dim rs, number, model, brand, name, image, wdn, i
          wdn = wd_en
          If IsNumeric(Right(wd_en,1)) and InStrRev(wd_en,"-") Then wdn = left(wd_en,InStrRev(wd_en,"-")-1)
          Set rs = Easp.db.GPR("sql", "SELECT number, model, brand, title, image FROM hot_selling WHERE title LIKE '%"&Replace(wd_en," ","%")&"%' OR number LIKE '"&wdn&"' ORDER BY number DESC")
          Easp.W Easp.db.GetPager("jqpager")
          %>
          <p class="notice">Buscar "<%=wd%>" obtener <%=rs.RecordCount%> resultados</p>
          <%
          For i = 1 To rs.PageSize
            If rs.eof or rs.bof Then Exit For
            number = rs("number")
            model = rs("model")
            brand = rs("brand")
            name = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(LCase(rs("title")),"case","Protector"),"bamboo","Bambú"),"wood","Madera"),"silicone","Silicon"),"hybrid","Híbrido"),"leather","Cuero"),"hard","Duro"),"metal","Metálico"),"bumper","Bamper"),"armband","Brazalete"),"diamond","Diamante"),"crystal","Cristal"),"electroplated","Electrochapado"),"sponge","Esponja")
            image = rs("image")
          %>
          <div class="box">
            <a class="img" href="/spanish/<%=brand%>_case_view.asp?name=<%=Server.URLEncode(name)%>&number=<%=number%>&model=<%=model%>" target="_blank" title='<%=name%>'><img src="<%=image%>" alt='<%=name%>' width="150" height="150" /></a>
            <h3><a href="/spanish/<%=brand%>_case_view.asp?name=<%=Server.URLEncode(name)%>&number=<%=number%>&model=<%=model%>" target="_blank" title='<%=name%>'><%=name%></a><span><%=number%></span></h3>
          </div>
          <%
            rs.MoveNext()
          Next
          If rs.eof and rs.bof Then
          %>
          <p class="error">No hay productos relacionados favor, elija otra marca.</p>
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