<%
dim style, title, keywords, description
style = "main"
title="Blackberry, iPad 2 & iPhone 4 para Telefonía Móvil Protector by LIEKO"
keywords="Blackberry Protector, iPhone 4 Protector, iPad 2 Protector, para Telefonía Móvil Protectors"
description="LIEKO es el proveedor de China de para Telefonía Móvil Protectors: iPhone 4 Protector, iPad 2 Protector, Blackberry 8520/8530/9700/9800/9500/9900 Protector, HTC Protector, Samsung Protector, Cuero/silicona/TPU /bambú/madera Protector."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="newpros">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->
        <h2 id="position"><a title="lieko.com" href="/spanish/index.asp">Inicio</a> <span>/</span> <strong>Productos</strong></h2>

        <div id="pros">
          <h1>Nuevos Productos</h1>
          <%
          Easp.db.PageSize = 16
          Easp.db.SetPager "jqpager", "<p id='jqpager'>{prev} <span class='qp_counter'>{pageindex} / {pagecount}</span> {next}</p>", Array("prev: &laquo; Anterior", "next: Siguiente &raquo;", "disabledclass: qp_disabled")
          Easp.db.SetPager "pager", "{prev}| <span class='list'>{liststart}{list}{listend}</span> |{next}", Array("prev:&laquo; Anterior", "next:Siguiente &raquo;", "listlong:5", "listsidelong:1")
          Dim new_rs, new_name, new_brand, new_model, new_number, new_pic, i
          Set new_rs = Easp.db.GPR("sql", "SELECT title, brand, model, number, image FROM hot_selling WHERE status LIKE '%NEW%' ORDER BY status DESC")
          Easp.W Easp.db.GetPager("jqpager")

          For i = 1 To new_rs.PageSize
            If new_rs.Eof and new_rs.PageCount > 1 Then Exit For
            new_name = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(LCase(new_rs("title")),"case","protector"),"list","lista"),"bamboo","bambú"),"silicone","silicon"),"hybrid","híbrido"),"bumper","bamper"),"design","diseño"),"hard","duro"),"leather","cuero"),"skin","piel"),"screen protector","Pantalla Protector")
            new_brand = new_rs("brand")
            new_model = new_rs("model")
            new_number = new_rs("number")
            new_pic = new_rs("image")
          %>
          <div class='new box'>
            <a class="img" href="/spanish/<%=new_brand%>_case_view.asp?name=<%=Server.URLEncode(new_name)%>&number=<%=new_number%>&model=<%=new_model%>" title='<%=new_name%>'><img src="<%=new_pic%>" alt='<%=new_name%>' width="150" height="150" /></a>
            <h3><a href="/spanish/<%=new_brand%>_case_view.asp?name=<%=Server.URLEncode(new_name)%>&number=<%=new_number%>&model=<%=new_model%>" title='<%=new_name%>'><%=new_name%></a><span><%=new_number%></span></h3>
          </div>
          <%
            new_rs.movenext()
          Next
          Easp.db.C(new_rs)
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