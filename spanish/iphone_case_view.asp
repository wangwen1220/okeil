<%
Dim style, title, keywords, description, pro_brand, pro_model, pro_name, pro_number
style = "main"
pro_brand = "iPhone"
pro_model = Request.QueryString("model")
pro_number = Request.QueryString("number")
pro_name = Request.QueryString("name")&" "&pro_number
title = pro_name&" - "&pro_brand&" "&pro_model&" Protector - Protectors para Telefonía Móvil"
keywords = pro_name&", "&pro_brand&" "&pro_model&" Protector - Protectors para Telefonía Móvil"
description = "Lieko Co., Ltd es el más grande proveedor de protector para celulares en Dongguan, los principales productos son "&pro_name&", "&pro_brand&" "&pro_model&" Protector y "&pro_brand&" "&pro_model&" Protectors para Telefonía Móvil."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="poster" class="pros hot jqpager">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->

        <h2 id="position"><a title="LIEKO Protectors para Telefonía Móvil" href="/spanish/index.asp">Inicio</a> <span>/</span> <a  title='<%=pro_brand&" "&pro_model%> Protector' href='/spanish/<%=pro_brand%>_case.asp?model=<%=pro_model%>'><%=pro_brand&" "&pro_model%> Protector</a> <span>/</span> <strong><%=pro_name%></strong></h2>

        <div id="pros">
          <h1 title='<%=pro_brand&" "&pro_model%>'><%=pro_name%></h1>

          <p id="jqpager"> </p>
          <%
          Dim rs, image, image_count, i
          On Error Resume Next
          Set rs = Easp.db.GRD("hot_selling", "number='"&pro_number&"'")
          image = rs("image")
          image_count = rs("image_count")
          For i = image_count To 1 Step -1
          %>
          <div class="poster">
            <img src='<%=Replace(image, ".jpg", "_poster_"&i&".jpg")%>' alt='<%=pro_name%>' width="780" />
          </div>
          <%
          Next
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