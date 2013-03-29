<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/lib/myasp.asp"-->
<%
dim style, title, keywords, description, pro_type, pro_type_rule, pro_title, pro_brand, brand_type, page_position
style = "view"
pro_brand = Trim(Request.QueryString("brand"))
pro_type = Request.QueryString("type")
pro_title = Request.QueryString("title")
pro_type_rule = Replace(LCase(pro_type), " ", "_")
brand_type = pro_type
If pro_brand <> "" Then brand_type = pro_brand&" "&pro_type
title = brand_type&" - Cell Phone Case"
keywords = brand_type&", Cell Phone Case"
description = "China manufacturer lieko.com for supply Cell Phone Cases: "&brand_type&", Cell Phone Case"
%>
<!--#include virtual="/inc/head.asp"-->

<body class="pros view">
  <div id="wrap">
    <!--#include virtual="/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/inc/banner.asp"-->

        <%
        Dim rs, image, brand, model, brand_model, name, number, color, color_str, i
        Set rs = Easp.db.GetRecord(pro_type_rule, "title='"&pro_title&"'", "id ASC")
        image = rs("image")
        brand = rs("brand")
        model =rs("model")
        brand_model = brand
        If model <> "" Then brand_model = brand&" "&model
        name = brand_model&" "&pro_type
        If rs.fields(1).name = "number" Then
          number = rs("number")
        Else
          number = Replace(LCase(pro_title), " ", "_")
        End If
        If rs("color") <> "" Then color = Split(rs("color"), " ")
        Easp.db.C(rs)
        If IsEmpty(color) Then
          color_str = "many color"
        Else
        For i = 0 To UBound(color)
          color_str = color_str + "<li><a class='"&color(i)&"' href='"&str_replace("_[a-z0-9A-Z]+\.", image, "_"&color(i)&".")&"' title='"&color(i)&"'>"&color(i)&"</a></li>"
        Next
        color_str = "<ul id='color' class='jqbox'>" + color_str + "</ul>"
        End If
        If pro_brand <> "" Then
        %>
        <h2 id="position"><a title='lieko.com' href='/'>Home</a> <span>/</span> <a title='<%=pro_type%>' href='/phone_case/<%=pro_type_rule%>.asp'><%=pro_type%></a> <span>/</span> <a title='<%=name%>' href='/phone_case/<%=Replace(LCase(brand_type), " ", "_")%>.asp'><%=brand_type%></a> <span>/</span> <strong><%=name%></strong></h2>
        <%Else%>
        <h2 id="position"><a title='lieko.com' href='/'>Home</a> <span>/</span> <a title='<%=pro_type%>' href='/phone_case/<%=pro_type_rule%>.asp'><%=pro_type%></a> <span>/</span> <strong><%=name%></strong></h2>
        <%End If%>

        <div id="pros">
          <h1 title="<%=name%>"><%=pro_title%></h1>

          <div id="picview">
            <a class="view" href="<%=image%>" title="<%=pro_title%>"><img src="<%=image%>" alt="<%=pro_title%>" /></a>
            <a class="view_more jbox" href="<%=image%>" title="<%=pro_title%>">See larger image</a>
          </div>

          <div class="prosinfo">
            <h3>Product Details</h3>
            <dl>
              <dt>Product Name:</dt>
              <dd><%=pro_title%></dd>
              <dt>Product Number:</dt>
              <dd><%=number%></dd>
              <dt>Compatible Model:</dt>
              <dd><%=model%></dd>
              <dt>Type:</dt>
              <dd><%=pro_type%></dd>
              <dt>Color:</dt>
              <dd><%Response.Write(color_str)%></dd>
            </dl>
          </div>

          <%
          Dim relatedpros_rs, relatedpros_title, relatedpros_image, relatedpros_str, condition
          condition = Easp.IIF(pro_brand = "", "", "brand = '"&pro_brand&"'")
          Set relatedpros_rs = Easp.db.GetRandRecord(pro_type_rule&":id:8", condition)
          For i = 1 to relatedpros_rs.RecordCount
            relatedpros_title = relatedpros_rs("title")
            relatedpros_image = Replace(relatedpros_rs("image"),".","_thumb.")
            relatedpros_str = relatedpros_str + "<li><a class='img' href='/public/view.asp?type="&pro_type&"&amp;title="&Replace(relatedpros_title," ","%20")&"&brand="&pro_brand&"' title='"&relatedpros_title&"'><img src='"&relatedpros_image&"' alt='"&relatedpros_title&"' /></a><h4><a href='/public/view.asp?type="&pro_type&"&amp;title="&Replace(relatedpros_title," ","%20")&"&brand="&pro_brand&"' title='"&relatedpros_title&"'>"&relatedpros_title&"</a></h4></li>"
            relatedpros_rs.movenext()
          Next
          Easp.db.C(relatedpros_rs)
          %>

          <%If Not relatedpros_str = "" Then%>
          <div id="relatedpros">
            <h2>Related Products</h2>

            <ul>
              <%Response.Write(relatedpros_str)%>
            </ul>
            <a class="more" title="See More <%=name%>" href="/phone_case/<%=pro_type_rule%>.asp?brand=<%=pro_brand%>"> more</a>
          </div>
          <%End If%>
        </div>

        <!--#include virtual="/inc/hot_products.asp"-->
      </div>

      <!--#include virtual="/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/inc/footer.asp"-->
</body>
</html>