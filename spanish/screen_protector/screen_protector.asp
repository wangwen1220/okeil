<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/lib/myasp.asp"-->
<%
Dim style, title, keywords, description, pro_type_rule, img_path, page_position, num_type, pro_type : pro_type = "Pantalla Protector"
num_type = "sp"
if Trim(Request.QueryString("type")) <> "" then
  pro_type = Trim(Request.QueryString("type"))
  pro_type_rule = Replace(Replace(Replace(LCase(pro_type),"pantalla protector","screen protector"),"privacidad","privacy"), " ", "_")
  img_path = "/img/products/screen_protector/"&pro_type_rule&"/"
  page_position = "<a title='lieko.com' href='/'>Home</a> <span>/</span> <a title='Pantalla Protector' href='screen_protector.asp'>Pantalla Protector</a> <span>/</span> <strong>"&pro_type&"</strong>"
else
  pro_type_rule = "screen_protector"
  img_path = "/img/products/screen_protector/privacy_screen_protector/"
  page_position = "<a title='lieko.com' href='/'>Home</a> <span>/</span> <strong>"&pro_type&"</strong>"
end if
style = "pros"
title = "iPad, iPhone, Blackberry, Nokia "&pro_type
keywords = pro_type
description = "Lieko iPad, iPhone, Blackberry, Nokia"&pro_type&"."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="<%=pro_type_rule%>" class="pros jqbox jqpager">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/spanish/inc/banner.asp"-->

        <h2 id="position"><%=page_position%></h2>

        <div id="pros" class="case jqbox">
          <h1 title="<%=pro_type%>"><%=pro_type%></h1>

          <p id="jqpager"> </p>
          <%
          dim img, pros_model, num, i, j, images, files : files = file_list(img_path)
          images = split(files(1), "|")
          i = UBound(images) + 1
          j = 0
          for each img in images
            pros_model = Replace(replace(replace(LCase(img), ".jpg", ""), "_", " "),"screen protector","Pantalla Protector")
            dim ilen : ilen = Len(i)
            if ilen < 3 then
              num = String(3 - ilen, "0") & i
            else
              num = i
            end if
            i = i - 1
            j = j + 1
          %>
          <div<%if j > 20 then echo " class='hide'"%>>
            <h3><%=pros_model%></h3>
            <a href="<%=img_path&img%>" title="<%=pros_model%>"><img src="<%=img_path&"thumb\"&replace(img, ".jpg", "_thumb.jpg")%>" alt='<%=pros_model%>' /></a>
            <span class="number"><%=num_type&num%></span>
          </div>
          <%next%>
        </div>

        <div class="types box">
          <h2><a href="iphone_screen_protector.asp" title="iPhone screen protector">iPhone Pantalla Protector</a></h2>
          <ul>
            <li><a href="screen_protector.asp" title="iPhone 4 Pantalla Protector">4</a></li>
            <li><a href="screen_protector.asp" title="iPhone 3GS Pantalla Protector">3GS</a></li>
            <li><a href="screen_protector.asp" title="iPhone 3G Pantalla Protector">3G</a></li>
          </ul>
        </div>

        <div class="types box">
          <h2>Blackberry Pantalla Protector</h2>
          <ul>
            <li><a href="screen_protector.asp" title="Blackberry 9800 Pantalla Protector">9800</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9700 Pantalla Protector">9700</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9650 Pantalla Protector">9650</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9630 Pantalla Protector">9630</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9600 Pantalla Protector">9600</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9550 Pantalla Protector">9550</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9530 Pantalla Protector">9530</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9520 Pantalla Protector">9520</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9500 Pantalla Protector">9500</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9300 Pantalla Protector">9300</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9100 Pantalla Protector">9100</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 9000 Pantalla Protector">9000</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8900 Pantalla Protector">8900</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8830 Pantalla Protector">8830</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8820 Pantalla Protector">8820</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8800 Pantalla Protector">8800</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8707 Pantalla Protector">8707</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8700 Pantalla Protector">8700</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8530 Pantalla Protector">8530</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8520 Pantalla Protector">8520</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8500 Pantalla Protector">8500</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8350i Pantalla Protector">8350i</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8330 Pantalla Protector">8330</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8320 Pantalla Protector">8320</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8310 Pantalla Protector">8310</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8300 Pantalla Protector">8300</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8200 Pantalla Protector">8200</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8130 Pantalla Protector">8130</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8120 Pantalla Protector">8120</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8110 Pantalla Protector">8110</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 8100 Pantalla Protector">8100</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7750 Pantalla Protector">7750</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7730 Pantalla Protector">7730</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7520 Pantalla Protector">7520</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7510 Pantalla Protector">7510</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7290 Pantalla Protector">7290</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7280 Pantalla Protector">7280</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7270 Pantalla Protector">7270</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7250 Pantalla Protector">7250</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7230 Pantalla Protector">7230</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7210 Pantalla Protector">7210</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7200 Pantalla Protector">7200</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 7100 Pantalla Protector">7100</a></li>
            <li><a href="screen_protector.asp" title="Blackberry 6210 Pantalla Protector">6210</a></li>
          </ul>
        </div>

        <div class="types box nokia">
          <h2>Nokia Pantalla Protector</h2>
          <ul>
            <li><a href="screen_protector.asp" title="Nokia 9500 Pantalla Protector">9500</a></li>
            <li><a href="screen_protector.asp" title="Nokia 8800 Pantalla Protector">8800</a></li>
            <li><a href="screen_protector.asp" title="Nokia 8310 Pantalla Protector">8310</a></li>
            <li><a href="screen_protector.asp" title="Nokia 7610 Pantalla Protector">7610</a></li>
            <li><a href="screen_protector.asp" title="Nokia 7280 Pantalla Protector">7280</a></li>
            <li><a href="screen_protector.asp" title="Nokia 7270 Pantalla Protector">7270</a></li>
            <li><a href="screen_protector.asp" title="Nokia 7260 Pantalla Protector">7260</a></li>
            <li><a href="screen_protector.asp" title="Nokia 7250 Pantalla Protector">7250</a></li>
            <li><a href="screen_protector.asp" title="Nokia 7210 Pantalla Protector">7210</a></li>
            <li><a href="screen_protector.asp" title="Nokia 7200 Pantalla Protector">7200</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6822 Pantalla Protector">6822</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6820 Pantalla Protector">6820</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6810 Pantalla Protector">6810</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6800 Pantalla Protector">6800</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6682 Pantalla Protector">6682</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6670 Pantalla Protector">6670</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6650 Pantalla Protector">6650</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6610 Pantalla Protector">6610</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6600 Pantalla Protector">6600</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6590 Pantalla Protector">6590</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6585 Pantalla Protector">6585</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6560 Pantalla Protector">6560</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6310 Pantalla Protector">6310</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6256i Pantalla Protector">6256i</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6255i Pantalla Protector">6255i</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6235i Pantalla Protector">6235i</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6230i Pantalla Protector">6230i</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6230 Pantalla Protector">6230</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6225 Pantalla Protector">6225</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6220 Pantalla Protector">6220</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6200 Pantalla Protector">6200</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6170 Pantalla Protector">6170</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6102 Pantalla Protector">6102</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6101 Pantalla Protector">6101</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6100 Pantalla Protector">6100</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6021 Pantalla Protector">6021</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6020 Pantalla Protector">6020</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6019i Pantalla Protector">6019i</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6015i Pantalla Protector">6015i</a></li>
            <li><a href="screen_protector.asp" title="Nokia 6010 Pantalla Protector">6010</a></li>
            <li><a href="screen_protector.asp" title="Nokia 5140 Pantalla Protector">5140</a></li>
            <li><a href="screen_protector.asp" title="Nokia 5100 Pantalla Protector">5100</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3660 Pantalla Protector">3660</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3650 Pantalla Protector">3650</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3600 Pantalla Protector">3600</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3595 Pantalla Protector">3595</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3590 Pantalla Protector">3590</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3586i Pantalla Protector">3586i</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3586 Pantalla Protector">3586</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3585 Pantalla Protector">3585</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3560 Pantalla Protector">3560</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3300 Pantalla Protector">3300</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3220 Pantalla Protector">3220</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3200 Pantalla Protector">3200</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3155i Pantalla Protector">3155i</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3155 Pantalla Protector">3155</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3120 Pantalla Protector">3120</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3108 Pantalla Protector">3108</a></li>
            <li><a href="screen_protector.asp" title="Nokia 3100 Pantalla Protector">3100</a></li>
            <li><a href="screen_protector.asp" title="Nokia 2651 Pantalla Protector">2651</a></li>
            <li><a href="screen_protector.asp" title="Nokia 2600 Pantalla Protector">2600</a></li>
            <li><a href="screen_protector.asp" title="Nokia 2285 Pantalla Protector">2285</a></li>
            <li><a href="screen_protector.asp" title="Nokia 2270 Pantalla Protector">2270</a></li>
            <li><a href="screen_protector.asp" title="Nokia 2260 Pantalla Protector">2260</a></li>
            <li><a href="screen_protector.asp" title="Nokia 2116i Pantalla Protector">2116i</a></li>
            <li><a href="screen_protector.asp" title="Nokia 2115 Pantalla Protector">2115</a></li>
            <li><a href="screen_protector.asp" title="Nokia 2100 Pantalla Protector">2100</a></li>
            <li><a href="screen_protector.asp" title="Nokia 1260 Pantalla Protector">1260</a></li>
            <li><a href="screen_protector.asp" title="Nokia 1221 Pantalla Protector">1221</a></li>
            <li><a href="screen_protector.asp" title="Nokia 1220 Pantalla Protector">1220</a></li>
            <li><a href="screen_protector.asp" title="Nokia 1100 Pantalla Protector">1100</a></li>
          </ul>
        </div>

        <!--#include virtual="/spanish/inc/hot_products.asp"-->
      </div>

      <!--#include virtual="/spanish/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/spanish/inc/footer.asp"-->
</body>
</html>