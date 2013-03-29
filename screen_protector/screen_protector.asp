<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%><!--#include virtual="/lib/myasp.asp"-->
<%
Dim style, title, keywords, description, pro_type_rule, img_path, page_position, num_type, pro_type : pro_type = "Screen Protector"
num_type = "sp"
if Trim(Request.QueryString("type")) <> "" then
  pro_type = Trim(Request.QueryString("type"))
  pro_type_rule = Replace(LCase(pro_type), " ", "_")
  img_path = "/img/products/screen_protector/"&pro_type_rule&"/"
  page_position = "<a title='lieko.com' href='/'>Home</a> <span>/</span> <a title='Screen Protector' href='screen_protector.asp'>Screen Protector</a> <span>/</span> <strong>"&pro_type&"</strong>"
else
  pro_type_rule = Replace(LCase(pro_type), " ", "_")
  img_path = "/img/products/screen_protector/privacy_screen_protector/"
  page_position = "<a title='lieko.com' href='/'>Home</a> <span>/</span> <strong>"&pro_type&"</strong>"
end if
style = "pros"
title = "iPad, iPhone, Blackberry, Nokia "&pro_type
keywords = pro_type&", Mobile Phone Cover"
description = "Lieko Co., Ltd. is the largest mobile phone cases supplier in Dongguan, the main products are iPad, iPhone, Blackberry, Nokia"&pro_type&" and Mobile Phone "&pro_type&"."
%>
<!--#include virtual="/inc/head.asp"-->

<body id="<%=pro_type_rule%>" class="pros jqbox jqpager">
  <div id="wrap">
    <!--#include virtual="/inc/header.asp"-->

    <div id="content">
      <div id="main">
        <!--#include virtual="/inc/banner.asp"-->

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
            pros_model = replace(replace(img, ".jpg", ""), "_", " ")
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
          <h2><a href="iphone_screen_protector.asp" title="iPhone screen protector">iPhone Screen Protector</a></h2>
          <ul>
            <li><a href="##" title="iPhone 4 Screen Protector">4</a></li>
            <li><a href="##" title="iPhone 3GS Screen Protector">3GS</a></li>
            <li><a href="##" title="iPhone 3G Screen Protector">3G</a></li>
          </ul>
        </div>

        <div class="types box">
          <h2>Blackberry Screen Protector</h2>
          <ul>
            <li><a href="##" title="Blackberry 9800 Screen Protector">9800</a></li>
            <li><a href="##" title="Blackberry 9700 Screen Protector">9700</a></li>
            <li><a href="##" title="Blackberry 9650 Screen Protector">9650</a></li>
            <li><a href="##" title="Blackberry 9630 Screen Protector">9630</a></li>
            <li><a href="##" title="Blackberry 9600 Screen Protector">9600</a></li>
            <li><a href="##" title="Blackberry 9550 Screen Protector">9550</a></li>
            <li><a href="##" title="Blackberry 9530 Screen Protector">9530</a></li>
            <li><a href="##" title="Blackberry 9520 Screen Protector">9520</a></li>
            <li><a href="##" title="Blackberry 9500 Screen Protector">9500</a></li>
            <li><a href="##" title="Blackberry 9300 Screen Protector">9300</a></li>
            <li><a href="##" title="Blackberry 9100 Screen Protector">9100</a></li>
            <li><a href="##" title="Blackberry 9000 Screen Protector">9000</a></li>
            <li><a href="##" title="Blackberry 8900 Screen Protector">8900</a></li>
            <li><a href="##" title="Blackberry 8830 Screen Protector">8830</a></li>
            <li><a href="##" title="Blackberry 8820 Screen Protector">8820</a></li>
            <li><a href="##" title="Blackberry 8800 Screen Protector">8800</a></li>
            <li><a href="##" title="Blackberry 8707 Screen Protector">8707</a></li>
            <li><a href="##" title="Blackberry 8700 Screen Protector">8700</a></li>
            <li><a href="##" title="Blackberry 8530 Screen Protector">8530</a></li>
            <li><a href="##" title="Blackberry 8520 Screen Protector">8520</a></li>
            <li><a href="##" title="Blackberry 8500 Screen Protector">8500</a></li>
            <li><a href="##" title="Blackberry 8350i Screen Protector">8350i</a></li>
            <li><a href="##" title="Blackberry 8330 Screen Protector">8330</a></li>
            <li><a href="##" title="Blackberry 8320 Screen Protector">8320</a></li>
            <li><a href="##" title="Blackberry 8310 Screen Protector">8310</a></li>
            <li><a href="##" title="Blackberry 8300 Screen Protector">8300</a></li>
            <li><a href="##" title="Blackberry 8200 Screen Protector">8200</a></li>
            <li><a href="##" title="Blackberry 8130 Screen Protector">8130</a></li>
            <li><a href="##" title="Blackberry 8120 Screen Protector">8120</a></li>
            <li><a href="##" title="Blackberry 8110 Screen Protector">8110</a></li>
            <li><a href="##" title="Blackberry 8100 Screen Protector">8100</a></li>
            <li><a href="##" title="Blackberry 7750 Screen Protector">7750</a></li>
            <li><a href="##" title="Blackberry 7730 Screen Protector">7730</a></li>
            <li><a href="##" title="Blackberry 7520 Screen Protector">7520</a></li>
            <li><a href="##" title="Blackberry 7510 Screen Protector">7510</a></li>
            <li><a href="##" title="Blackberry 7290 Screen Protector">7290</a></li>
            <li><a href="##" title="Blackberry 7280 Screen Protector">7280</a></li>
            <li><a href="##" title="Blackberry 7270 Screen Protector">7270</a></li>
            <li><a href="##" title="Blackberry 7250 Screen Protector">7250</a></li>
            <li><a href="##" title="Blackberry 7230 Screen Protector">7230</a></li>
            <li><a href="##" title="Blackberry 7210 Screen Protector">7210</a></li>
            <li><a href="##" title="Blackberry 7200 Screen Protector">7200</a></li>
            <li><a href="##" title="Blackberry 7100 Screen Protector">7100</a></li>
            <li><a href="##" title="Blackberry 6210 Screen Protector">6210</a></li>
          </ul>
        </div>

        <div class="types box nokia">
          <h2>Nokia Screen Protector</h2>
          <ul>
            <li><a href="##" title="Nokia 9500 Screen Protector">9500</a></li>
            <li><a href="##" title="Nokia 8800 Screen Protector">8800</a></li>
            <li><a href="##" title="Nokia 8310 Screen Protector">8310</a></li>
            <li><a href="##" title="Nokia 7610 Screen Protector">7610</a></li>
            <li><a href="##" title="Nokia 7280 Screen Protector">7280</a></li>
            <li><a href="##" title="Nokia 7270 Screen Protector">7270</a></li>
            <li><a href="##" title="Nokia 7260 Screen Protector">7260</a></li>
            <li><a href="##" title="Nokia 7250 Screen Protector">7250</a></li>
            <li><a href="##" title="Nokia 7210 Screen Protector">7210</a></li>
            <li><a href="##" title="Nokia 7200 Screen Protector">7200</a></li>
            <li><a href="##" title="Nokia 6822 Screen Protector">6822</a></li>
            <li><a href="##" title="Nokia 6820 Screen Protector">6820</a></li>
            <li><a href="##" title="Nokia 6810 Screen Protector">6810</a></li>
            <li><a href="##" title="Nokia 6800 Screen Protector">6800</a></li>
            <li><a href="##" title="Nokia 6682 Screen Protector">6682</a></li>
            <li><a href="##" title="Nokia 6670 Screen Protector">6670</a></li>
            <li><a href="##" title="Nokia 6650 Screen Protector">6650</a></li>
            <li><a href="##" title="Nokia 6610 Screen Protector">6610</a></li>
            <li><a href="##" title="Nokia 6600 Screen Protector">6600</a></li>
            <li><a href="##" title="Nokia 6590 Screen Protector">6590</a></li>
            <li><a href="##" title="Nokia 6585 Screen Protector">6585</a></li>
            <li><a href="##" title="Nokia 6560 Screen Protector">6560</a></li>
            <li><a href="##" title="Nokia 6310 Screen Protector">6310</a></li>
            <li><a href="##" title="Nokia 6256i Screen Protector">6256i</a></li>
            <li><a href="##" title="Nokia 6255i Screen Protector">6255i</a></li>
            <li><a href="##" title="Nokia 6235i Screen Protector">6235i</a></li>
            <li><a href="##" title="Nokia 6230i Screen Protector">6230i</a></li>
            <li><a href="##" title="Nokia 6230 Screen Protector">6230</a></li>
            <li><a href="##" title="Nokia 6225 Screen Protector">6225</a></li>
            <li><a href="##" title="Nokia 6220 Screen Protector">6220</a></li>
            <li><a href="##" title="Nokia 6200 Screen Protector">6200</a></li>
            <li><a href="##" title="Nokia 6170 Screen Protector">6170</a></li>
            <li><a href="##" title="Nokia 6102 Screen Protector">6102</a></li>
            <li><a href="##" title="Nokia 6101 Screen Protector">6101</a></li>
            <li><a href="##" title="Nokia 6100 Screen Protector">6100</a></li>
            <li><a href="##" title="Nokia 6021 Screen Protector">6021</a></li>
            <li><a href="##" title="Nokia 6020 Screen Protector">6020</a></li>
            <li><a href="##" title="Nokia 6019i Screen Protector">6019i</a></li>
            <li><a href="##" title="Nokia 6015i Screen Protector">6015i</a></li>
            <li><a href="##" title="Nokia 6010 Screen Protector">6010</a></li>
            <li><a href="##" title="Nokia 5140 Screen Protector">5140</a></li>
            <li><a href="##" title="Nokia 5100 Screen Protector">5100</a></li>
            <li><a href="##" title="Nokia 3660 Screen Protector">3660</a></li>
            <li><a href="##" title="Nokia 3650 Screen Protector">3650</a></li>
            <li><a href="##" title="Nokia 3600 Screen Protector">3600</a></li>
            <li><a href="##" title="Nokia 3595 Screen Protector">3595</a></li>
            <li><a href="##" title="Nokia 3590 Screen Protector">3590</a></li>
            <li><a href="##" title="Nokia 3586i Screen Protector">3586i</a></li>
            <li><a href="##" title="Nokia 3586 Screen Protector">3586</a></li>
            <li><a href="##" title="Nokia 3585 Screen Protector">3585</a></li>
            <li><a href="##" title="Nokia 3560 Screen Protector">3560</a></li>
            <li><a href="##" title="Nokia 3300 Screen Protector">3300</a></li>
            <li><a href="##" title="Nokia 3220 Screen Protector">3220</a></li>
            <li><a href="##" title="Nokia 3200 Screen Protector">3200</a></li>
            <li><a href="##" title="Nokia 3155i Screen Protector">3155i</a></li>
            <li><a href="##" title="Nokia 3155 Screen Protector">3155</a></li>
            <li><a href="##" title="Nokia 3120 Screen Protector">3120</a></li>
            <li><a href="##" title="Nokia 3108 Screen Protector">3108</a></li>
            <li><a href="##" title="Nokia 3100 Screen Protector">3100</a></li>
            <li><a href="##" title="Nokia 2651 Screen Protector">2651</a></li>
            <li><a href="##" title="Nokia 2600 Screen Protector">2600</a></li>
            <li><a href="##" title="Nokia 2285 Screen Protector">2285</a></li>
            <li><a href="##" title="Nokia 2270 Screen Protector">2270</a></li>
            <li><a href="##" title="Nokia 2260 Screen Protector">2260</a></li>
            <li><a href="##" title="Nokia 2116i Screen Protector">2116i</a></li>
            <li><a href="##" title="Nokia 2115 Screen Protector">2115</a></li>
            <li><a href="##" title="Nokia 2100 Screen Protector">2100</a></li>
            <li><a href="##" title="Nokia 1260 Screen Protector">1260</a></li>
            <li><a href="##" title="Nokia 1221 Screen Protector">1221</a></li>
            <li><a href="##" title="Nokia 1220 Screen Protector">1220</a></li>
            <li><a href="##" title="Nokia 1100 Screen Protector">1100</a></li>
          </ul>
        </div>

        <!--#include virtual="/inc/hot_products.asp"-->
      </div>

      <!--#include virtual="/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/inc/footer.asp"-->
</body>
</html>