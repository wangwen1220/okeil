<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="x-ua-compatible" content="ie=7" />
<!--#include file="connections/conn_szweben.asp"-->
<title>VIP Member Registration - LIEKO.COM,<%=kw1%></title>
<meta name="keywords" content="<%=kw2%>">
<meta name="description" content="<%=kw3%>">
<meta name="author" content="<%=kw4%>" >
<meta name="copyright" content="<%=kw5%>">
<meta name="Abstrat" content="<%=kw6%>">
<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript" src="js/jquery-1.6.4.min.js"></script>
<script type="text/javascript" src="js/ready.min.js"></script>
<link href="css/base_min.css" rel="stylesheet" type="text/css" />
<link href="css/main.css" rel="stylesheet" type="text/css" />
<link rel="shortcut icon" type="image/x-icon" href="/favicon.ico">
</head>

<body id="vip">

  <div id="wrap">
    <!--#include file="top.asp"-->


    <div id="content">
      <div id="main">
        <h2 id="position"><a title="lieko.com" href="/index.asp">Home</a> <span>/</span> <strong>VIP Service</strong></h2>

        <div id="pros">
          <h1>VIP Member Service</h1>

          <p id="vip_intro"><em>Dear customers,</em><br />Welcome you to visit LIEKO company website! We will reply you within 8 hours after we get your inquiry. You will become our regular members after you register in our website. Then you can take free samples and discount of our products in stock. If you will become our VIPI, then here are much creditworthiness and privilege items waiting for you! Our sales are very gald to talk about the details with you!</p>

          <form id="vip_form" action="public/vip_reg.asp" method="post">
            <fieldset>
              <legend>Enter your Inquiry Information</legend>

              <div>
                <label for="inquiry">Your Inquiry:</label>
                <textarea name="inquiry" id="inquiry"></textarea>
              </div>
            </fieldset>

            <fieldset>
              <legend>Enter your Country & Company Name</legend>

              <div>
                <label for="country">Country:</label>
                <input name="country" id="country" type="text" /><em class="required">*</em>
              </div>

              <div>
                <label for="company">Company Name:</label>
                <input name="company" id="company" type="text" /><em class="required">*</em>
              </div>

              <div>
                <h3>Company Tpye:</h3>
                <label for="com_tpye" class="radio"><input name="com_tpye" id="com_tpye" type="radio" value="Trader" />Trader</label>
                <label for="com_tpye" class="radio"><input name="com_tpye" id="com_tpye" type="radio" value="Refurbishing Factory" />Refurbishing Factory</label>
                <label for="com_tpye" class="radio"><input name="com_tpye" id="com_tpye" type="radio" value="Wholesaler" />Wholesaler</label>
              </div>
            </fieldset>

            <fieldset>
              <legend>Enter your Address & Contact Information</legend>
              <div>
                <label for="name">Your Name:</label>
                <input name="name" id="name" type="text" /><em class="required">*</em>
              </div>

              <div>
                <label for="email">Email:</label>
                <input name="email" id="email" type="text" /><em class="required">*</em>
              </div>

              <div>
                <label for="tel">Tel:</label>
                <input name="tel" id="tel" type="text" /><em class="required">*</em>
              </div>

              <div>
                <label for="fax">Fax:</label>
                <input name="fax" id="fax" type="text" /><em class="required">*</em>
              </div>

              <div>
                <label for="address">Address:</label>
                <input name="address" id="address" type="text" /><span>(for shipping)</span>
              </div>
            </fieldset>

            <fieldset>
              <legend>Enter your other Contact Information</legend>
              <div>
                <label for="msn">MSN:</label>
                <input name="msn" id="msn" type="text" />
              </div>

              <div>
                <label for="skype">Skype:</label>
                <input name="skype" id="skype" type="text" />
              </div>

              <div>
                <label for="website">Website:</label>
                <input name="website" id="website" type="text" />
              </div>
            </fieldset>

            <fieldset class="btn">
              <input class="submit" name="submit" type="submit" value="Submit" />
            </fieldset>
          </form>
        </div>
      </div>

      
      <!--#include file="left.asp"-->
    </div>
  </div>

	<!--#include file="end.asp"-->
</body>
</html>
