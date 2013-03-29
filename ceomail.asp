<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="x-ua-compatible" content="ie=7" />
<!--#include file="connections/conn_szweben.asp"-->
<title>CEO-mail - LIEKO.COM,<%=kw1%></title>
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
<link href="css/ceomail.css" rel="stylesheet" type="text/css" />
<link rel="shortcut icon" type="image/x-icon" href="/favicon.ico">
</head>

<body id="ceomail">

  <div id="wrap">
    <!--#include file="top.asp"-->

    <div id="content">
      <div id="main">
        <h2 id="position"><a title="lieko.com" href="/index.asp">Home</a> <span>/</span> <strong>CEO-mail</strong></h2>

        <div id="pros" class="ceomail">
          <h1>CEO-mail</h1>

          <form action="public/ceomail_send.asp" method="post">
            <fieldset id="intro">
              <legend>respect to customers</legend>
              <p><strong>Dear respected friends,</strong></p>
              <p class="text">We are deeply grateful that you have been strong supporting LIEKO. To better service honored customers &amp; friends, In constant pursuit of better service system, hope you give me feedback about our deficiencies. Please complete the form below. and email it to <a class="mail" href="mailto:CEO@lieko.com">CEO@lieko.com</a>, I will reply to you within 24 hours.</p>
              <p>Best regards</p>
              <p><em>—— LIEKO CEO <span>Mark Zhang</span></em></p>
            </fieldset>

            <fieldset id="mainContent">
              <legend>your complaint</legend>
              <div>
                <label for="complaints">the complaint: <span class="required">(required)</span></label>
                <textarea id="complaints" name="complaints">I would like to complaint</textarea>
              </div>
            </fieldset>

            <fieldset id="info">
              <legend>your contact information</legend>
              <p>In order to <strong>contact you as soon as possible</strong>, please fill in your contact information.</p>
              <ul>
                <li>
                  <label for="name">name:</label>
                  <input type="text" id="name" name="name" />
                </li>

                <li>
                  <label for="company">company:</label>
                  <input type="text" id="company" name="company" />
                </li>

                <li>
                  <label for="phone">phone:</label>
                  <input type="text" id="phone" name="phone" />
                </li>

                <li>
                  <label for="email">E-mail:</label>
                  <input type="text" id="email" name="email" />
                </li>

                <li>
                  <label for="city">city:</label>
                  <input type="text" id="city" name="city" />
                </li>

                <li>
                  <label for="country">country:</label>
                  <input type="text" id="country" name="country" />
                </li>
              </ul>
            </fieldset>

            <div class="btn">
              <input type="submit" id="send" name="send" value="send" />
              <input type="reset" id="reset" name="reset" value="reset" />
            </div>
          </form>
        </div>
      </div>

      
      <!--#include file="left.asp"-->
    </div>
  </div>

	<!--#include file="end.asp"-->
</body>
</html>
