<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="x-ua-compatible" content="ie=7" />
<!--#include file="connections/conn_szweben.asp"-->
<title>CEO-Correo - LIEKO.COM,<%=kw1%></title>
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
    <!--#include file="top_sp.asp"-->

    <div id="content">
      <div id="main">
        <h2 id="position"><a title="lieko.com" href="/index_sp.asp">Inicio</a> <span>/</span> <strong>CEO-Correo</strong></h2>

        <div id="pros" class="ceomail">
          <h1>CEO-Correo</h1>

          <form action="public/ceomail_send.asp" method="post">
            <fieldset id="intro">
              <legend>Respecto al cliente</legend>
              <p><strong>Queridos respectados amigos,</strong></p>
              <p class="text">Estamos profundamente agradecidos de que usted ha sido fuerte apoyo de Lieko. Para un mejor servicio a los clientes honrados y amigos, en búsqueda constante del sistema de un mejor servicio, espero que me dé sus comentarios sobre nuestras deficiencias. Por favor completa el siguiente formulario. y envíe usted correo electrónico a <a class="mail" href="mailto:CEO@lieko.com">CEO@lieko.com</a>, voy a responder a usted dentro de 24 horas.</p>
              <p>Saludos Crodiales</p>
              <p><em>—— LIEKO CEO <span>Mark Zhang</span></em></p>
            </fieldset>

            <fieldset id="mainContent">
              <legend>Su Queja</legend>
              <div>
                <label for="complaints">La Queja: <span class="required">(requerido)</span></label>
                <textarea id="complaints" name="complaints">Me quiero quejar con</textarea>
              </div>
            </fieldset>

            <fieldset id="info">
              <legend>Su Contacto Información</legend>
              <p>Para entrar en contacto con usted lo antes posible, por favor llene su información de contacto</p>
              <ul>
                <li>
                  <label for="name">Nombre:</label>
                  <input type="text" id="name" name="name" />
                </li>

                <li>
                  <label for="company">Compañía:</label>
                  <input type="text" id="company" name="company" />
                </li>

                <li>
                  <label for="phone">Celular:</label>
                  <input type="text" id="phone" name="phone" />
                </li>

                <li>
                  <label for="email">E-mail:</label>
                  <input type="text" id="email" name="email" />
                </li>

                <li>
                  <label for="city">Ciudad:</label>
                  <input type="text" id="city" name="city" />
                </li>

                <li>
                  <label for="country">País:</label>
                  <input type="text" id="country" name="country" />
                </li>
              </ul>
            </fieldset>

            <div class="btn">
              <input type="submit" id="send" name="send" value="Enviar" />
              <input type="reset" id="reset" name="reset" value="Borrar" />
            </div>
          </form>
        </div>
      </div>

      
      <!--#include file="left_sp.asp"-->
    </div>
  </div>

	<!--#include file="end_sp.asp"-->
</body>
</html>
