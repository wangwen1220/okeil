<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="x-ua-compatible" content="ie=7" />
<!--#include file="connections/conn_szweben.asp"-->
<title>Servicio al miembro VIP - LIEKO.COM,<%=kw1%></title>
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
    <!--#include file="top_sp.asp"-->


    <div id="content">
      <div id="main">
        <h2 id="position"><a title="lieko.com" href="/index_sp.asp">Inicio</a> <span>/</span> <strong>VIP Servicio</strong></h2>

        <div id="pros">
          <h1>Servicio al miembro VIP</h1>

          <p id="vip_intro"><em>Queridos Clientes,</em><br />Bienvenidos a visitor al sitio web de Lieko compañía!Le responderemos dentro de 8 horas después de recibir su consulta. Usted llegará a ser nuestro regular mienbro después de registrarse en nuestra página web. Así usted puede tomar las muestras gratuitas y descuentos de nuestros productos que estén en stock. Si usted se convertirá en nuestra vip, entonces aquí hay mucha solvencia y artículos privilegio esperando! A nuestros representante de ventas les encantanrian de hablar con usted sobre los detalles!</p>

          <form id="vip_form" action="public/vip_reg.asp" method="post">
            <fieldset>
              <legend>Entre su información de investigación</legend>

              <div>
                <label for="inquiry">Su consulta:</label>
                <textarea name="inquiry" id="inquiry"></textarea>
              </div>
            </fieldset>

            <fieldset>
              <legend>Escríba su país y Compañía Nombre</legend>

              <div>
                <label for="country">País:</label>
                <input name="country" id="country" type="text" /><em class="required">*</em>
              </div>

              <div>
                <label for="company">Compañía nombre:</label>
                <input name="company" id="company" type="text" /><em class="required">*</em>
              </div>

              <div>
                <h3>Tipo de Compañía:</h3>
                <label for="com_tpye" class="radio"><input name="com_tpye" id="com_tpye" type="radio" value="Trader" />Comerciante</label>
                <label for="com_tpye" class="radio"><input name="com_tpye" id="com_tpye" type="radio" value="Refurbishing Factory" />Restauración de Fábrica</label>
                <label for="com_tpye" class="radio"><input name="com_tpye" id="com_tpye" type="radio" value="Wholesaler" />Mayorista</label>
              </div>
            </fieldset>

            <fieldset>
              <legend>Escríba su Contacto Información</legend>
              <div>
                <label for="name">Nombre Completo:</label>
                <input name="name" id="name" type="text" /><em class="required">*</em>
              </div>

              <div>
                <label for="email">Correo:</label>
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
                <label for="address">Dirección:</label>
                <input name="address" id="address" type="text" /><span>(para el transporte)</span>
              </div>
            </fieldset>

            <fieldset>
              <legend>Escríba su otra Contacto Información</legend>
              <div>
                <label for="msn">MSN:</label>
                <input name="msn" id="msn" type="text" />
              </div>

              <div>
                <label for="skype">Skype:</label>
                <input name="skype" id="skype" type="text" />
              </div>

              <div>
                <label for="website">Sitio Web:</label>
                <input name="website" id="website" type="text" />
              </div>
            </fieldset>

            <fieldset class="btn">
              <input class="submit" name="submit" type="submit" value="Presentar" />
            </fieldset>
          </form>
        </div>
      </div>

      
      <!--#include file="left_sp.asp"-->
    </div>
  </div>

	<!--#include file="end_sp.asp"-->
</body>
</html>
