<%
Dim style, title, keywords, description
style = "ceomail"
title = "CEO-Correo - LIEKO.COM"
keywords = "LIEKO CEO Correo"
description = "Lieko Co., Ltd CEO Correo."
%>
<!--#include virtual="/spanish/inc/head.asp"-->

<body id="ceomail">
  <div id="wrap">
    <!--#include virtual="/spanish/inc/header.asp"-->
    <div id="content">
      <div id="main">
        <h2 id="position"><a title="LIEKO Protectors para Telefonía Móvil" href="/spanish/index.asp">Inicio</a> <span>/</span> <strong>CEO-Correo</strong></h2>

        <div id="pros" class="ceomail">
          <h1>CEO-Correo</h1>

          <form action="public/ceomail_send.asp" method="post">
            <fieldset id="intro">
              <legend>Respecto al cliente</legend>
              <p><strong>Queridos amigos respetado,</strong></p>
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

      <!--#include virtual="/spanish/inc/sidebar.asp"-->
    </div>
  </div>

  <!--#include virtual="/spanish/inc/footer.asp"-->
</body>
</html>