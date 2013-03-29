<!doctype html><%@codepage="65001"%><!--#include file="../../lib/easp/easp.asp"-->
<html>
<head>
  <title>Miembro VIP Registro resultado</title>
  <meta http-equiv="conten-type" content="text/html; charset=utf-8" />
  <script type="text/javascript">
    window.onload=refresh;
    var i=5;
    function refresh(){
      document.getElementById("refresh").innerHTML="After <strong>"+i+"</strong> seconds will automatically return to the previous page.";
      i-->0?setTimeout("refresh();",1000):document.location.href="http://www.lieko.com/spanish/";//history.go(-1);//window.history.go(-1);
    }
  </script>

  <style type="text/css" media="screen">
    @import "/css/mail.css";
  </style>
  <link rel="Shortcut Icon" type="image/x-icon" href="/favicon.ico" />
</head>

<body id="vip_reg">
  <h1>Gracias por su registion</h1>

  <%
  country=Easp.CheckForm(Easp.Post("country:s"),"",1,"Your country field is empty")
  company=Easp.CheckForm(Easp.Post("company:s"),"",1,"Your company name field is empty")
  com_tpye=Request.form("com_tpye")
  customer=Easp.CheckForm(Easp.Post("name:s"),"",1,"Your full name field is empty")
  email=Easp.CheckForm(Easp.Post("email:s"),"email",1,"Your email address field is empty:Your email address is invalid")
  tel=Easp.CheckForm(Easp.Post("tel:s"),"",1,"Your tel field is empty")
  fax=Easp.CheckForm(Easp.Post("fax:s"),"",1,"Your fax field is empty")
  address=Request.form("address")
  msn=Request.form("msn")
  skype=Request.form("skype")
  website=Request.form("website")
  inquiry=Request.form("inquiry")

  if trim(Request.form("submit"))="Presentar" then
    msg = msg & "VIP Member Registration Information" & vbcrlf
    msg = msg & "--------------------------------------------------" & vbcrlf
    msg = msg & "Country: " & country & vbcrlf
    msg = msg & "Company: " & company & vbcrlf
    msg = msg & "Company Tpye: " & com_tpye & vbcrlf & vbcrlf
    msg = msg & "Customer: " & customer & vbcrlf
    msg = msg & "Email: " & email & vbcrlf
    msg = msg & "Tel: " & tel & vbcrlf
    msg = msg & "Fax: " & fax & vbcrlf
    msg = msg & "Address: " & address & vbcrlf & vbcrlf
    msg = msg & "MSN: " & msn & vbcrlf
    msg = msg & "Skype: " & skype & vbcrlf
    msg = msg & "Website:" & website & vbcrlf
    msg = msg & "--------------------------------------------------" & vbcrlf
    msg = msg & inquiry & vbcrlf

    sendUrl = "http://schemas.microsoft.com/cdo/configuration/sendusing"
    smtpUrl = "http://schemas.microsoft.com/cdo/configuration/smtpserver"

    On Error Resume Next
    'Set the mail server configuration
    Set cdoConfig = CreateObject("CDO.Configuration")
    cdoConfig.Fields.Item(sendUrl)=2 'cdoSendUsingPort
    cdoConfig.Fields.Item(smtpUrl)="relay-hosting.secureserver.net"
    cdoConfig.Fields.Update

    'Create and send the mail
    Set cdoMail=CreateObject("CDO.Message")
    'Use the config object created above
    Set cdoMail.Configuration=cdoConfig
    cdoMail.From = "lieko@godaddy.com"
    cdoMail.To = "ceo@lieko.com"
    cdoMail.Bcc = "lieko@lieko.com"
    cdoMail.BodyPart.Charset = "utf-8"
    cdoMail.Subject = "LIEKO VIP Member Registration"
    cdoMail.TextBody = msg
    cdoMail.Send
    Set cdoConfig = Nothing
    Set cdoMail = Nothing

    if err then
      err.clear
      result="La escritura es activada, pero el servidor no es compatible con ellos o dirección de correo electrónico en la carta no fue enviada Su mensaje se ha enviado,se lo responderémos a la brevedad."
    else
      result="Su mensaje se ha enviado, se lo responderémos a la brevedad."
    end if
  end if
  %>

  <p id="result"><%=result%></p>
  <p id="refresh"></p>
  <p>Haz clic aquí para regresar al <a href="http://www.lieko.com/spanish/">inicio</a>.</p>
</body>
</html>