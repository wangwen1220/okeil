<!doctype html><%@codepage="65001"%>
<html>
<head>
<title>CEO-Correo - LIEKO</title>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="/css/ceomail.css" />
<script type="text/javascript">
window.onload=refresh;
var i=5;
function refresh(){
  document.getElementById("refresh").innerHTML="After <strong>"+i+"</strong> seconds will automatically return to the previous page.";
  i-->0?setTimeout("refresh();",1000):window.history.go(-1);//window.location.href="http://www.lieko.com/spanish/";
}
</script>
<link rel="shortcut icon" type="image/x-icon" href="/favicon.ico" />
</head>

<body id="send">
  <h1>Gracias por sus quejas</h1>

  <%
  customer=request.form("name")
  company=request.form("company")
  phone=request.form("phone")
  email=request.form("email")
  city=request.form("city")
  country=request.form("country")
  complaints=request.form("complaints")

  if Trim(Request.form("send")) = "Enviar" Then
    msg = msg & "Customer Information" & vbcrlf
    msg = msg & "--------------------------------------------------" & vbcrlf
    msg = msg & "Customer: " & customer & vbcrlf
    msg = msg & "Company: " & company & vbcrlf
    msg = msg & "Phone: " & phone & vbcrlf
    msg = msg & "Email: " & email & vbcrlf
    msg = msg & "City:" & city & vbcrlf
    msg = msg & "Country:" & country & vbcrlf & vbcrlf
    msg = msg & "Complaint Content" & vbcrlf
    msg = msg & "--------------------------------------------------" & vbcrlf
    msg = msg & " " & complaints & vbcrlf

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
    cdoMail.Subject = "LIEKO.COM Customer Complaints"
    cdoMail.TextBody = msg
    cdoMail.Send
    Set cdoMail = Nothing

    if err <> 0 then
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