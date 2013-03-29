<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>软件上传</title>
<link type="text/css" href="css.css" rel="stylesheet">
</head>

<body text="#000000" leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<table width="100%" border="1" cellspacing="10" cellpadding="10" bordercolor="#CCCCCC" height="100%">
  <tr>
    <td bgcolor="#f7f7f7"> 
      <div align="center"><font color="#FF6600"><span class="14p">文件上传</span></font> 
        <br>
      </div>
      <form name="form" method="post" action="img_upfile.asp" enctype="multipart/form-data" >
        <div align="center">
          <p>
            <input type="hidden" name="act" value="upload">
            <input type="hidden" name="filepath" value="../Upfiles/down">
            <font color="#FF6600"></font></p>
          <p> 
            <input type="file" name="file1">
            <input type="submit" name="Submit" value="上传" onSubmit="javascript:document.form1.submit;" class="buttonface">
            <br>
            <br>
            <br>
            <br>
          </p>
        </div>
      </form>
    </td>
  </tr>
</table>
</body>
</html>
