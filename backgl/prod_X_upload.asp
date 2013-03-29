<%@ codepage="65001"%>
<html>
<head>
<title>图片上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link type="text/css" href="css.css" rel="stylesheet">
</head>

<body>
<table width="100%" border="1" cellspacing="10" cellpadding="10" bordercolor="#CCCCCC" height="100%">
  <tr>
    <td bgcolor="#f7f7f7"> 
      <div align="center"><font color="#FF6600"><span class="14p">图片上传</span></font> 
        <br>
      </div>
      <form name="form" method="post" action="Prod_X_upfile.asp" enctype="multipart/form-data" >
        <div align="center">
          <p>
		  	<input type="hidden" name="txt" value="<%=request("txt")%>">
            <input type="hidden" name="act" value="upload">
            <input type="hidden" name="filepath" value="../Upfiles/Prod_X">
            <font color="#FF6600"></font></p>
          <p> 
            <input type="file" name="file1">
            <input type="submit" name="Submit" value="上传" onSubmit="javascript:document.form1.submit;" class="buttonface">
            <br>
            <br>
            <br>
            <font color="#FF0000">注：文件类型：gif/jpg，大小限制：250K<br>
            <br>
            图片尺寸请将宽度限制在500像素之内，否则会影响效果</font><br>
          </p>
        </div>
      </form>
    </td>
  </tr>
</table>
</body>
</html>
