<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Server.scripttimeout=900%>
<!--#include file="downupload.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>文件上传</title>
<link type="text/css" href="css.css" rel="stylesheet">
</head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" marginwidth="0" marginheight="0" >
<table width="100%" border="1" cellspacing="10" cellpadding="5" bordercolor="#CCCCCC" height="100%">
  <tr> 
    <td bgcolor="#f7f7f7" align="center"> 
      <%
	dim upload,file,formName,formPath,iCount,filename,fileExt,fileSize
	set upload=new upload_5xSoft ''建立上传对象

	 formPath=upload.form("filepath")
	 '--------------------在目录后加(/)--------------------
	 if right(formPath,1)<>"/" then formPath=formPath&"/" 

	iCount=0
	for each formName in upload.file ''列出所有上传了的文件
	 set file=upload.file(formName)  ''生成一个文件对象
	 if file.filesize<100 then
		response.write "<font size=2>请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
		response.end
	 end if
		
	 if file.filesize>9000000 then
		response.write "<font size=2>文件大小超过了限制 9000K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
		response.end
	 end if
	
	 fileExt=lcase(right(file.filename,4))
	 if fileEXT=".exe" or fileEXT=".asp" or fileEXT=".php" or fileEXT=".jsp" or fileEXT=".dll" then
 	  response.write "<font size=2>系统不允许此文件格式上传　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	  response.end
	 end if 

	 fileSize=Int(file.filesize/1024)
	 If fileSize=0 then fileSize=1
 
 'response.Write filename
 'response.End()
 randomize
 ranNum=int(90000*rnd)+10000
 'filename = formPath&file.filename
 filename=formPath&year(now)&month(now)&day(now)&ranNum&fileExt
    'filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
 file.SaveAs Server.mappath(FileName)   ''保存文件
	'response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" 成功!<br>"
 weizhi=FileName 
 iCount=iCount+1
 end if
 
	set file=nothing
	next
	set upload=nothing  ''删除此对象
	response.write "<br>"
%>
      <div align="center">文件已经成功上传！<br>
        <br>
      </div>
        
      <table border="0" cellspacing="8" cellpadding="0" width="264">
        <form method="post" action="">
          <tr> 
            <td> <div align="center"> 
                <input type="hidden" value=<%=weizhi%> name="weizhi">
                <input type="button" name="Submit" value="关闭窗口" class="buttonface" onClick="javascript:CloseCopy();">
              </div></td>
          </tr>
        </form>
      </table>   
</td></tr></table>

<script language="javascript">
function CloseCopy(){
	window.opener.myform.pic2.value="<%=weizhi%>";
	window.opener.myform.down1.value="<%=weizhi%>";
	//window.opener.myform.down2.value="<%=weizhi%>";
	window.opener.myform.Size.value="<%=fileSize%>";
	window.close();
}
</script>
</body>
</html>