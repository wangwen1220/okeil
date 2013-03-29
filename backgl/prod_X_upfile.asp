<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="downupload.inc"-->
<!--#include file="smallpic.asp"-->
<html>
<head>
<title>文件上传</title>
<link type="text/css" href="css.css" rel="stylesheet">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" marginwidth="0" marginheight="0" >
<table width="100%" border="1" cellspacing="10" cellpadding="5" bordercolor="#CCCCCC" height="100%">
  <tr> 
    <td bgcolor="#f7f7f7" align="center"> 
<%
dim upload,file,formName,formPath,iCount,filename,fileExt
dim txt
set upload=new upload_5xSoft ''建立上传对象

 formPath=upload.form("filepath")
 txt=upload.form("txt")
 ''在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 


iCount=0
for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.filesize<100 then
 	response.write "<font size=2>请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>250000 then
 	response.write "<font size=2>文件大小超过了限制 250K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" then
 	response.write "<font size=2>文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if 

 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&ranNum&fileExt
' filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName

 if file.FileSize>0 then     
 file.SaveAs Server.mappath(FileName)   


'---------------------获得到上传图片的长和宽-----------------------
'Dim Jpeg3
'Set Jpeg3 = Server.CreateObject("Persits.Jpeg")
'Jpeg3.Open server.mappath(FileName)
'ImgPrevWidth=Jpeg3.OriginalWidth
'ImgPrevHeight=Jpeg3.OriginalHeight
'Set Jpeg3=nothing
'---------------------获得上传图片的长宽---------------------------

'---------------------给图片加水印的新方法，图片会比较清楚开始---------------------------
'Dim photo,logo,photopath,logopath
'Set Photo = Server.CreateObject("Persits.Jpeg")
'PhotoPath = Server.MapPath(FileName)  'Path为路片路径及名称
'Photo.Open PhotoPath   '打开图片
'Set Logo = Server.CreateObject("Persits.Jpeg")
'LogoPath = Server.MapPath("02.gif")   '水印图片的路径
'Logo.Open LogoPath 
'Logo.Width = 445   '水印图片的宽度
'Logo.Height = Logo.Width * Logo.OriginalHeight / Logo.OriginalWidth
'Photo.Canvas.Pen.Color  = &H000000    '水印背景颜色
'Photo.Canvas.Pen.Width  = 1    
'Photo.Canvas.Brush.Solid = False  
'Photo.DrawImage (photo.width-445)/2, (photo.height-355)/2, Logo,0.05
'photo.Save Server.MapPath(FileName)  ''水印显示在图片上的XY位置
'Set logo = Nothing 
'Set photo = Nothing     
'---------------------给图片加水印的新方法，图片会比较清楚结束---------------------------

 	if fileEXT=".gif" then
 	weizhi=FileName
 	elseif fileEXT=".jpg" then
 	weizhi=FileName
 	end if
 iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''删除此对象

'--------------生成缩略图-----------------
'sSmallPath = BuildSmallPic(weizhi, "../Upfiles/prod_s", 118, 118)
'mSmallPath = BuildSmallPic(weizhi, "../Upfiles/prod_m", 400, 400)
%>
<div align="center">图片已经成功上传！<br><br></div>
        
      <table border="0" cellspacing="4" cellpadding="0" width="264">
        <form method="post" action="">
          <tr> 
            <td><div align="center"><a href="<%=weizhi%>" target="_blank"><img src="<%=weizhi%>" border="0"></a></div></td>
          </tr>
          <tr> 
            <td> <div align="center"> 
                <input type="hidden" value=<%=weizhi%> name="weizhi">
				<input type="hidden" value=<%=sSmallPath%> name="sSmallPath">
				<input type="hidden" value=<%=mSmallPath%> name="mSmallPath">
                <input type="button" name="Submit" value="关闭窗口" class="buttonface" onClick="javascript:CloseCopy();">
              </div></td>
          </tr>
        </form>
      </table>   
    </td>
  </tr>
</table>
<script language="javascript">
function CloseCopy(){
	var txt = "<%=txt%>";
	var obj = eval("window.opener.prodtable.photo"+txt);
	obj.value="<%=weizhi%>";
	//var obj1 = eval("window.opener.prodtable.sSmallPath"+txt);
	//obj1.value="<%=sSmallPath%>";
	//var obj2 = eval("window.opener.prodtable.mSmallPath"+txt);
	//obj2.value="<%=mSmallPath%>";
	//window.opener.prodtable.ImgPrevWidth.value="<%=ImgPrevWidth%>";
	//window.opener.prodtable.ImgPrevHeight.value="<%=ImgPrevHeight%>";
	window.close();
}
</script>
</body>
</html>