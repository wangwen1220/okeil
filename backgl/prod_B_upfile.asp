﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="downupload.inc"-->
<%
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

 Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = true
	Set xTestObj = Nothing
	Err = 0
End Function

Const SaveUpFilesPath="Upfiles/prod_B"        '存放上传文件的目录
Const MaxPerPage=20
dim strFileName
dim totalPut,CurrentPage,TotalPages
dim UploadDir,truePath,fso,theFolder,theFile,whichfile,thisfile,FileCount,TotleSize
strFileName="admin_img.asp"

if right(SaveUpFilesPath,1)<>"/" then
	UploadDir="../" & SaveUpFilesPath & "/"
else
	UploadDir="../" & SaveUpFilesPath
end if

truePath=Server.MapPath(UploadDir)

	If not IsObjInstalled("Scripting.FileSystemObject") Then
		Response.Write "<br><br><br><b><font color=red><center>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</center></font></b>"
	Else
		set fso=CreateObject("Scripting.FileSystemObject")
		if request("photo")<>"" then
			whichfile=server.mappath(Request("photo")) 
			Set thisfile = fso.GetFile(whichfile) 
			thisfile.Delete true
		end if
		
		if request("photo2")<>"" then
			whichfile=server.mappath(Request("photo2")) 
			Set thisfile = fso.GetFile(whichfile) 
			thisfile.Delete true
		end if
		
		if request("photo3")<>"" then
			whichfile=server.mappath(Request("photo3")) 
			Set thisfile = fso.GetFile(whichfile) 
			thisfile.Delete true
		end if

		if request("photo4")<>"" then
			whichfile=server.mappath(Request("photo4")) 
			Set thisfile = fso.GetFile(whichfile) 
			thisfile.Delete true
		end if
		
		if request("photo5")<>"" then
			whichfile=server.mappath(Request("photo5")) 
			Set thisfile = fso.GetFile(whichfile) 
			thisfile.Delete true
		end if

		if request("photo6")<>"" then
			whichfile=server.mappath(Request("photo6")) 
			Set thisfile = fso.GetFile(whichfile) 
			thisfile.Delete true
		end if

		if request("photo7")<>"" then
			whichfile=server.mappath(Request("photo7")) 
			Set thisfile = fso.GetFile(whichfile) 
			thisfile.Delete true
		end if

		if request("photo8")<>"" then
			whichfile=server.mappath(Request("photo8")) 
			Set thisfile = fso.GetFile(whichfile) 
			thisfile.Delete true
		end if
		
	end if 
%>

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
dim upload,file,formName,formPath,iCount,filename,fileExt,filesize
set upload=new upload_5xSoft ''建立上传对象

 formPath=upload.form("filepath")
 ''在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 
 t=upload.form("t")
 gb=upload.form("gb")



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

 if fileEXT<>".jpg" and  fileEXT<>".swf" then
 	response.write "<font size=2>文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if 

 randomize
 ranNum=int(90000*rnd)+10000
 filename=formpath&"banner"&t&gb&fileEXT
 'filename=formpath&file.filename
 'filename=formPath&year(now)&month(now)&day(now)&ranNum&fileExt
 
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveAs Server.mappath(FileName)   ''保存文件
'  response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" 成功!<br>"
 	if fileEXT=".swf" then
 	weizhi=FileName
 	elseif fileEXT=".jpg" then
 	weizhi=FileName
 	end if
	filesize=file.filesize
 iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''删除此对象
response.write "<br>"

%>
      <div align="center">图片已经成功上传！<br>
        <br>
      </div>
        
      <table border="0" cellspacing="8" cellpadding="0" width="264">
        <form method="post" action="">
          <tr> 
            <td><div align="center"></div></td>
          </tr>
          <tr> 
            <td> <div align="center"> 
                <input type="hidden" value=<%=weizhi%> name="weizhi">
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
    var tt;
	tt="<%=t%>";
	if (tt=="1"){
	window.opener.prodtable.photo.value="<%=weizhi%>";
	}else if (tt=="2"){
	window.opener.prodtable.photo2.value="<%=weizhi%>";
	}else if (tt=="3"){
	window.opener.prodtable.photo3.value="<%=weizhi%>";
	}else if (tt=="4"){
	window.opener.prodtable.photo4.value="<%=weizhi%>";
	}else if (tt=="5"){
	window.opener.prodtable.photo5.value="<%=weizhi%>";
	}else if (tt=="6"){
	window.opener.prodtable.photo6.value="<%=weizhi%>";
	}else if (tt=="7"){
	window.opener.prodtable.photo7.value="<%=weizhi%>";
	}else if (tt=="8"){
	window.opener.prodtable.photo8.value="<%=weizhi%>";
	}

	window.close();
}
</script>
</body>
</html>