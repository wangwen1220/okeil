<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<%
dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
else
	if session("flag")<>"0" then
	
		cls = Instr(session("flag"), "tem")
		if cls <= 0 then
	%>
		<script language="javascript">
			if (confirm("您的操作权限不够,系统拒绝你的访问,请点确定返回,或者点取消退出重新登录"))
			  location.href="index.asp?err=5";
			else
			  location.href="quit.asp";
		</script>
	<%
		end if
	end if
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>管理中心——>系统数据库管理</title>
<link type="text/css" href="css.css" rel="stylesheet">
</head>
<body onmouseover="self.status='欢迎使用网站后台管理系系统';return true"bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" height="100%" align="center">
  <tr>
    <td valign="top" width="100%"> 
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#222222" style="border-collapse: collapse">
        <tr> 
          <td width="100%" height="25">当前位置:网站管理中心--系统数据库管理</td>
        </tr>
        <tr> 
          <td width="100%"> <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#222222" width="100%">
              <tr> 
                <td> <%if request("back")="" then%> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
                    <tr> 
                      <td align="center" valign="top"> <table width="100%" border="0" height="100%" cellspacing="0" cellpadding="4">
                          <tr>
                            <td height="60">&nbsp;</td>
                          </tr>
                          <tr> 
                            <td height="25"> <div align="center"><strong><A HREF="admin_backup.asp">备份数据库</A></strong> 
                                |　<strong><A HREF="admin_backup.asp?back=h">还原数据库</A></strong></div>
                              <%
			if request("action")="Backup" then
			call backupdata()
			else
			%></td>
                          </tr>
                          <tr> 
                            <form method="post" action="admin_backup.asp?action=Backup">
                              <td> <table width="80%" border="1" align="center" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#cccccc">
                                  <tr> 
                                    <td height="25" background="images/backtop_2.jpg" class="toptext"><div align="center">备份系统数据文件[需要FSO权限]</div></td>
                                  </tr>
                                  <tr> 
                                    <td height="22"> 当前数据库路径</td>
                                  </tr>
                                  <tr> 
                                    <td height="22"><input type=text size=50 name=DBpath value="<%=db%>"></td>
                                  </tr>
                                  <tr> 
                                    <td height="22"> 备份数据库目录[如目录不存在，程序将自动创建]</td>
                                  </tr>
                                  <tr> 
                                    <td height="22"><input type=text size=50 name=bkfolder value=Databackup></td>
                                  </tr>
                                  <tr> 
                                    <td height="22">备份数据库名称[如备份目录有该文件，将覆盖，如没有，将自动创建]</td>
                                  </tr>
                                  <tr> 
                                    <td height="22"><input type=text size=30 name=bkDBname value=DataShop.mdb></td>
                                  </tr>
                                  <tr> 
                                    <td height="22"><div align="center"> 
                                        <input type=submit value="备份数据库">
                                      </div></td>
                                  </tr>
                                  <tr> 
                                    <td height="22"> 本程序的默认数据库文件为<%=db%><br>
                                      您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
                                      注意：所有路径都是相对与程序空间根目录的相对路径</td>
                                  </tr>
                                  <tr> 
                                    <td height="22">&nbsp;</td>
                                  </tr>
                                </table></td>
                            </form>
                          </tr>
                        </table>
                        <%end if%> <% 
sub backupdata() 
Dbpath=request.form("Dbpath") 
Dbpath=server.mappath(Dbpath) 
bkfolder=request.form("bkfolder") 
bkdbname=request.form("bkdbname")



Response.Write "<table width=100% height=200 border=0 cellpadding=0 style=border-collapse: collapse bordercolor=#000000><tr><td align=center>"
If not IsObjInstalled("Scripting.FileSystemObject") Then
	Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
Else

Set Fso=server.createobject("scripting.filesystemobject") 
if fso.fileexists(dbpath) then 
If CheckDir(bkfolder) = true Then 
fso.copyfile dbpath,bkfolder& "\"& bkdbname 
else 
MakeNewsDir bkfolder 
fso.copyfile dbpath,bkfolder& "\"& bkdbname 
end if 
response.write "备份数据库成功，您备份的数据库路径为" &bkfolder& "\"& bkdbname 
Response.write "<blockquote><br><a href='Admin_backup.asp'><b>点击返回</b></a><br></blockquote>"

Else 
response.write "找不到您所需要备份的文件。" 
End if
end if
Response.Write "</td></tr></table>"

end sub 
'------------------检查某一目录是否存在------------------- 
Function CheckDir(FolderPath) 
folderpath=Server.MapPath(".")&"\"&folderpath 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
If fso1.FolderExists(FolderPath) then 
'存在 
CheckDir = true 
Else 
'不存在 
CheckDir = False 
End if 
Set fso1 = nothing 
End Function 
'-------------根据指定名称生成目录--------- 
Function MakeNewsDir(foldername) 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
Set f = fso1.CreateFolder(foldername) 
MakeNewsDir = true 
Set fso1 = nothing 
End Function

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
%> </td>
                    </tr>
                  </table></td>
              </tr>
            </table>
            <%else%> <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
              <tr> 
                <td align="center" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="4">
                    <tr>
                      <td height="60">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="25"> <div align="center"><strong><A HREF="admin_backup.asp">备份数据库</A></strong>　
                          <strong><A HREF="admin_backup.asp?back=h">还原数据库</A></strong></div>
                        <%
if request("action")="huanyuan" then
call huanyuan()
else
%></td>
                    </tr>
                    <tr> 
                      <form method="post" action="admin_backup.asp?action=huanyuan&back=h" name="huanyuan">
                        <td> <table width="80%" border="1" align="center" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#cccccc">
                            <tr> 
                              <td height="25" background="images/backtop_2.jpg" class="toptext"><div align="center">还原系统数据文件[需要FSO权限]</div></td>
                            </tr>
                            <tr> 
                              <td height="22"> 备份数据库路径</td>
                            </tr>
                            <tr> 
                              <td height="22"><input type=text size=50 name=DBpath value="Databackup/DataShop.mdb"></td>
                            </tr>
                            <tr> 
                              <td height="22"> 还原数据库目录[如目录不存在，程序将自动创建]</td>
                            </tr>
                            <tr> 
                              <td height="22"><input type=text size=50 name=bkfolder value="data_@#$%DB"></td>
                            </tr>
                            <tr> 
                              <td height="22">还原数据库名称</td>
                            </tr>
                            <tr> 
                              <td height="22"><input type=text size=30 name=bkDBname value=databs_ch.mdb></td>
                            </tr>
                            <tr> 
                              <td height="22"><div align="center"> 
                                  <input type=submit value="还原数据库"onclick="{if(confirm('注意，该操作不可恢复！\n\n确定还原数据库吗？')){this.document.huanyuan.submit();return true;}return false;}">
                                </div></td>
                            </tr>
                            <tr> 
                              <td height="22"> 本程序的默认数据库文件为<%=db%><br>
                                您可以用这个功能来还原您的法规数据，以保证您的数据安全！<br>
                                注意：所有路径都是相对与程序空间根目录的相对路径</td>
                            </tr>
                            <tr> 
                              <td height="22">&nbsp;</td>
                            </tr>
                          </table></td>
                      </form>
                    </tr>
                  </table>
                  <%end if%> 
<% 
sub huanyuan() 
Dbpath=request.form("Dbpath") 
Dbpath=server.mappath(Dbpath) 
bkfolder=request.form("bkfolder") 
bkdbname=request.form("bkdbname")



Response.Write "<table width=100% height=200 border=0 cellpadding=0 style=border-collapse: collapse bordercolor=#000000><tr><td align=center>"
If not IsObjInstalled("Scripting.FileSystemObject") Then
	Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
Else

Set Fso=server.createobject("scripting.filesystemobject") 
if fso.fileexists(dbpath) then 
If CheckDir(bkfolder) = true Then 
fso.copyfile dbpath,bkfolder& "\"& bkdbname 
else 
MakeNewsDir bkfolder 
fso.copyfile dbpath,bkfolder& "\"& bkdbname 
end if 
response.write "还原数据库成功，您还原的数据库路径为" &bkfolder& "\"& bkdbname 
Response.write "<blockquote><br><a href='Admin_backup.asp'><b>点击返回</b></a><br></blockquote>"

Else 
response.write "找不到您所需要备份的文件。" 
End if
end if
Response.Write "</td></tr></table>"

end sub 
'------------------检查某一目录是否存在------------------- 
Function CheckDir(FolderPath) 
folderpath=Server.MapPath(".")&"\"&folderpath 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
If fso1.FolderExists(FolderPath) then 
'存在 
CheckDir = true 
Else 
'不存在 
CheckDir = False 
End if 
Set fso1 = nothing 
End Function 
'-------------根据指定名称生成目录--------- 
Function MakeNewsDir(foldername) 
Set fso1 = CreateObject("Scripting.FileSystemObject") 
Set f = fso1.CreateFolder(foldername) 
MakeNewsDir = true 
Set fso1 = nothing 
End Function


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
%> </td>
              </tr>
            </table></td>
        </tr>
      </table>
      <%end if%> </td>
  </tr>
</table>   
</td></tr></table>   
</td></tr></table>
</body>
</html>
