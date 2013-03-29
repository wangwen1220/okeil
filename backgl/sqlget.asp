<%
'--------定义部份------------------
Dim Fy_Get,Fy_In,Fy_Inf,Fy_Xh,Fy_db,Fy_dbstr
'自定义需要过滤的字串,用 "郭" 分隔
Fy_In = "'郭;郭exec郭insert郭select郭delete郭update郭count郭*郭chr郭master郭truncate郭char郭declare郭{郭}"
'----------------------------------
%>

<%
Fy_Inf = split(Fy_In,"郭")

'--------GET部份-------------------
If Request.QueryString<>"" Then
For Each Fy_Get In Request.QueryString

For Fy_Xh=0 To Ubound(Fy_Inf)
If Instr(LCase(Request.QueryString(Fy_Get)),Fy_Inf(Fy_Xh))<>0 Then
'--------写入数据库----------头-----
'Fy_dbstr="DBQ="+server.mappath("/SqlIn.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
'Set Fy_db=Server.CreateObject("ADODB.CONNECTION")
'Fy_db.open Fy_dbstr
'Fy_db.Execute("insert into SqlIn(Sqlin_IP,SqlIn_Web,SqlIn_FS,SqlIn_CS,SqlIn_SJ) values('"&Request.ServerVariables("REMOTE_ADDR")&"','"&Request.ServerVariables("URL")&"','GET','"&Fy_Get&"','"&replace(Request.QueryString(Fy_Get),"'","''")&"')")
'Fy_db.close
'Set Fy_db = Nothing
'--------写入数据库----------尾-----

Response.Write "<Script Language=JavaScript>alert('系统友情提示!\n请不要在参数中包含非法字符尝试注入！\n\n');</Script>"
Response.Write "<b>非法操作！系统做了如下记录</b>↓<br>"
Response.Write "操作ＩＰ："&Request.ServerVariables("REMOTE_ADDR")&"<br>"
Response.Write "操作时间："&Now&"<br>"
Response.Write "操作页面："&Request.ServerVariables("URL")&"<br>"
Response.Write "提交方式：WEB GET<br>"
Response.Write "提交参数："&Fy_Get&"<br>"
Response.Write "提交数据："&Request.QueryString(Fy_Get)
Response.End
End If
Next
Next
End If
%>