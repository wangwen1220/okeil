<!--#include file="../backgl/constCls.asp"-->
<!--#include file="../backgl/sqlget.asp"-->
<% 
Dim Fy_Url,Fy_a,Fy_x,Fy_Cs(),Fy_Cl,Fy_Ts,Fy_Zx 
'---定义部份 头------ 
Fy_Cl = 1 '处理方式：1=提示信息,2=转向页面,3=先提示再转向 
Fy_Zx = "index.Asp" '出错时转向的页面 
'---定义部份 尾------ 

On Error Resume Next 
Fy_Url=Request.ServerVariables("QUERY_STRING") 
Fy_a=split(Fy_Url,"&") 
redim Fy_Cs(ubound(Fy_a)) 
On Error Resume Next 
for Fy_x=0 to ubound(Fy_a) 
Fy_Cs(Fy_x) = left(Fy_a(Fy_x),instr(Fy_a(Fy_x),"=")-1) 
Next 
For Fy_x=0 to ubound(Fy_Cs) 
If Fy_Cs(Fy_x)<>"" Then 
If Instr(LCase(Request(Fy_Cs(Fy_x))),"'")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"select")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"update")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"chr")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"delete%20from")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),";")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"insert")<>0 or Instr(LCase(Request(Fy_Cs(Fy_x))),"mid")<>0 Or Instr(LCase(Request(Fy_Cs(Fy_x))),"master.")<>0 Then 
Select Case Fy_Cl 
Case "1" 
Response.Write "<Script Language=JavaScript>alert(' 出现错误！参数 "&Fy_Cs(Fy_x)&" 的值中包含非法字符串！\n\n 请不要在参数中出现：and,select,update,insert,delete,chr 等非法字符！\n\n我已经设置了不能SQL注入，请不要对我进行非法手段！');window.close();</Script>" 
Case "2" 
Response.Write "<Script Language=JavaScript>location.href='"&Fy_Zx&"'</Script>" 
Case "3" 
Response.Write "<Script Language=JavaScript>alert(' 出现错误！参数 "&Fy_Cs(Fy_x)&"的值中包含非法字符串！\n\n 请不要在参数中出现：,and,select,update,insert,delete,chr 等非法字符！\n\n设计了门，非法侵入请离开，谢谢！');location.href='"&Fy_Zx&"';</Script>" 
End Select 
Response.End 
End If 
End If 
Next 
%> 
<%
'生成对象
Dim Downasp
Set Downasp = New const_Cls
%>
<%
sqlUsername = "admin"
sqlPassword = "www.wxweb.cn"          

DB="backgl/data_@#$%DB/databs_ch.mdb"
set conn=server.createobject("adodb.Connection")
connstr="provider=Microsoft.Jet.OLEDB.4.0; User ID = " & SqlUsername & "; Jet OleDB:Database Password  = " & SqlPassword & "; Data Source=" & Server.MapPath(DB)
'如果服务器是老的ACCESS，请用下面的连接
'connstr= "driver={Microsoft Access Driver(*.mdb)};dbq=" & Server.MapPath(DB)
conn.Open connstr
%>

<%
dim gbbase(9)
	gbbase(0)="1"
%>

<%
'---------------设置网站默认关键字和标题栏---------------
set rs=server.createobject("adodb.recordset")
sql="select top 1 * from base_info where gb='"&gbbase(0)&"' order by id desc " 
rs.open sql,conn,1,3
	if not rs.eof  then
	kw1= rs("A_title")
	kw2= rs("A_keywords")
	kw3= rs("A_meta")
	kw4= rs("A_notes")
	kw5= rs("A_company")
	kw6= rs("A_tel")
	end if
rs.close
set rs = nothing

sqlprod="select top 1 * from banner where gb='"&gbbase(0)&"' order by id desc"
Set rsprod=Server.CreateObject("ADODB.RecordSet") 
rsprod.open sqlprod,conn,1,1
	if not rsprod.eof then
		gotourl1=rsprod("link1")
		gotourl2=rsprod("link2")	
		gotourl3=rsprod("link3")
		gotourl4=rsprod("link4")
		gotourl5=rsprod("link5")
		gotourl6=rsprod("link6")
		gotourl7=rsprod("link7")
		gotourl8=rsprod("link8")	
	end if 
rsprod.close
set rsprod = nothing

%>

<%
function checkuserstr(str)
   '检测用户注册名
	if len(str) = 0 or isnull(str) then
		checkuserstr = ""
		exit function
	end if
	dim i 
	dim lb_user,strfilter,regEx,str1,toustr,strfilter1
		
	strfilter="\$|\(|\)|\*|\+|\-|\.|\[|]|\?|\\|\^|\{|\||}|~|`|!|@|#|%|&|_|=|<|>|/|,|'| |　|\:"
	toustr="^ddn.*?$"  '检查开头字符

	Set regEx = New RegExp
	regEx.Pattern =	strfilter
	regEx.IgnoreCase = True
	lb_user = regex.Test(str)
	if not lb_user then
		regex.Pattern="法[ 　]*轮[ 　]*功"
		checkuserstr = regex.Replace(str,"")
	else
		checkuserstr = ""
	end if
	

	regEx.Pattern =	toustr
	regEx.IgnoreCase = True
	lb_user = regex.Test(str)
	if  lb_user then		
		checkuserstr = ""
	end if
	
	dim regEx2
	strfilter1 = "^[A-Za-z0-9\u4e00-\u9fa5]+$" '输入的数据必须是中文、英文和数字
	Set regEx2 = New RegExp
	regEx2.Pattern =	strfilter1
	regEx2.IgnoreCase = True
	
	if not regex2.Test(str) then		
		checkuserstr = ""
	end if
	set regex2 = nothing
		
	set regex=nothing	
end function 

function checkint(str,def)
   '检测输入的是否是整数
   'str 输入的字符串，def如果str非法则返回的整数
   str = trim(str)
   if len(str)= 0 or isnull(str) then
       checkint = def
       exit function
   end if
   if isnumeric(str) then
       checkint=clng(str)
   else
       checkint=def
   end if
end function

function checksqlstr(getstr)
'检测输入的参数是否含有sql敏感字符，如果有返回空字符串
	dim strfilter,strtmp,i,regEx
	if len(getstr) = 0 or isnull(getstr) then
		checksqlstr = ""
		exit function
	end if
	Set regEx = New RegExp
	strfilter = "select|delete|update|drop|create|exec"
	regEx.Pattern =	strfilter
	regEx.IgnoreCase = True
	regEx.Global = True
	getstr = trim(regex.Replace(getstr,""))
	strfilter="'"
	regEx.Pattern=   strfilter
	getstr = trim(regex.Replace(getstr,"''"))
	
	strfilter="0x"
	regEx.Pattern=   strfilter
	getstr = trim(regex.Replace(getstr,""))
	
	
	regex.Pattern="法[\s　]*轮[\s　]*功"
	getstr=regex.Replace(getstr,"*轮*")
	set regex=nothing
	checksqlstr = getstr
end function

sub jstop(strmsg)
'显示信息并回退一步
	dim html
	html = "<script>alert('" & strmsg & "');if(history.length==0){window.opener='';window.close();}else{history.go(-1);}</script>"
	response.Write html
	response.End 
end sub

sub showmsgbox(strmsg,strurl)
'显示信息
	dim html
	
	html = "<script>alert('" & strmsg & "');window.location.href='"&strurl&"'</script>"
	response.Write html
	
	response.End 
end sub


function nohtml(str) 
	dim re 
	Set re=new RegExp 
	re.IgnoreCase =true 
	re.Global=True 
	re.Pattern="(\<.[^\<]*\>)" 
	str=re.replace(str," ") 
	re.Pattern="(\<\/[^\<]*\>)" 
	str=re.replace(str," ") 
	nohtml=str 
	set re=nothing 
end function 

Function gb2utf8(Chinese) 
	if isnull(Chinese) then 
	gb2utf8="" 
	exit function 
	end if 
	
	For J = 1 to Len (Chinese) 
	a=Mid(Chinese, J, 1) 
	gb2utf8=gb2utf8 & "&#x" & Hex(Ascw(a)) & ";" 
	next 
End Function 

sub Jmail(email,subject,content)
	dim Fromer,sender,MailAddress
	Fromer = "webmaster@szfongda.com" 
	sender = "szfongda.com"
	MailAddress = "mail.szfongda.com"
	
	on error resume next
	Set myJmail = Server.CreateObject("JMAIL.SMTPMail") 
		myJmail.silent = true 
		myJmail.logging = true
		myJmail.Charset = "utf-8" 
		myJmail.ContentType = "text/html"
		'myJmail.ServerAddress = mailaddress 
		myJmail.AddRecipient  Email 
		myJmail.SenderName = sender 
		myJmail.Sender = fromer 
		'myJmail.MailServerUserName = "UserName of Email" 
		'myJmail.MailServerPassword = "Password of Email"
		myJmail.Priority = 1 
		myJmail.Subject = subject 
		myJmail.Body = content 
		'myJmail.AddRecipientBCC Email 
		'myJmail.AddRecipientCC Email 
		myJmail.Execute() 
		myJmail.Close  
end sub

%>