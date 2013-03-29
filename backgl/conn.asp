<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="constCls.asp"-->
<!--#include file="../connections/config.asp"-->
<%
'生成对象
Dim Downasp
Set Downasp = New const_Cls
%>

<%
sqlUsername = "admin"
sqlPassword = "www.wxweb.cn"          
DB="data_@#$%DB/databs_ch.mdb"
set conn=server.createobject("adodb.Connection")
connstr="provider=Microsoft.Jet.OLEDB.4.0; User ID = " & SqlUsername & "; Jet OleDB:Database Password  = " & SqlPassword & "; Data Source=" & Server.MapPath(DB)
'如果服务器是老的ACCESS，请用下面的连接
'connstr= "driver={Microsoft Access Driver(*.mdb)};dbq=" & Server.MapPath(DB)
conn.Open connstr

'---------------------------确定网站的版本要求-----------------------------
dim gb
 if request("gb")<>"" then 
 	gb=request("gb")
	session("gb")=gb
 end if 
 if gb="" then 
 	gb="1"
 end if 
 
dim gbbase(9)
	gbbase(0)="0"
	gbbase(1)="1"
'------------------------------------------------------------------------------------	 
%>

<%
Sub CloseConn()
	If IsObject(Downasp) Then
		Set Downasp = Nothing
	End If
	If IsObject(conn) Then
		conn.Close
		Set conn = Nothing
	End If
End Sub

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
%>