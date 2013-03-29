<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include file="conn.asp"-->
<%
dim cls
if session("admin")="" then
	Response.Redirect("index.asp?err=4")	
else
	if session("flag")<>"0" then
	
		cls = Instr(session("flag"), "prod")
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

<%
	if request("add") ="ok" then
	
		session("photo1") = request("photo1")
		session("prodid") = request("prodid")
		session("model") = request("model")
		session("content1") = request("content1")
		session("SearchType") = request("SearchType")
		session("prodname") = request("prodname")
		session("ImgFull") = request("ImgFull")
	
	'判断是不是最底层分类
	set rs = conn.execute("select * from pclass where ParentID="&Request("ParentID")&"")
	If not rs.eof then
		response.Write "<script language=JavaScript>{window.alert('所属模块必须选择最底层目录分类!');window.history.go(-1);}</script>"
		response.end
	End If
	rs.close
	set rs = nothing

	if request("ProdId")="" then
	response.Write "<script language=JavaScript>{window.alert('产品名称不得为空!');window.history.go(-1);}</script>"
	response.end
	end if

	set rs = conn.execute("select * from prodmain where prodid='"&Request("prodid")&"'")
	if not rs.eof then 
	response.Write "<script language=JavaScript>{window.alert('产品名称不能重复，请重新填写!');window.history.go(-1);}</script>"
	response.End()
	end if
	rs.close
	set rs = nothing
	

'查找该分类ID的类别名称
set rs = conn.execute("select * from pclass where id="&Request("ParentID")&"")
classname = rs("classname")
classname_sp = rs("classname_sp")
rs.close

'产生随即数
'dim num1
'dim rndnum
'Randomize
'for i=0 to 6
'num1=CStr(Chr((57-48)*rnd+48))
'rndnum=rndnum&num1
'next
'日期，格式：YYYYMMDD
'yy=year(date)
'mm=right("00"&month(date),2)
'dd=right("00"&day(date),2)
'riqi=yy & mm & dd
'生成元素,格式为：小时，分钟，秒
'xiaoshi=right("00"&hour(time),2)
'fenzhong=right("00"&minute(time),2)
'miao=right("00"&second(time),2)
'shijian=xiaoshi & fenzhong & miao
'生成订单号
'inBillNo=riqi & shijian & rndnum

'-------------------------------将此产品的大类和中类和小类也查找出来一并写入数据库只能用于三级分类开始-------------------------------
set rsclass1=server.createobject("adodb.recordset")
sqlclass1="select * from PClass where id="&request("parentid")&" and gb='"&gb&"' "
rsclass1.open sqlclass1,conn,1,3
	set rsclass2=server.CreateObject("adodb.recordset")
	sqlclass2="select * from PClass where id="&rsclass1("parentID")&" and gb='"&gb&"' "
	rsclass2.open sqlclass2,conn,1,3
		if rsclass2.eof and rsclass2.bof then 
			midcode=rsclass1("classname")
			midcode_sp=rsclass1("classname_sp")
			midseq=rsclass1("id")
			larcode=rsclass1("classname")
			larcode_sp=rsclass1("classname_sp")
			larseq=rsclass1("id")
		else
			set rsclass3=server.CreateObject("adodb.recordset")
			sqlclass3="select * from PClass where id="&rsclass2("parentID")&" and gb='"&gb&"' "
			rsclass3.open sqlclass3,conn,1,3
				if rsclass3.eof and rsclass3.bof then 
					midcode=rsclass2("classname")
					midcode_sp=rsclass2("classname_sp")
					midseq=rsclass2("id")
					larcode=rsclass2("classname")
					larcode_sp=rsclass2("classname_sp")
					larseq=rsclass2("id")
				else
					midcode=rsclass2("classname")
					midcode_sp=rsclass2("classname_sp")
					midseq=rsclass2("id")
					larcode=rsclass3("classname")
					larcode_sp=rsclass3("classname_sp")
					larseq=rsclass3("id")
				end if 
			rsclass3.close
			set rsclass3 = nothing
		end if 
		rsclass2.close
		set rsclass2 = nothing
	rsclass1.close
	set rsclass1 = nothing
	
'-----------------------------------查找大类和中类结束----------------------------------------------------------

set rs=server.createobject("adodb.recordset")
sql="select * from ProdMain" 
rs.open sql,conn,1,3

		rs.Addnew
		'rs("Id")=inBillNo
		rs("ProdId")=Trim(request.form("ProdId"))
		if Trim(request.form("ProdId_sp"))<>"" then
		rs("ProdId_sp")=Trim(request.form("ProdId_sp"))
		else
		rs("ProdId_sp")=Trim(request.form("ProdId"))
		end if
		rs("ProdName")=Trim(request.form("ProdName"))
		rs("Model")=Trim(request.form("Model"))
		if Trim(request.form("Model_sp"))<>"" then
		rs("Model_sp")=Trim(request.form("Model_sp"))
		else
		rs("Model_sp")=Trim(request.form("Model"))
		end if
		rs("gb")=request.form("gb")
		rs("gbCode")=classname	
		rs("gbCode_sp")=classname_sp		
		rs("gbseq")=request.form("ParentId")
		rs("LarCode")=larcode
		rs("LarCode_sp")=larcode_sp	
		rs("MidCode")=midcode
		rs("MidCode_sp")=midcode_sp
		rs("larseq")=LarSeq
		rs("midseq")=MidSeq
		rs("MemoSpec")=request.form("content1")
		rs("xuhao")=Trim(request.form("xuhao"))
		rs("SearchType")=trim(request.Form("SearchType"))
		rs("ImgFull")=trim(request.Form("ImgFull"))
		rs("photo1")=trim(request.Form("photo1"))
		rs("FileOther")=0
		rs("online")=true
		
		rs.update	
		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
		
		session("photo1") = ""
		session("prodid") = ""
		session("model") = ""
		session("content1") = ""
		session("SearchType")=""
		session("prodname")=""
		session("ImgFull")=""
		
'		Response.Redirect"admin_prod.asp?action=detail&id="&request.form("ProdId")
		response.Write "<script language=JavaScript>{window.alert('提交成功！');window.location.href='prodadd.asp?gb="&gb&"';}</script>"	
		response.end
	else
	response.write " 操作产品列表生成错误！！"	
	end if
%>