<%
Class const_Cls
	Public LongTime, ShortTime, GetUserIP
	
	Private Sub Class_Initialize()
		LongTime = Now()
		ShortTime = Date()
		GetUserIP = Request.ServerVariables("REMOTE_ADDR")
		statusStr = "当前管理员："+Session("adminname")+"______IP："+GetUserIP
		'Response.Write("<script>var msg='"+statusStr+"';window.status=msg;")
		'以上语言如果需要，则在“msg;”后面加上/script与前面对应。
	End Sub
	'*************************************************************
	'函数作用：判断字符串是否空
	'*************************************************************
	Public Function IsEmp(Str)
		IsEmp = True
		Str = Trim(Str)
		If IsNull(Str) Then
			IsEmp = False
			Exit Function
		End If
		If Len(Str)<1 Then
			IsEmp = False
			Exit Function
		End If
		IsEmp = True
	End Function	
	'*************************************************************
	'函数作用：判断非法字符
	'*************************************************************
	Public Function IsValidStr(Str)
		IsValidStr = True
		If IsNull(Str) Then Exit Function
		Dim ForbidStr, i
		'ForbidStr = ":|=|%|&|$|#|@|+|-|*|/|\|<|>|;|,|"& Chr(32) &"|"& Chr(34) &"|"& Chr(39) &"|"& Chr(9)
		ForbidStr = ":|=|%|&|$|#|@|+|*|\|<|>|;|,|"& Chr(39) &"|"& Chr(9)
		ForbidStr = Split(ForbidStr, "|")
		For i = 0 To UBound(ForbidStr)
			If InStr(Str, ForbidStr(i)) > 0 Then
				IsValidStr = False
				Exit Function
			End If
		Next
		IsValidStr = True
	End Function
	'*************************************************************
	'函数作用：判断密码非法字符
	'*************************************************************
	Public Function IsValidPassword(Str)
		IsValidPassword = True
		If IsNull(Str) Then Exit Function
		Dim ForbidStr, i
		'ForbidStr = "=|%|&|;|,|"& Chr(32) &"|"& Chr(34) &"|"& Chr(39) &"|"& Chr(9)
		ForbidStr = "=|%|&|;|,|"& Chr(39) &"|"& Chr(9)
		ForbidStr = Split(ForbidStr, "|")
		For i = 0 To UBound(ForbidStr)
			If InStr(Str, ForbidStr(i)) > 0 Then
				IsValidPassword = False
				Exit Function
			End If
		Next
		IsValidPassword = True
	End Function
	'*************************************************************
	'函数作用：过滤SQL非法字符
	'*************************************************************
	Public Function checkStr(Str)
		If IsNull(Str) Then
			checkStr = ""
			Exit Function
		End If
		checkStr = Replace(Str, "'", "''")
	End Function
	'*************************************************************
	'函数作用：显示提示信息
	'*************************************************************
	Public Function ShowMsg(Str)
		Response.write("<script>alert('"+Str+"');</script>")
	End Function
	'*************************************************************
	'函数作用：跳转到某个页面
	'*************************************************************
	Public Function GotoUrl(Str)
		Response.write("<script>location.href('"+Str+"');</script>")
	End Function
	'*************************************************************
	'函数作用：返回前页
	'*************************************************************
	Public Function Back(ID)
		Response.write("<script>history.back();</script>")
	End Function
	'*************************************************************
	'函数作用：返回下载类别列表（无限级别）
	'*************************************************************
	Function getcatalogs(parentid,str)
		Dim Sql,rstemp,TemptempStr,tempStr
		if str="" then 
			tempStr=str&"<img src='images/N.png' align='absmiddle'>"
		else
			tempStr=str&"<img src='images/I.png' align='absmiddle'>"
		end if
		sql="select * from pclass where parentid="&parentid&" order by orders asc"
		Set rstemp = conn.Execute(Sql)
		if rstemp.eof then exit function
		do while Not rstemp.Eof
			Set rsP = conn.execute("select * from pclass where parentid="&rstemp("id"))
			if rsP.eof then
				getcatalogs = getcatalogs &"<tr bgcolor=f6f6f6 onmouseover=""this.style.backgroundColor='#C1DEFB'"" onmouseout=""this.style.backgroundColor=''""><td height=""25"" class=""underline"" align=""center"" width=""5%""><b>&nbsp;</b></td><td height=""25"" class=""underline"">"&tempStr&"<img src='images/T.png' align='absmiddle'><img src='images/file.png' align='absmiddle'>"&rstemp("orders")&"."&rstemp("classname")&"&nbsp;&nbsp;(<font color='#ff6600'>"&rstemp("classname_sp")&"</font>)</td><td height=""25"" class=""underline"" width=""30%"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=""admin_prodclass.asp?action=AddClass&ParentID="&rstemp("id")&"&gb="&gb&""">添加</a> <a href=""admin_prodclass.asp?action=EditClass&ID="&rstemp("id")&"&ParentID="&rstemp("parentid")&"&gb="&gb&""">编辑</a> <a href=""admin_prodclass.asp?action=DelClass&ParentID="&rstemp("id")&"&gb="&gb&""" onclick=""{if(confirm('确定删除此项目吗?')){return true;}return false;}"">删除</a></td></tr>"
			else
				getcatalogs = getcatalogs &"<tr bgcolor=f6f6f6 onmouseover=""this.style.backgroundColor='#C1DEFB'"" onmouseout=""this.style.backgroundColor=''""><td height=""25"" class=""underline"" align=""center"" width=""5%""><b>&nbsp;</b></td><td height=""25"" class=""underline"">"&tempStr&"<img src='images/T.png' align='absmiddle'><img src='images/e.png' align='absmiddle'>"&rstemp("orders")&"."&rstemp("classname")&"&nbsp;&nbsp;(<font color='#ff6600'>"&rstemp("classname_sp")&"</font>)</td><td height=""25"" class=""underline"" width=""30%"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=""admin_prodclass.asp?action=AddClass&ParentID="&rstemp("id")&"&gb="&gb&""">添加</a> <a href=""admin_prodclass.asp?action=EditClass&ID="&rstemp("id")&"&ParentID="&rstemp("parentid")&"&gb="&gb&""">编辑</a> <a href=""admin_prodclass.asp?action=DelClass&ParentID="&rstemp("id")&"&gb="&gb&""" onclick=""{if(confirm('确定删除此项目吗?')){return true;}return false;}"">删除</a></td></tr>"
			end if	
			getcatalogs = getcatalogs & getcatalogs(rstemp("id"),tempStr)
		rstemp.MoveNext
		loop
		Set rstemp = Nothing
	End Function


	Function getDownlist(parentid,str)
		Dim Sql,rstemp,TemptempStr,tempStr 
		tempStr=str&"│　"
		sql="select * from pclass where parentid="&parentid&""
		Set rstemp = conn.Execute(Sql)
		if rstemp.eof then exit function
		do while Not rstemp.Eof
			dim mySelect
			mySelect = ""
			if cint(request("ParentID"))=cint(rstemp("id")) Then mySelect = "selected"
			getDownlist = getDownlist & "<option value="& rstemp("id") &" "& mySelect &">" & tempStr &"├"& rstemp("classname") & "</option>"
			getDownlist = getDownlist & getDownlist(rstemp("id"),tempStr)
		rstemp.MoveNext
		loop
		Set rstemp = Nothing
	End Function
	
	'*************************************************************
	'函数作用：查询出来某一类别下的所有下载类别（无限搜索）
	'*************************************************************
	Function selectclass(ID)
		Dim sql,rstemp
		sql="select * from pclass where parentid="&ID
		Set rstemp = conn.Execute(Sql)
		if rstemp.eof then exit function
		do while Not rstemp.Eof
			if selectclass = "" then
				selectclass = rstemp("id")& "," & selectclass(rstemp("id")) 
			else
				selectclass = selectclass & "," & rstemp("id") & "," & selectclass(rstemp("id"))
			end if
		rstemp.Movenext
		loop
		Set rstemp = Nothing
	End Function
	
	
		'*************************************************************
	'函数作用：返回下载类别列表（无限级别）
	'*************************************************************
	Function getcatalogs2(parentid,str)
		Dim Sql,rstemp,TemptempStr,tempStr
		if str="" then 
			tempStr=str&"<img src='images/N.png' align='absmiddle'>"
		else
			tempStr=str&"<img src='images/I.png' align='absmiddle'>"
		end if
		sql="select * from wclass where parentid="&parentid&" order by orders asc"
		Set rstemp = conn.Execute(Sql)
		if rstemp.eof then exit function
		do while Not rstemp.Eof
		Set rsP = conn.execute("select * from wclass where parentid="&rstemp("id"))
		if session("admin")="admin" then 
			if rsP.eof then
				getcatalogs2 = getcatalogs2 &"<tr bgcolor=f6f6f6 onmouseover=""this.style.backgroundColor='#C1DEFB'"" onmouseout=""this.style.backgroundColor=''""><td height=""25"" class=""underline"" align=""center"" width=""5%""><b>&nbsp;</b></td><td height=""25"" class=""underline"">"&tempStr&"<img src='images/T.png' align='absmiddle'><img src='images/file.png' align='absmiddle'>"&rstemp("orders")&"."&rstemp("classname")&"&nbsp;("&rstemp("id")&")</td><td height=""25"" class=""underline"" width=""30%"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""admin_wzclass.asp?action=EditClass&ID="&rstemp("id")&"&ParentID="&rstemp("parentid")&"&gb="&gb&""">编辑内容</a> <a href=""admin_wzclass.asp?action=DelClass&ParentID="&rstemp("id")&"&gb="&gb&""" onclick=""{if(confirm('确定删除此项目吗?')){return true;}return false;}"">删除栏目</a></td></tr>"
			else
				getcatalogs2 = getcatalogs2 &"<tr bgcolor=f6f6f6 onmouseover=""this.style.backgroundColor='#C1DEFB'"" onmouseout=""this.style.backgroundColor=''""><td height=""25"" class=""underline"" align=""center"" width=""5%""><b>&nbsp;</b></td><td height=""25"" class=""underline"">"&tempStr&"<img src='images/T.png' align='absmiddle'><img src='images/e.png' align='absmiddle'>"&rstemp("orders")&"."&rstemp("classname")&"&nbsp;("&rstemp("id")&")</td><td height=""25"" class=""underline"" width=""30%"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=""admin_wzclass.asp?action=EditClass&ID="&rstemp("id")&"&ParentID="&rstemp("parentid")&"&gb="&gb&""">编辑内容</a> <a href=""admin_wzclass.asp?action=DelClass&ParentID="&rstemp("id")&"&gb="&gb&""" onclick=""{if(confirm('确定删除此项目吗?')){return true;}return false;}"">删除栏目</a></td></tr>"
			end if	
		else
			if rsP.eof then
				getcatalogs2 = getcatalogs2 &"<tr bgcolor=f6f6f6 onmouseover=""this.style.backgroundColor='#C1DEFB'"" onmouseout=""this.style.backgroundColor=''""><td height=""25"" class=""underline"" align=""center"" width=""5%""><b>&nbsp;</b></td><td height=""25"" class=""underline"">"&tempStr&"<img src='images/T.png' align='absmiddle'><img src='images/file.png' align='absmiddle'>"&rstemp("orders")&"."&rstemp("classname")&"</td><td height=""25"" class=""underline"" width=""30%"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""admin_wzclass.asp?action=EditClass&ID="&rstemp("id")&"&ParentID="&rstemp("parentid")&"&gb="&gb&""">编辑内容</a> <a href=""admin_wzclass.asp?action=DelClass&ParentID="&rstemp("id")&"&gb="&gb&""" onclick=""{if(confirm('确定删除此项目吗?')){return true;}return false;}"">删除栏目</a></td></tr>"
			else
				getcatalogs2 = getcatalogs2 &"<tr bgcolor=f6f6f6 onmouseover=""this.style.backgroundColor='#C1DEFB'"" onmouseout=""this.style.backgroundColor=''""><td height=""25"" class=""underline"" align=""center"" width=""5%""><b>&nbsp;</b></td><td height=""25"" class=""underline"">"&tempStr&"<img src='images/T.png' align='absmiddle'><img src='images/e.png' align='absmiddle'>"&rstemp("orders")&"."&rstemp("classname")&"</td><td height=""25"" class=""underline"" width=""30%"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href=""admin_wzclass.asp?action=EditClass&ID="&rstemp("id")&"&ParentID="&rstemp("parentid")&"&gb="&gb&""">编辑内容</a> <a href=""admin_wzclass.asp?action=DelClass&ParentID="&rstemp("id")&"&gb="&gb&""" onclick=""{if(confirm('确定删除此项目吗?')){return true;}return false;}"">删除栏目</a></td></tr>"
			end if	
		end if
			getcatalogs2 = getcatalogs2 & getcatalogs2(rstemp("id"),tempStr)
		rstemp.MoveNext
		loop
		Set rstemp = Nothing
	End Function
	
End Class
%>