﻿<%
'######################################################################
'## easp.db.asp
'## -------------------------------------------------------------------
'## Feature     :   Database Control
'## Version     :   v2.2 Alpha
'## Author      :   Coldstone(coldstone[at]qq.com)
'## Update Date :   2010/03/22 21:44:16
'## Description :   EasyAsp database controller
'##
'######################################################################
Class EasyAsp_db
	Private s_dbType, s_pageParam, s_pageSpName
	Private i_queryType, i_errNumber, i_pageIndex, i_pageSize, i_pageCount, i_recordCount
	Private o_conn, o_pageDic
	Private Sub Class_Initialize()
		Easp.Error(11) = "无效的查询条件，无法获取记录集！"
		Easp.Error(12) = "数据库服务器端连接错误，请检查数据库连接信息（用户名/密码/服务器地址/端口号/数据库名）是否正确！"
		Easp.Error(13) = "无效的数据库连接！"
		Easp.Error(14) = "无效的查询条件，无法获取新的ID号！"
		Easp.Error(15) = "生成Json格式代码出错！"
		Easp.Error(16) = "生成不重复的随机字符串出错！"
		Easp.Error(17) = "生成不重复的随机数出错！"
		Easp.Error(18) = "获取随机记录失败，请输入要取的记录数量！"
		Easp.Error(19) = "获取随机记录失败，请在表名后输入:ID字段的名称！"
		Easp.Error(20) = "向数据库添加记录出错！"
		Easp.Error(21) = "更新数据库记录出错！"
		Easp.Error(22) = "从数据库删除数据出错！"
		Easp.Error(23) = "从数据库获取数据出错！"
		Easp.Error(32) = "仅支持从MS SQL Server数据库调用存储过程！"
		Easp.Error(24) = "调用存储过程出错！"
		Easp.Error(25) = "执行SQL语句出错！"
		Easp.Error(26) = "生成SQL语句出错！"
		Easp.Error(27) = "获取分页数据出错，数组必须是4个元素（必须提供数据库表的主键）！"
		Easp.Error(28) = "获取分页数据出错，使用数组条件获取分页数据时条件参数必须为数组！"
		Easp.Error(29) = "获取分页数据出错，使用自带分页存储过程时条件数组参数必须为6个元素！"
		Easp.Error(30) = "获取分页数据出错，使用自定义分页存储过程时必须包含@@RecordCount和@@PageCount输出参数！"
		Easp.Error(31) = "获取分页数据出错，使用存储过程获取分页数据时条件参数必须为数组！"
		s_dbType       = ""
		i_queryType    = 0
		s_pageParam    = "page"
		i_pageSize     = 20
		s_pageSpName   = "easp_sp_pager"
		Set o_pageDic = Server.CreateObject("Scripting.Dictionary")
		o_pageDic("default_html") = "<div class=""pager"">{first}{prev}{liststart}{list}{listend}{next}{last} 跳转到{jump}页</div>"
		o_pageDic("default_config") = ""
	End Sub
	Private Sub Class_Terminate()
		If TypeName(o_conn) = "Connection" Then
			If o_conn.State = 1 Then o_conn.Close()
			Set o_conn = Nothing
		End If
		Set o_pageDic = Nothing
	End Sub
	Public Property Let Conn(ByVal pdbConn)
		If TypeName(pdbConn) = "Connection" Then
			Set o_conn = pdbConn
			s_dbType = GetDataType(pdbConn)
		Else
			Easp.Error.Raise 13
		End If
	End Property
	Public Property Get Conn()
		If TypeName(o_conn) = "Connection" Then
			Set Conn = o_conn
		Else
			Easp.Error.Raise 13
		End If
	End Property
	Public Property Get DatabaseType()
		DatabaseType = s_dbType
	End Property
	Public Property Let QueryType(ByVal str)
		str = Lcase(str)
		If str = "1" or str = "command" Then
			i_queryType = 1
		Else
			i_queryType = 0
		End If
	End Property
	Public Property Let PageSize(ByVal num)
		i_pageSize = num
	End Property
	Public Property Get PageSize()
		PageSize = i_pageSize
	End Property
	Public Property Get PageCount()
		PageCount = i_pageCount
	End Property
	Public Property Get PageIndex()
		PageIndex = Easp.IIF(Easp.isN(i_pageIndex),GetCurrentPage,i_pageIndex)
	End Property
	Public Property Get PageRecordCount()
		PageRecordCount = i_recordCount
	End Property
	Public Property Let PageParam(ByVal str)
		s_pageParam = str
	End Property
	Public Property Let PageSpName(ByVal str)
		s_pageSpName = str
	End Property
	Public Function [New]()
		Set [New] = New EasyASP_db
	End Function
	Public Function OpenConn(ByVal dbType, ByVal strDB, ByVal strServer)
		Dim TempStr, objConn, s, u, p, port
		s = "" : u = "" : p = "" : port = ""
		If Instr(strServer,"@")>0 Then
			s = Trim(Mid(strServer,InstrRev(strServer,"@")+1))
			u = Trim(Left(strServer,InstrRev(strServer,"@")-1))
			If Instr(s,":")>0 Then : port = Trim(Mid(s,Instr(s,":")+1)) : s = Trim(Left(s,Instr(s,":")-1))
			If Instr(u,":")>0 Then : p = Trim(Mid(u,Instr(u,":")+1)) : u = Trim(Left(u,Instr(u,":")-1))
		Else
			If Instr(strServer,":")>0 Then
				u = Trim(Left(strServer,Instr(strServer,":")-1))
				p = Trim(Mid(strServer,Instr(strServer,":")+1))
			Else
				p = Trim(strServer)
			End If
		End If
		s_dbType = UCase(Cstr(dbType))
		Select Case s_dbType
			Case "0","MSSQL"
				TempStr = "Provider=sqloledb;Data Source=" & s & Easp.IfThen(Easp.Has(port), "," & port) & ";Initial Catalog="&strDB&";User Id="&u&";Password="&p&";"
			Case "1","ACCESS"
				Dim tDb : If Instr(strDB,":")>0 Then : tDb = strDB : Else : tDb = Server.MapPath(strDB) : End If
				TempStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&tDb&";Jet OLEDB:Database Password="&p&";"
			Case "2","MYSQL"
				If port = "" Then port = "3306"
				TempStr = "Driver={mySQL};Server="&s&";Port="&port&";Option=131072;Stmt=;Database="&strDB&";Uid="&u&";Pwd="&p&";"
			Case "3","ORACLE"
				TempStr = "Provider=msdaora;Data Source="&s&";User Id="&u&";Password="&p&";"
		End Select
		Set OpenConn = CreatConn(TempStr)
	End Function
	Public Function CreatConn(ByVal ConnStr)
		On Error Resume Next
		Dim objConn : Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open ConnStr
		If Err.number <> 0 Then
			Easp.Error.Msg = "<br />(""" & ConnStr & """)"
			Easp.Error.Raise 12
			objConn.Close
			Set objConn = Nothing
		End If
		Set CreatConn = objConn
	End Function
	Private Function GetDataType(ByVal connObj)
		Dim str,i : str = UCase(connObj.Provider)
		Dim MSSQL, ACCESS, MYSQL, ORACLE
		MSSQL = Split("SQLNCLI10, SQLXMLOLEDB, SQLNCLI, SQLOLEDB, MSDASQL",", ")
		ACCESS = Split("MICROSOFT.ACE.OLEDB.12.0, MICROSOFT.JET.OLEDB.4.0",", ")
		MYSQL = "MYSQLPROV"
		ORACLE = Split("MSDAORA, OLEDB.ORACLE",", ")
		For i = 0 To Ubound(MSSQL)
			If Instr(str,MSSQL(i))>0 Then
				GetDataType = "MSSQL" : Exit Function
			End If
		Next
		For i = 0 To Ubound(ACCESS)
			If Instr(str,ACCESS(i))>0 Then
				GetDataType = "ACCESS" : Exit Function
			End If
		Next
		If Instr(str,MYSQL)>0 Then
			GetDataType = "MYSQL" : Exit Function
		End If
		For i = 0 To Ubound(ORACLE)
			If Instr(str,ORACLE(i))>0 Then
				GetDataType = "ORACLE" : Exit Function
			End If
		Next
	End Function
	Public Function AutoID(ByVal TableName)
		On Error Resume Next
		Dim rs, tmp, fID, tmpID : fID = "" : tmpID = 0
		tmp = Easp_Param(TableName)
		If Easp.Has(tmp(1)) Then : TableName = tmp(0) : fID = tmp(1) : tmp = "" : End If
		Set rs = GRS("Select " & Easp.IIF(fID<>"", "Max("&fID&")", "Top 1 *") & " From ["&TableName&"]")
		If rs.eof Then
			AutoID = 1 : Exit Function
		Else
			If fID<>"" Then
				If Easp.isN(rs.Fields.Item(0).Value) Then AutoID = 1 : Exit Function
				AutoID = rs.Fields.Item(0).Value + 1 : Exit Function
			Else
				Dim newRs
				Set newRs = GRS("Select Max("&rs.Fields.Item(0).Name&") From ["&TableName&"]")
				tmpID = newRS.Fields.Item(0).Value + 1
				newRs.Close() : Set newRs = Nothing
			End If
		End If
		If Err.number <> 0 Then Easp.Error.Raise 14
		rs.Close() : Set rs = Nothing
		AutoID = tmpID
	End Function
	Public Function GetRecord(ByVal TableName,ByVal Condition,ByVal OrderField)
		Set GetRecord = GRS(wGetRecord(TableName,Condition,OrderField))
	End Function
	Public Function wGetRecord(ByVal TableName,ByVal Condition,ByVal OrderField)
		Dim strSelect, FieldsList, ShowN, o, p
		FieldsList = "" : ShowN = 0
		o = Easp_Param(TableName)
		If Easp.Has(o(1)) Then
			TableName = Trim(o(0)) : FieldsList = Trim(o(1)) : o = ""
			p = Easp_Param(FieldsList)
			If Easp.Has(p(1)) Then
				FieldsList = Trim(p(0)) : ShowN = Int(Trim(p(1))) : p = ""
			Else
				If isNumeric(FieldsList) Then ShowN = Int(FieldsList) : FieldsList = ""
			End If
		End If
		strSelect = "Select "
		If ShowN > 0 Then strSelect = strSelect & "Top " & ShowN & " "
		strSelect = strSelect & Easp.IIF(FieldsList <> "", FieldsList, "* ")
		strSelect = strSelect & " From [" & TableName & "]"
		If isArray(Condition) Then
			strSelect = strSelect & " Where " & ValueToSql(TableName,Condition,1)
		Else
			If Condition <> "" Then strSelect = strSelect & " Where " & Condition
		End If
		If OrderField <> "" Then strSelect = strSelect & " Order By " & OrderField
		wGetRecord = strSelect
	End Function
	Public Function GR(ByVal TableName,ByVal Condition,ByVal OrderField)
		Set GR = GetRecord(TableName, Condition, OrderField)
	End Function
	Public Function wGR(ByVal TableName,ByVal Condition,ByVal OrderField)
		wGR = wGetRecord(TableName, Condition, OrderField)
	End Function
	Public Function GetRecordBySQL(ByVal s)
		On Error Resume Next
		If i_queryType = 1 Then
			Dim cmd : Set cmd = Server.CreateObject("ADODB.Command")
			With cmd
				.ActiveConnection = o_conn
				.CommandText = s
				Set GetRecordBySQL = .Execute
			End With
			Set cmd = Nothing
		Else
			Dim rs : Set rs = Server.CreateObject("Adodb.Recordset")
			With rs
				.ActiveConnection = o_conn
				.CursorType = 1
				.LockType = 1
				.Source = s
				.Open
			End With
			Set GetRecordBySQL = rs
		End If
		Easp_DbQueryTimes = Easp_DbQueryTimes + 1
		If Err.number <> 0 Then
			Easp.Error.Msg = "<br />" & s
			Easp.Error.Raise 11
			Err.Clear
		End If
	End Function
	Public Function GRS(ByVal s)
		Set GRS = GetRecordBySQL(s)
	End Function
	Public Function Json(ByVal jRs, ByVal jName)
		On Error Resume Next
		Dim tmpStr, rs, fi, o, totalName, total, tName, tValue
		o = Easp_Param(jName)
		If Easp.Has(o(1)) Then
			jName = o(0) : totalName = o(1)
		End If
		Set rs = jRs.Clone
		Easp.Use "JSON"
		Set o = Easp.Json.New(0)
		total = 0
		If Easp.Has(rs) Then
			total = rs.RecordCount
			If Easp.Has(totalName) Then o(totalName) = total
			o(jName) = Easp.Json.New(1)
			While Not rs.Eof
				o(jName)(Null) = Easp.Json.New(0)
				For Each fi In rs.Fields
					o(jName)(Null)(fi.Name) = fi.Value
				Next
				rs.MoveNext
			Wend
		End If
		If Err.number <> 0 Then Easp.Error.Raise 15
		tmpStr = o.JsString
		Set o = Nothing
		rs.Close() : Set rs = Nothing
		Json = tmpStr
	End Function
	Public Function RandStr(length,TableField)
		On Error Resume Next
		Dim tb, fi, tmpStr, rs
		tb = Easp_Param(TableField)(0)
		fi = Easp_Param(TableField)(1)
		tmpStr = Easp.RandStr(length)
		Do While (True)
			Set rs = GR(tb&":"&fi&":1",fi&"='"&tmpStr&"'","")
			If Not rs.Bof And Not rs.Eof Then
				tmpStr = Easp.RandStr(length)
			Else
				RandStr = tmpStr
				Exit Do
			End If
			C(rs)
		Loop
		If Err.number <> 0 Then Easp.Error.Raise 16
	End Function
	Public Function Rand(min,max,TableField)
		On Error Resume Next
		Dim tb, fi, tmpInt, rs
		tb = Easp_Param(TableField)(0)
		fi = Easp_Param(TableField)(1)
		tmpInt = Easp.Rand(min,max)
		Do While (True)
			Set rs = GR(tb&":"&fi&":1",Array(fi&":"&tmpInt),"")
			If Not rs.Bof And Not rs.Eof Then
				tmpInt = Easp.Rand(min,max)
			Else
				Rand = tmpInt
				Exit Do
			End If
			C(rs)
		Loop
		If Err.number <> 0 Then Easp.Error.Raise 17
	End Function
	Public Function GetRecordDetail(ByVal TableName,ByVal Condition)
		Dim strSelect
		strSelect = "Select * From [" & TableName & "] Where " & ValueToSql(TableName,Condition,1)
		Set GetRecordDetail = GRS(strSelect)
	End Function
	Public Function GRD(ByVal TableName,ByVal Condition)
		Set GRD = GetRecordDetail(TableName, Condition)
	End Function
	Public Function GetRandRecord(ByVal TableName,ByVal Condition)
		Dim sql,o,p,fi,IdField,showN,where
		o = Easp_Param(TableName)
		If Easp.Has(o(1)) Then
			TableName = o(0)
			p = Easp_Param(o(1))
			If Easp.isN(p(1)) Then
				Easp.Error.Raise 18
				Exit Function
			Else
				fi = p(0) : showN = p(1)
				If Instr(fi,",")>0 Then
					IdField = Trim(Left(fi,Instr(fi,",")-1))
				Else
					IdField = fi : fi = "*"
				End If
			End If
		Else
			Easp.Error.Raise 19
			Exit Function
		End If
		Condition = Easp.IIF(Easp.isN(Condition),""," Where " & ValueToSql(TableName,Condition,1))
		sql = "Select Top " & showN & " " & fi & " From ["&TableName&"]" & Condition
		Select Case s_dbType
			Case "ACCESS" : Randomize
				sql = sql & " Order By Rnd(-(" & IdField & "+" & Rnd() & "))"
			Case "MSSQL"
				sql = sql & " Order By newid()"
			Case "MYSQL"
				sql = "Select " & fi & " From " & TableName & Condition & " Order By rand() limit " & showN
			Case "ORACLE"
				sql = "Select " & fi & " From (Select " & fi & " From "&TableName&" Order By dbms_random.value) " & Easp.IIF(Easp.isN(Condition),"Where",Condition & " And") & " rownum < " & Int(showN)+1
		End Select
		Set GetRandRecord = GRS(sql)
	End Function
	Public Function GRR(ByVal TableName,ByVal Condition)
		Set GRR = GetRandRecord(TableName,Condition)
	End Function
	Public Function AddRecord(ByVal TableName,ByVal ValueList)
		On Error Resume Next
		Dim o,s : o = Easp_Param(TableName)
		If Easp.Has(o(1)) Then TableName = o(0)
		s = wAddRecord(TableName,ValueList)
		DoExecute s
		If Err.number <> 0 Then
			Easp.Error.Msg = "<br />" & s
			Easp.Error.Raise 20
			AddRecord = 0
			Exit Function
		End If
		If Easp.Has(o(1)) Then
			AddRecord = AutoID(o(0)&":"&o(1))-1
		Else
			AddRecord = 1
		End If
	End Function
	Public Function wAddRecord(ByVal TableName,ByVal ValueList)
		Dim TempSQL, TempFiled, TempValue, o
		o = Easp_Param(TableName) : If Easp.Has(o(1)) Then TableName = o(0)
		TempFiled = ValueToSql(TableName,ValueList,2)
		TempValue = ValueToSql(TableName,ValueList,3)
		TempSQL = "Insert Into [" & TableName & "] (" & TempFiled & ") Values (" & TempValue & ")"
		wAddRecord = TempSQL
	End Function
	Public Function AR(ByVal TableName,ByVal ValueList)
		AR = AddRecord(TableName,ValueList)
	End Function
	Public Function wAR(ByVal TableName,ByVal ValueList)
		wAR = wAddRecord(TableName,ValueList)
	End Function
	Public Function UpdateRecord(ByVal TableName,ByVal Condition,ByVal ValueList)
		On Error Resume Next
		Dim s : s = wUpdateRecord(TableName,Condition,ValueList)
		DoExecute s
		If Err.number <> 0 Then
			Easp.Error.Msg = "<br />" & s
			Easp.Error.Raise 21
			UpdateRecord = 0
			Exit Function
		End If
		UpdateRecord = 1
	End Function
	Public Function wUpdateRecord(ByVal TableName,ByVal Condition,ByVal ValueList)
		Dim TmpSQL
		TmpSQL = "Update ["&TableName&"] Set "
		TmpSQL = TmpSQL & ValueToSql(TableName,ValueList,0)
		If Easp.Has(Condition) Then TmpSQL = TmpSQL & " Where " & ValueToSql(TableName,Condition,1)
		wUpdateRecord = TmpSQL
	End Function
	Public Function UR(ByVal TableName,ByVal Condition,ByVal ValueList)
		UR = UpdateRecord(TableName, Condition, ValueList)
	End Function
	Public Function wUR(ByVal TableName,ByVal Condition,ByVal ValueList)
		wUR = wUpdateRecord(TableName, Condition, ValueList)
	End Function
	Public Function DeleteRecord(ByVal TableName,ByVal Condition)
		On Error Resume Next
		Dim s : s = wDeleteRecord(TableName,Condition)
		DoExecute s
		If Err.number <> 0 Then
			Easp.Error.Msg = "<br />" & s
			Easp.Error.Raise 22
			DeleteRecord = 0
			Exit Function
		End If
		DeleteRecord = 1
	End Function
	Public Function wDeleteRecord(ByVal TableName,ByVal Condition)
		Dim IDFieldName, IDValues, Sql, p : IDFieldName = "" : IDValues = ""
		If Not isArray(Condition) Then
			p = Easp_Param(Condition)
			If Easp.Has(p(1)) Then
				IDFieldName = p(0)
				If Instr(IDFieldName," ")=0 Then
					IDValues = p(1)
				Else
					IDFieldName = ""
				End If
			End If
		End If
		Sql = "Delete From ["&TableName&"] Where " & Easp.IIF(IDFieldName="", ValueToSql(TableName,Condition,1), "["&IDFieldName&"] In (" & IDValues & ")")
		wDeleteRecord = Sql
	End Function
	Public Function DR(ByVal TableName,ByVal Condition)
		DR = DeleteRecord(TableName, Condition)
	End Function
	Public Function wDR(ByVal TableName,ByVal Condition)
		wDR = wDeleteRecord(TableName, Condition)
	End Function
	Public Function ReadTable(ByVal TableName,ByVal Condition,ByVal GetFieldNames)
		On Error Resume Next
		Dim rs,Sql,arrTemp,arrStr,TempStr,i
		TempStr = "" : arrStr = ""
		Sql = "Select "&GetFieldNames&" From ["&TableName&"] Where " & ValueToSql(TableName,Condition,1)
		Set rs = GRS(Sql)
		If Not rs.Eof Then
			If Instr(GetFieldNames,",") > 0 Then
				arrTemp = Split(GetFieldNames,",")
				For i = 0 To Ubound(arrTemp)
					If i<>0 Then arrStr = arrStr & Chr(0)
					arrStr = arrStr & rs.Fields.Item(i).Value
				Next
				TempStr = Split(arrStr,Chr(0))
			Else
				TempStr = rs.Fields.Item(0).Value
			End If
		End If
		If Err.number <> 0 Then Easp.Error.Raise 23
		rs.close() : Set rs = Nothing : Err.Clear
		ReadTable = TempStr
	End Function
	Public Function RT(ByVal TableName,ByVal Condition,ByVal GetFieldNames)
		RT = ReadTable(TableName, Condition, GetFieldNames)
	End Function
	Public Function doSP(ByVal spName, ByVal spParam)
		On Error Resume Next
		Dim p, spType, cmd, outParam, i, NewRS : spType = ""
		If Not s_dbType="0" And Not s_dbType="MSSQL" Then
			Easp.Error.Raise 32
			Exit Function
		End If
		p = Easp_Param(spName)
		If Easp.Has(p(1)) Then : spType = UCase(Trim(p(1))) : spName = Trim(p(0)) : p = "" : End If
		Set cmd = Server.CreateObject("ADODB.Command")
			With cmd
				.ActiveConnection = o_conn
				.CommandText = spName
				.CommandType = 4
				.Prepared = true
				.Parameters.append .CreateParameter("return",3,4)
				outParam = "return"
				If Not IsArray(spParam) Then
					If spParam<>"" Then
						spParam = Easp.IIF(Instr(spParam,",")>0, spParam = Split(spParam,","), Array(spParam))
					End If
				End If
				If IsArray(spParam) Then
					For i = 0 To Ubound(spParam)
						Dim pName, pValue
						If (spType = "1" or spType = "OUT" or spType = "3" or spType = "ALL") And Instr(spParam(i),"@@")=1 Then
							.Parameters.append .CreateParameter(spParam(i),200,2,8000)
							outParam = outParam & "," & spParam(i)
						Else
							If Instr(spParam(i),"@")=1 And Instr(spParam(i),":")>2 Then
								pName = Left(spParam(i),Instr(spParam(i),":")-1)
								outParam = outParam & "," & pName
								pValue = Mid(spParam(i),Instr(spParam(i),":")+1)
								If pValue = "" Then pValue = NULL
								.Parameters.append .CreateParameter(pName,200,1,8000,pValue)
							Else
								.Parameters.append .CreateParameter("@param"&(i+1),200,1,8000,spParam(i))
								outParam = outParam & "," & "@param"&(i+1)
							End If
						End If
					Next
				End If
			End With
			outParam = Easp.IIF(Instr(outParam,",")>0, Split(outParam,","), Array(outParam))
			If spType = "1" or spType = "OUT" Then
				cmd.Execute : doSP = cmd
			ElseIf spType = "2" or spType = "RS" Then
				Set doSP = cmd.Execute
			ElseIf spType = "3" or spType = "ALL" Then
				Dim NewOut,pa : Set NewOut = Server.CreateObject("Scripting.Dictionary")
				Set NewRS = cmd.Execute : NewRS.close
				For i = 0 To Ubound(outParam)
					NewOut(Trim(outParam(i))) = cmd(i)
				Next
				NewRs.open : doSP = Array(NewRS,NewOut)
				Set NewOut = Nothing
			Else
				cmd.Execute : doSP = cmd(0)
			End If
		If Err.number <> 0 Then Easp.Error.Raise 24
		Set cmd = Nothing
		Err.Clear
	End Function
	Public Function C(ByRef ObjRs)
		On Error Resume Next
		ObjRs.close()
		Set ObjRs = Nothing
	End Function
	Public Function Exec(ByVal s)
		On Error Resume Next
		If Lcase(Left(s,6)) = "select" Then
			Dim i : i = i_queryType
			i_queryType = 1
			Set Exec = GRS(s)
			i_queryType = i
		Else
			Exec = 1 : DoExecute(s)
			If Err.number <> 0 Then Exec = 0
		End If
		If Err.number <> 0 Then
			Easp.Error.Msg = "<br />" & s
			Easp.Error.Raise 25
			Err.Clear
		End If
	End Function
	
	Private Function ValueToSql(ByVal TableName, ByVal ValueList, ByVal sType)
		On Error Resume Next
		Dim StrTemp : StrTemp = ValueList
		If IsArray(ValueList) Then
			StrTemp = ""
			Dim rsTemp, CurrentField, CurrentValue, i
			Set rsTemp = GRS("Select Top 1 * From [" & TableName & "] Where 1 = -1")
			For i = 0 to Ubound(ValueList)
				CurrentField = Easp_Param(ValueList(i))(0)
				CurrentValue = Easp_Param(ValueList(i))(1)
				If i <> 0 Then StrTemp = StrTemp & Easp.IIF(sType=1, " And ", ", ")
				If sType = 2 Then
					StrTemp = StrTemp & "[" & CurrentField & "]"
				Else
					Select Case rsTemp.Fields(CurrentField).Type
						Case 8,129,130,133,134,200,201,202,203
							StrTemp = StrTemp & Easp.IIF(sType = 3, "'"&CurrentValue&"'", "[" & CurrentField & "] = '"&CurrentValue&"'")
						Case 7,135
							CurrentValue = Easp.IIF(Easp.IsN(CurrentValue),"NULL","'"&CurrentValue&"'")
							StrTemp = StrTemp & Easp.IIF(sType = 3, CurrentValue, "[" & CurrentField & "] = " & CurrentValue)
						Case 11
							Dim tmpTF, tmpTFV : tmpTFV = UCase(cstr(Trim(CurrentValue)))
							tmpTF = Easp.IIF(tmpTFV="TRUE" or tmpTFV = "1", Easp.IIF(s_dbType="ACCESS","True","1"), Easp.IIF(s_dbType="ACCESS",Easp.IIF(tmpTFV="","NULL","False"),Easp.IIF(tmpTFV="","NULL","0")))
							StrTemp = StrTemp & Easp.IIF(sType = 3, tmpTF, "[" & CurrentField & "] = " & tmpTF)
						Case Else
							CurrentValue = Easp.IIF(Easp.IsN(CurrentValue),"NULL",CurrentValue)
							StrTemp = StrTemp & Easp.IIF(sType = 3, CurrentValue, "[" & CurrentField & "] = " & CurrentValue)
					End Select
				End If
			Next
			If Err.number <> 0 Then Easp.Error.Raise 26
			rsTemp.Close() : Set rsTemp = Nothing : Err.Clear
		End If
		ValueToSql = StrTemp
	End Function
	Private Function DoExecute(ByVal sql)
		Dim ExecuteCmd : Set ExecuteCmd = Server.CreateObject("ADODB.Command")
		With ExecuteCmd
			.ActiveConnection = o_conn
			.CommandText = sql
			.Execute
		End With
		Set ExecuteCmd = Nothing
	End Function
	Public Function GetPageRecord(ByVal PageSetup, ByVal Condition)
		On Error Resume Next
		Dim pType,spResult,rs,o,p,Sql,n,i,spReturn
		o = Easp_Param(Cstr(PageSetup))
		pType = o(0)
		If Easp.Has(o(1)) Then
			p = Easp_Param(o(1))
			If Easp.Has(p(1)) Then
				s_pageParam = Lcase(p(0))
				i_pageSize = Int(p(1))
			Else
				If isNumeric(o(1)) Then
					i_pageSize = Int(o(1))
				Else
					s_pageParam = Lcase(o(1))
				End If
			End If
		End If
		i_pageIndex = GetCurrentPage()
		Select Case Lcase(pType)
			Case "array","0"
				If isArray(Condition) Then
					Dim Table,Fi,Where
					o = Easp_Param(Condition(0))
					If Easp.Has(o(1)) Then
						Table = o(0) : Fi = o(1)
					Else
						Table = Condition(0) : Fi = "*"
					End If
					If isArray(Condition(1)) Then
						Where = ValueToSql(Table,Condition(1),1)
					Else
						Where = Condition(1)
					End If
					i_recordCount = Int(RT(Table, Easp.IIF(Easp.isN(Where),"1=1",Where), "Count(0)"))
					n = i_recordCount / i_pageSize
					i_pageCount = Easp.IIF(n=Int(n), n, Int(n)+1)
					i_pageIndex = Easp.IIF(i_pageIndex > i_pageCount, i_pageCount, i_pageIndex)
					If s_dbType = "1" or s_dbType = "ACCESS" Then
						Set rs = GR(Table&":"&Fi,Where,Condition(2))
						rs.PageSize = i_pageSize
						If i_recordCount>0 Then rs.AbsolutePage = i_pageIndex
						Set GetPageRecord = rs : Exit Function
					ElseIf s_dbType = "2" or s_dbType = "MYSQL" Then
						Sql = "Select "& fi & " From [" & Table & "]"
						If Easp.Has(Where) Then Sql = Sql & " Where " & Where
						If Easp.Has(Condition(2)) Then Sql = Sql & " Order By " & Condition(2)
						Sql = Sql & " Limit " & i_pageSize*(i_pageIndex-1) & ", " & i_pageSize
					Else
						If Ubound(Condition)<>3 Then Easp.Error.Raise 27
						Sql = "Select Top " & i_pageSize & " " & fi
						Sql = Sql & " From [" & Table & "]"
						If Easp.Has(Where) Then Sql = Sql & " Where " & Where
						If i_pageIndex > 1 Then
							Sql = Sql & " " & Easp.IIF(Easp.isN(Where), "Where", "And") & " " & Condition(3) & " Not In ("
							Sql = Sql & "Select Top " & i_pageSize * (i_pageIndex-1) & " " & Condition(3) & " From [" & Table & "]"
							If Easp.Has(Where) Then Sql = Sql & " Where " & Where
							If Easp.Has(Condition(2)) Then Sql = Sql & " Order By " & Condition(2)
							Sql = Sql & ") "
						End If
						If Easp.Has(Condition(2)) Then Sql = Sql & " Order By " & Condition(2)
					End If
					Set GetPageRecord = GRS(Sql)
				Else
					Easp.Error.Raise 28
				End If
			Case "sql","1" Set rs = GRS(Condition)
			Case "rs","2" Set rs = Condition
			Case Else
				If isArray(Condition) Then
					If pType = "" Then pType = s_pageSpName
					Select Case pType
						Case "easp_sp_pager"
							If Ubound(Condition)<>5 Then Easp.Error.Raise 29
							spResult = doSP("easp_sp_pager:3",Array("@TableName:"&Condition(0),"@FieldList:"&Condition(1),"@Where:"&Condition(2),"@Order:"&Condition(3),"@PrimaryKey:"&Condition(4),"@SortType:"&Condition(5),"@RecorderCount:0","@pageSize:"&i_pageSize,"@PageIndex:"&i_pageIndex,"@@RecordCount","@@PageCount"))
						Case Else
							spReturn = Array(False,False)
							For i = 0 To Ubound(Condition)
								If LCase(Condition(i)) = "@@recordcount" Then spReturn(0) = True
								If LCase(Condition(i)) = "@@pagecount" Then spReturn(1) = True
								If spReturn(0) And spReturn(1) Then Exit For
							Next
							If spReturn(0) And spReturn(1) Then
								spResult = doSP(pType&":3",Condition)
							Else
								Easp.Error.Raise 30
							End If
					End Select
					Set GetPageRecord = spResult(0)
					i_recordCount = int(spResult(1)("@@RecordCount"))
					i_pageCount = int(spResult(1)("@@PageCount"))
					i_pageIndex = Easp.IIF(i_pageIndex > i_pageCount, i_pageCount, i_pageIndex)
				Else
					Easp.Error.Raise 31
				End If
		End Select
		If Instr(",sql,rs,1,2,", "," & pType & ",")>0 Then
			i_recordCount = rs.RecordCount
			rs.PageSize = i_pageSize
			i_pageCount = rs.PageCount
			i_pageIndex = Easp.IIF(i_pageIndex > i_pageCount, i_pageCount, i_pageIndex)
			If i_recordCount>0 Then rs.AbsolutePage = i_pageIndex
			Set GetPageRecord = rs
		End If
	End Function
	Public Function GPR(ByVal PageSetup, ByVal Condition)
		Set GPR = GetPageRecord(PageSetup, Condition)
	End Function
	Public Function Pager(ByVal PagerHtml, ByRef PagerConfig)
		On Error Resume Next
		Dim pList, pListStart, pListEnd, pFirst, pPrev, pNext, pLast
		Dim pJump, pJumpLong, pJumpStart, pJumpEnd, pJumpValue
		Dim i, j, tmpStr, pStart, pEnd, cfg, pcfg(1)
		tmpStr = Easp.IIF(PagerHtml="",o_pageDic("default_html"),PagerHtml)
		Set cfg = Server.CreateObject("Scripting.Dictionary")
		cfg("recordcount")	= i_recordCount
		cfg("pageindex")	= i_pageIndex
		cfg("pagecount")	= i_pageCount
		cfg("pagesize")		= i_pageSize
		cfg("listlong")		= 9
		cfg("listsidelong")	= 2
		cfg("list")			= "*"
		cfg("currentclass")	= "current"
		cfg("link")			= GetRQ("*")
		cfg("first")		= "&laquo;"
		cfg("prev")			= "&#8249;"
		cfg("next")			= "&#8250;"
		cfg("last")			= "&raquo;"
		cfg("more")			= "..."
		cfg("disabledclass")= "disabled"
		cfg("jump")			= "input"
		cfg("jumpplus")		= ""
		cfg("jumpaction")	= ""
		cfg("jumplong")		= 50
		PagerConfig = Easp.IIF(isArray(PagerConfig),PagerConfig, Easp.IIF(Easp.isN(PagerConfig),o_pageDic("default_config"),Array(PagerConfig,"pagerconfig:1")))
		If isArray(PagerConfig) Then
			Dim ConfigName, ConfigValue
			For i = 0 To Ubound(PagerConfig)
				ConfigName = LCase(Left(PagerConfig(i),Instr(PagerConfig(i),":")-1))
				ConfigValue = Mid(PagerConfig(i),Instr(PagerConfig(i),":")+1)
				If Instr(",recordcount,pageindex,pagecount,pagesize,listlong,listsidelong,jumplong,", ","&ConfigName&",") > 0 Then
					cfg(ConfigName) = Int(ConfigValue)
				Else
					cfg(ConfigName) = ConfigValue
				End If
			Next
		End If
		pStart = cfg("pageindex") - ((cfg("listlong") \ 2) + (cfg("listlong") Mod 2)) + 1
		pEnd = cfg("pageindex") + (cfg("listlong") \ 2)
		If pStart < 1 Then
			pStart = 1 : pEnd = cfg("listlong")
		End If
		If pEnd > cfg("pagecount") Then
			pStart = cfg("pagecount") - cfg("listlong") + 1 : pEnd = cfg("pagecount")
		End If
		If pStart < 1 Then pStart = 1
		For i = pStart To pEnd
			If i = cfg("pageindex") Then
				pList = pList & " <span class="""&cfg("currentclass")&""">" & Replace(cfg("list"),"*",i) & "</span> "
			Else
				pList = pList & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
			End If
		Next
		If cfg("listsidelong")>0 Then
			If cfg("listsidelong") < pStart Then
				For i = 1 To cfg("listsidelong")
					pListStart = pListStart & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
				pListStart = pListStart & Easp.IIF(cfg("listsidelong")+1=pStart,"",cfg("more") & " ")
			ElseIf cfg("listsidelong") >= pStart And pStart > 1 Then
				For i = 1 To (pStart - 1)
					pListStart = pListStart & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
			End If
			If (cfg("pagecount") - cfg("listsidelong")) > pEnd Then
				pListEnd = " " & cfg("more") & pListEnd
				For i = ((cfg("pagecount") - cfg("listsidelong"))+1) To cfg("pagecount")
					pListEnd = pListEnd & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
			ElseIf (cfg("pagecount") - cfg("listsidelong")) <= pEnd And pEnd < cfg("pagecount") Then
				For i = (pEnd+1) To cfg("pagecount")
					pListEnd = pListEnd & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
			End If
		End If
		If cfg("pageindex") > 1 Then
			pFirst = " <a href="""&Replace(cfg("link"),"*","1")&""">" & cfg("first") & "</a> "
			pPrev = " <a href="""&Replace(cfg("link"),"*",cfg("pageindex")-1)&""">" & cfg("prev") & "</a> "
		Else
			pFirst = " <span class="""&cfg("disabledclass")&""">" & cfg("first") & "</span> "
			pPrev = " <span class="""&cfg("disabledclass")&""">" & cfg("prev") & "</span> "
		End If
		If cfg("pageindex") < cfg("pagecount") Then
			pLast = " <a href="""&Replace(cfg("link"),"*",cfg("pagecount"))&""">" & cfg("last") & "</a> "
			pNext = " <a href="""&Replace(cfg("link"),"*",cfg("pageindex")+1)&""">" & cfg("next") & "</a> "
		Else
			pLast = " <span class="""&cfg("disabledclass")&""">" & cfg("last") & "</span> "
			pNext = " <span class="""&cfg("disabledclass")&""">" & cfg("next") & "</span> "
		End If
		Select Case LCase(cfg("jump"))
			Case "input"
				pJumpValue = "this.value"
				pJump = "<input type=""text"" size=""3"" title=""请输入要跳转到的页数并回车""" & Easp.IIF(cfg("jumpplus")="",""," "&cfg("jumpplus"))
				pJump = pJump & " onkeydown=""javascript:if(event.charCode==13||event.keyCode==13){if(!isNaN(" & pJumpValue & ")){"
				pJump = pJump & Easp.IIF(cfg("jumpaction")="",Easp.IIF(Lcase(Left(cfg("link"),11))="javascript:",Replace(Mid(cfg("link"),12),"*",pJumpValue),"document.location.href='" & Replace(cfg("link"),"*","'+" & pJumpValue & "+'") & "';"),Replace(cfg("jumpaction"),"*", pJumpValue))
				pJump = pJump & "}return false;}"" />"
			Case "select"
				pJumpValue = "this.options[this.selectedIndex].value"
				pJump = "<select" & Easp.IIF(cfg("jumpplus")="",""," "&cfg("jumpplus")) & " onchange=""javascript:"
				pJump = pJump & Easp.IIF(cfg("jumpaction")="",Easp.IIF(Lcase(Left(cfg("link"),11))="javascript:",Replace(Mid(cfg("link"),12),"*",pJumpValue),"document.location.href='" & Replace(cfg("link"),"*","'+" & pJumpValue & "+'") & "';"),Replace(cfg("jumpaction"),"*",pJumpValue))
				pJump = pJump & """ title=""请选择要跳转到的页数""> "
				If cfg("jumplong")=0 Then
					For i = 1 To cfg("pagecount")
						pJump = pJump & "<option value=""" & i & """" & Easp.IIF(i=cfg("pageindex")," selected=""selected""","") & ">" & i & "</option> "
					Next
				Else
					pJumpLong = Int(cfg("jumplong") / 2)
					pJumpStart = Easp.IIF(cfg("pageindex")-pJumpLong<1, 1, cfg("pageindex")-pJumpLong)
					pJumpStart = Easp.IIF(cfg("pagecount")-cfg("pageindex")<pJumpLong, pJumpStart-(pJumpLong-(cfg("pagecount")-cfg("pageindex")))+1, pJumpStart)
					pJumpStart = Easp.IIF(pJumpStart<1,1,pJumpStart)
					j = 1
					For i = pJumpStart To cfg("pageindex")
						pJump = pJump & "<option value=""" & i & """" & Easp.IIF(i=cfg("pageindex")," selected=""selected""","") & ">" & i & "</option> "
						j = j + 1
					Next
					pJumpLong = Easp.IIF(cfg("pagecount")-cfg("pageindex")<pJumpLong, pJumpLong, pJumpLong + (pJumpLong-j)+1)
					pJumpEnd = Easp.IIF(cfg("pageindex")+pJumpLong>cfg("pagecount"), cfg("pagecount"), cfg("pageindex")+pJumpLong)
					For i = cfg("pageindex")+1 To pJumpEnd
						pJump = pJump & "<option value=""" & i & """>" & i & "</option> "
					Next
				End If
				pJump = pJump & "</select>"
		End Select
		tmpStr = Replace(tmpStr,"{recordcount}",cfg("recordcount"))
		tmpStr = Replace(tmpStr,"{pagecount}",cfg("pagecount"))
		tmpStr = Replace(tmpStr,"{pageindex}",cfg("pageindex"))
		tmpStr = Replace(tmpStr,"{pagesize}",cfg("pagesize"))
		tmpStr = Replace(tmpStr,"{list}",pList)
		tmpStr = Replace(tmpStr,"{liststart}",pListStart)
		tmpStr = Replace(tmpStr,"{listend}",pListEnd)
		tmpStr = Replace(tmpStr,"{first}",pFirst)
		tmpStr = Replace(tmpStr,"{prev}",pPrev)
		tmpStr = Replace(tmpStr,"{next}",pNext)
		tmpStr = Replace(tmpStr,"{last}",pLast)
		tmpStr = Replace(tmpStr,"{jump}",pJump)
		Set cfg = Nothing
		Pager = vbCrLf & tmpStr & vbCrLf
	End Function
	Public Sub SetPager(ByVal PagerName, ByVal PagerHtml, ByRef PagerConfig)
		If PagerName = "" Then PagerName = "default"
		If Easp.Has(PagerHtml) Then o_pageDic.item(PagerName&"_html") = PagerHtml
		If Easp.Has(PagerConfig) Then o_pageDic.item(PagerName&"_config") = PagerConfig
	End Sub
	Public Function GetPager(ByVal PagerName)
		If PagerName = "" Then PagerName = "default"
		GetPager = Pager(o_pageDic(PagerName&"_html"),o_pageDic(PagerName&"_config"))
	End Function
	Private Function GetCurrentPage()
		Dim rqParam, thisPage : thisPage = 1
		rqParam = Easp.Get(s_pageParam)
		If isNumeric(rqParam) Then
			If Int(rqParam) > 0 Then thisPage = Int(rqParam)
		End If
		GetCurrentPage = thisPage
	End Function
	Private Function GetRQ(pageNumer)
		GetRQ = Easp.ReplaceUrl(s_pageParam, pageNumer)
	End Function
End Class
%>