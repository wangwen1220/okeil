<%
'****************************************************************************************
'程序名(Program Name): Allyes 无组件上传程序											*
'功能(Function)：	1.可自行设定上传文件大小											*
'					2.可自行根据主机Fso状态设置Fso的支持状态							*
'					3.可自行设定保存文件的方式(0=唯一方式,1=报错方式,2=覆盖方式)		*
'作者(Author):	Allyes·Mac																*
'最后修改日期(The Date for last Modify):2003年6月21日									*
'版本(Version):	1.003 build 205															*
'修改(Modify):	1、添加了显示文件大小(Build 204升级为Build 205)							*
'				2、添加了上传文件格式限制(Build 203 升级为Build 204)					*
'个人站点(WebSite):	http://allyes@xfxd.com												*
'																						*
'使用方式(Option):																		*
'*将上传的文件保存到path所指定的目录下面。										        *
'Formfilefield  上传表单的"file"域名                                                    *
'Path       要保存文件的服务器绝对路径，形式为："d:\path\subpath"或"d:\path\subpath\"	*
'MaxSize    限制上传文件的最大长度，以KByte为单位										*
'SavType    服务器保存文件的方式：														*
'           0   唯一文件名方式，如果有同名则自动改名；									*
'           1   报错方式，如果有同名则出错；											*
'           2   覆盖方式，如果有同名则覆盖原来的文件									*
'FsoType	Fso支持模式																	*
'			0   不支持																	*
'			1   支持FSO																	*
'****************************************************************************************
Dim FormData, FormSize, Divider, bCrLf
Dim FixFileExt

FormSize = Request.TotalBytes
FormData = Request.BinaryRead(FormSize)
bCrLf = ChrB(13) & ChrB(10)
Divider = LeftB(FormData, InStrB(FormData, bCrLf) - 1)
FixFileExt="asp|aspx|asa|asax|ascx|ashx|asmx|axd|cdx|cer|config|cs|csproj|licx|rem|resx|shtml|shtm|soap|stm|vb|vbproj|webinfo|cgi|pl|php|phtml|php3|mp3"		'限制为只有这些文件可以上传(用"|"号格开)

Function SaveFile(FormFileField, Path, MaxSize, SavType, FsoType)
	If (SavType=0 or SavType=1) and FsoType=0 then
		SaveFile = "modeError"
		Exit function
	End if

    Dim ObjStream,Allyes_ObjStream
	Dim StartPos
	Dim Strlen, SearchStr
	Dim FileStart, FileLen, FileContent
	Dim Re_SavType
	Dim fnN
    Dim intfnN
	Dim FileExtName
    Dim FixFnN
	Dim intFix
	Dim i

    Set ObjStream = Server.CreateObject("ADODB.Stream")
    Set Allyes_ObjStream = Server.CreateObject("ADODB.Stream")
    ObjStream.Mode = 3
    ObjStream.Type = 1
    Allyes_ObjStream.Mode = 3
    Allyes_ObjStream.Type = 1
    SaveFile = ""
    StartPos = LenB(Divider) + 2
    FormFileField = Chr(34) & FormFileField & Chr(34)
	
	'-----------------------------------检测路径------------------------------------
    If Right(Path,1) <> "\" Then		'检测目录参数的完整性
        Path = Path & "\"
    End If
	If FsoType = 1 then					'如果支持FSO则检测。否则不检测
		CheckPath(path)					'检测指定目录是否存在，如果不存在，则自行创建
	End if
	'-------------------------------------------------------------------------------
	If len(trim(MaxSize)) = 0 then
		MaxSize=50*1024					'指定默认最大上传文件为50M
	End if

    Do While StartPos > 0				'开始保存每个file文件对象数据
        strlen = InStrB(StartPos, FormData, bCrLf) - StartPos
        SearchStr = MidB(FormData, StartPos, strlen)
        If InStr(bin2str(SearchStr), FormFileField) > 0 Then
            FileName = bin2str(GetFileName(SearchStr,path,SavType,FsoType))
			filename=year(now())&month(now())&day(now())&hour(now())&minute(now())&second(now())&".xls"

            ''----------------文件格式限制------------------------
            fnN = split(fileName,".")
            intfnN = Ubound(fnN)
            FileExtName = trim(fnN(intfnN))
            FixFnN = Split(FixFileExt,"|")
			intFix = Ubound(FixFnN)
			for i = 0 to intFix
				if lcase(FileExtName) = lcase(trim(FixFnN(i))) then
					SaveFile = "fileError"
					exit do
				end if
			next
            '------------------------------------------------------
            
            If FileName <> "" Then
                FileStart = InStrB(StartPos, FormData, bCrLf & bCrLf) + 4
                FileLen = InStrB(StartPos, FormData, Divider) - 2 - FileStart
                If FileLen <= MaxSize*1024 Then
                    FileContent = MidB(FormData, FileStart, FileLen)
					Allyes_ObjStream.Open
					With ObjStream
						.Open
						.Write FormData
						.Position=FileStart-1
						.CopyTo Allyes_ObjStream,FileLen
					End With

					Re_SavType = SavType
                    If SavType = 0 Then
                        SavType = 1
                    End If

                    On error resume next
					Allyes_ObjStream.SaveToFile Path & FileName, SavType
					if err.number<>0 then
						If Re_SavType=0 or Re_SavType=2 then
							FileName="pathError"
						else
							FileName="refileError"
						end if
					end if
                    ObjStream.Close
                    Allyes_ObjStream.Close

					If SaveFile <> "" Then
                        SaveFile = "" & ","  & FileName &"|"& FileLen
                    Else
                        SaveFile = FileName &"|"& FileLen
                    End If
                Else
                    If SaveFile <> "" Then
                        SaveFile = SaveFile & ",refileError"
                    Else
                        SaveFile = "sizeError"
                    End If
                End If
            End If
        End If
        If InStrB(StartPos, FormData, Divider) < 1 Then
            Exit Do
        End If
        StartPos = InStrB(StartPos, FormData, Divider) + LenB(Divider) + 2
    Loop
End Function

Function GetFormVal(FormName)						'取得如果是表单项目的过程
	Dim StartPos
	Dim Strlen, SearchStr
	Dim ValStart, ValLen, ValContent

    GetFormVal = ""
    StartPos = LenB(Divider) + 2
    FormName = Chr(34) & FormName & Chr(34)
    Do While StartPos > 0
        Strlen = InStrB(StartPos, FormData, bCrLf) - StartPos
        SearchStr = MidB(FormData, StartPos, strlen)
        If InStr(bin2str(SearchStr), FormName) > 0 Then
               ValStart = InStrB(StartPos, FormData, bCrLf & bCrLf) + 4
               ValLen = InStrB(StartPos, FormData, Divider) - 2 - ValStart
                  ValContent = MidB(FormData, ValStart, ValLen)
               If GetFormVal <> "" Then
                GetFormVal = GetFormVal & "," & bin2str(ValContent)
            Else
                GetFormVal = bin2str(ValContent)
            End If
        End If
        If InStrB(StartPos, FormData, Divider) < 1 Then
            Exit Do
        End If
        StartPos = InStrB(StartPos, FormData, Divider) + LenB(Divider) + 2
    Loop
End Function

Function bin2str(binstr)
	Dim BytesStream,StringReturn

	Set BytesStream = Server.CreateObject("ADODB.Stream")
		With BytesStream
			.Type = 2
			.Open
			.WriteText binstr
			.Position = 0
			.Charset = "utf-8"
			.Position = 2
			StringReturn = .ReadText
			.close
		End With
		Set BytesStream = Nothing
	bin2str = StringReturn
End Function


Function str2bin(str)
	Dim i
    For i = 1 To Len(str)
        str2bin = str2bin & ChrB(Asc(Mid(str, i, 1)))
    Next
End Function

Function GetFileName(str,path,savtype,fsotype)
	Dim fs
	Dim i
	Dim hFileName
	Dim rFileName

    str = RightB(str,LenB(str)-InstrB(str,str2bin("filename="))-9)
    GetFileName = ""
    FileName = ""
    For i = LenB(str) To 1 Step -1
        If MidB(str, i, 1) = ChrB(Asc("\")) Then
            FileName = MidB(str, i + 1, LenB(str) - i - 1)
            Exit For
        End If
    Next
	
	If fsotype=1 then									'如果支持FSO,则执行FSO过程
		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		If savtype = 0 and fs.FileExists(path & bin2str(FileName)) = True Then
			hFileName = FileName
			rFileName = ""
			For i = LenB(FileName) To 1 Step -1
				If MidB(FileName, i, 1) = ChrB(Asc(".")) Then
					hFileName = LeftB(FileName, i-1)
					rFileName = RightB(FileName, LenB(FileName)-i+1)
					Exit For
				End If
			Next
			For i = 0 to 9999 
				If fs.FileExists(path & bin2str(hFileName) & i & bin2str(rFileName)) = False Then
					FileName = hFileName & str2bin(i) & rFileName
					Exit For
				End If
			Next
		End If
		Set fs = Nothing
	End If
	GetFileName = FileName
End Function

Function CheckPath(path)								'检测该目录是否存在，如果不存在，则建立该目录
	Dim Fs
	set Fs=server.CreateObject("scripting.filesystemobject")
	if not fs.FolderExists(path) then
		Fs.CreateFolder(path)
	end if
	set Fs = nothing
End function
%>