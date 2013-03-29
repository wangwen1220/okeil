<% Option Explicit 
Dim NoUpFileExt,NoUpFileTxt,UpFileText,NoUpFileText
'任何情况下都不允许上传的文件类型
NoUpFileExt   = "Asp|Asa|Exe|cer|aspx|php"
'上传文件时出现黑名单文件时的提示语
NoUpFileTxt = "当前站点已将该文件列入完全禁止上传文件黑名单！"
'防止上传jpg等木马检测自定义代码 以|分列 不能含有% >"
UpFileText = "Response|.write"
'检测出恶意代码后的提示
NoUpFileText = "上传失败:\n\n检测出您上传的文件含有恶意代码,我们已将您的IP记录!\n\n您的IP地址是:"&Request.ServerVariables("Remote_Addr")&"\n\n超过三次后我们将禁止您访问本站！"
%>
<!--#include file="Config.Asp"-->
<!--#include file="UpfileClass.Asp"-->
<%
Server.ScriptTimeOut = 1800
Dim sType, sStyleName, sCusDir
Dim sAllowExt, nAllowSize, sUploadDir, nUploadObject, nAutoDir, sBaseUrl, sContentPath, nsUploadDir
Dim sFileExt, sOriginalFileName, sSaveFileName, sPathFileName, nFileNum
Dim nSLTFlag, nSLTMinSize, nSLTOkSize, nSYWZFlag, sSYText, sSYFontColor, nSYFontSize, sSYFontName, sSYPicPath, nSLTSYObject, sSLTSYExt, nSYWZMinWidth, sSYShadowColor, nSYShadowOffset, nSYWZMinHeight, nSYWZPosition, nSYWZTextWidth, nSYWZTextHeight, nSYWZPaddingH, nSYWZPaddingV, nSYTPFlag, nSYTPMinWidth, nSYTPMinHeight, nSYTPPosition, nSYTPPaddingH, nSYTPPaddingV, nSYTPImageWidth, nSYTPImageHeight, nSYTPOpacity, nCusDirFlag
'2008根据5.5新加
Dim sSywzFlag,sSytpFlag,sSltFlag,IsRemoteSY,FielNames,SmallFileName,IsFileClass,IsMackSmallSy
'初始化上传参数
Call InitUpload()

Dim sAction
sAction = UCase(Trim(Request.QueryString("action")))

Call DoCreateNewDir()

Select Case sAction
Case "LOCAL"
	Call DoLocal()
Case "REMOTE"
	Call DoRemote()
Case "SAVE"
	Call DoSave()
End Select



Sub DoSave()
	Response.Write "<html><head><title>eWebEditor</title><meta http-equiv='Content-Type' content='text/html; charset=UTF-8'></head><body>"

	Select Case nUploadObject
	Case 1
		Call DoUpload_ASPUpload()
	Case 2
		Call DoUpload_SAFileUP()
	Case 3
		Call DoUpload_LyfUpload()
	Case Else
		Call DoUpload_Class()
	End Select


	Response.write sUploadDir
'Call CheckUploadFile(sUploadDir&sSaveFileName)
	Dim s_SmallImageFile, s_SmallImagePathFile, s_SmallImageScript
	s_SmallImageFile = getSmallImageFile(sSaveFileName)
	s_SmallImagePathFile = ""
	s_SmallImageScript = ""
	If makeImageSLT(sUploadDir, sSaveFileName, s_SmallImageFile) = True Then
	  If IsMackSmallSy = 1 Then Call makeImageSY(sUploadDir, s_SmallImageFile)
		Call makeImageSY(sUploadDir, sSaveFileName)
		s_SmallImagePathFile = sContentPath & s_SmallImageFile
		s_SmallImageScript = "try{obj.addUploadFile('" & sOriginalFileName & "', '" & s_SmallImageFile & "', '" & s_SmallImagePathFile & "');} catch(e){} "
	Else
		s_SmallImageFile = ""
		Call makeImageSY(sUploadDir, sSaveFileName)
	End If

	sPathFileName = sContentPath & sSaveFileName
	sOriginalFileName = Replace(sOriginalFileName, "'", "\'")
	sOriginalFileName = Replace(sOriginalFileName, """", "\""")
	Call OutScript("parent.UploadSaved('" & sPathFileName & "','" & s_SmallImagePathFile & "');var obj=parent.dialogArguments;if((!obj.eWebEditor)||(!obj.eWebEditor_Temp_HTML)||(!obj.eWebEditor_UploadForm)){obj=parent.dialogArguments.dialogArguments;} try{obj.addUploadFile('" & sOriginalFileName & "', '" & sSaveFileName & "', '" & sPathFileName & "');} catch(e){} " & s_SmallImageScript)
End Sub

Sub DoLocal()
	Select Case nUploadObject
	Case 1
		Call DoUpload_ASPUpload()
	Case 2
		Call DoUpload_SAFileUP()
	Case 3
		Call DoUpload_LyfUpload()
	Case Else
		Call DoUpload_Class()
	End Select
	
	sPathFileName = sContentPath & sSaveFileName
	Response.Write sPathFileName
End Sub
Sub makeImageSY(s_Path, s_File)
    IF IsRemoteSY = 1 and sAction = "REMOTE" then
	if nSYWZFlag > 0 then nSYWZFlag = 1
	if nSYTPFlag > 0 then nSYTPFlag = 1
	end if
	If isValidSLTSYExt(s_File) = False Then Exit Sub
	On Error Resume Next
	Dim nOriginalWidth, nOriginalHeight, posX, posY
	Dim oImage, oLogo

	Select Case nSLTSYObject
	Case 0
		If IsObjInstalled("Persits.Jpeg") = False Then Exit Sub
		Set oImage = Server.CreateObject("Persits.Jpeg")

		If nSYWZFlag = 1 Then
		
			oImage.Open Server.Mappath(s_Path & s_File)
			nOriginalWidth = oImage.OriginalWidth
			nOriginalHeight = oImage.OriginalHeight
			If nOriginalWidth<nSYWZMinWidth Or nOriginalHeight<nSYWZMinHeight Then Exit Sub
			posX = getSYPosX(nSYWZPosition, nOriginalWidth, nSYWZTextWidth+nSYShadowOffset, nSYWZPaddingH)
			posY = getSYPosY(nSYWZPosition, nOriginalHeight, nSYWZTextHeight+nSYShadowOffset, nSYWZPaddingV)

			oImage.Canvas.Font.Color = Clng("&H" & sSYFontColor)
			oImage.Canvas.Font.Family = sSYFontName
			'oImage.Canvas.Font.Bold = True
			oImage.Canvas.Font.Size = nSYFontSize
			oImage.Canvas.Font.ShadowColor = Clng("&H" & sSYShadowColor)
			oImage.Canvas.Font.ShadowXOffset = nSYShadowOffset
			oImage.Canvas.Font.ShadowYOffset = nSYShadowOffset
			oImage.Canvas.Print posX, posY, sSYText
			oImage.Save Server.Mappath(s_Path & s_File)
		End If
		If nSYTPFlag = 1 Then
			oImage.Open Server.Mappath(s_Path & s_File)
			nOriginalWidth = oImage.OriginalWidth
			nOriginalHeight = oImage.OriginalHeight
			If nOriginalWidth<nSYTPMinWidth Or nOriginalHeight<nSYTPMinHeight Then Exit Sub
			posX = getSYPosX(nSYTPPosition, nOriginalWidth, nSYTPImageWidth, nSYTPPaddingH)
			posY = getSYPosY(nSYTPPosition, nOriginalHeight, nSYTPImageHeight, nSYTPPaddingV)

			Set oLogo = Server.CreateObject("Persits.Jpeg")
			oLogo.Open Server.Mappath(sSYPicPath)
			oImage.DrawImage posX, posY, oLogo, nSYTPOpacity, &HFFFFFF
			oImage.Save Server.Mappath(s_Path & s_File)
			Set oLogo = Nothing
		End If
		Set oImage = Nothing
	Case Else

	End Select

End Sub

Function getSYPosX(posFlag, originalW, syW, paddingH)
	Select Case posFlag
	Case 1, 2, 3
		getSYPosX = paddingH
	Case 4, 5, 6
		getSYPosX = (originalW - syW) \ 2
	Case 7, 8, 9
		getSYPosX = originalW - paddingH - syW
	End Select
End Function

Function getSYPosY(posFlag, originalH, syH, paddingV)
	Select Case posFlag
	Case 1, 4, 7
		getSYPosY = paddingV
	Case 2, 5, 8
		getSYPosY = (originalH - syH) \ 2
	Case 3, 6, 9
		getSYPosY = originalH - paddingV - syH
	End Select
End Function

Function makeImageSLT(s_Path, s_File, s_SmallFile)
	makeImageSLT = False
	If nSLTFlag <> 1 Then Exit Function
	If isValidSLTSYExt(s_File) = False Then Exit Function

	Dim nOriginalWidth, nOriginalHeight, nWidth, nHeight
	Dim oImage

	Select Case nSLTSYObject
	Case 0
		If IsObjInstalled("Persits.Jpeg") = False Then Exit Function
		Set oImage = Server.CreateObject("Persits.Jpeg")
		oImage.Open Server.Mappath(s_Path & s_File)
		nOriginalWidth = oImage.OriginalWidth
		nOriginalHeight = oImage.OriginalHeight
		If nOriginalWidth < nSLTMinSize And nOriginalHeight < nSLTMinSize Then Exit Function
		If nOriginalWidth > nOriginalHeight Then
			nWidth = nSLTOkSize
			nHeight = (nSLTOkSize / nOriginalWidth) * nOriginalHeight
		Else
			nHeight = nSLTOkSize
			nWidth = (nSLTOkSize / nOriginalHeight) * nOriginalWidth
		End If
		oImage.Width = nWidth
		oImage.Height = nHeight
		oImage.Save Server.Mappath(s_Path & s_SmallFile)
		Set oImage = Nothing
	Case Else

	End Select

	makeImageSLT = True
End Function

Function isValidSLTSYExt(s_File)
	Dim b, i, aExt, sExt
	b = False
	sExt = LCase(Mid(s_File, InstrRev(s_File, ".")+1))
	aExt = Split(LCase(sSLTSYExt), "|")
	For i = 0 To UBound(aExt)
		If aExt(i) = sExt Then
			b = True
			Exit For
		End If
	Next
	isValidSLTSYExt = b
End Function

Function getSmallImageFile(s_File)
	Dim n
	n = InstrRev(s_File, ".")
	if SmallFileName <> "" then
	getSmallImageFile = Left(s_File, n-1) & SmallFileName & "." & Mid(s_File, n+1)
	Else
	getSmallImageFile = Left(s_File, n-1) & "_s." & Mid(s_File, n+1)
	End IF
End Function


Sub DoRemote()
	Dim sContent, i
	For i = 1 To Request.Form("eWebEditor_UploadText").Count 
		sContent = sContent & Request.Form("eWebEditor_UploadText")(i) 
	Next

	If sAllowExt <> "" Then
		sContent = ReplaceRemoteUrl(sContent, sAllowExt)
	End If

	Response.Write "<html><head><title>eWebEditor</title><meta http-equiv='Content-Type' content='text/html; charset=UTF-8'></head><body>" & _
		"<input type=hidden id=UploadText value=""" & inHTML(sContent) & """>" & _
		"</body></html>"

	Call OutScriptNoBack("parent.setHTML(UploadText.value);try{parent.addUploadFile('" & sOriginalFileName & "', '" & sSaveFileName & "', '" & sPathFileName & "');} catch(e){} parent.remoteUploadOK();")

End Sub




Sub DoUpload_LyfUpload()
	On Error Resume Next
	Dim oUpload, sResult, sOriginalFile
	Set oUpload = Server.CreateObject("LyfUpload.UploadFile")
	oUpload.CodePage = 65001
	oUpload.ExtName = Replace(sAllowExt, "|", ",")
	oUpload.MaxSize = nAllowSize*2048
	sOriginalFile = oUpload.Request("originalfile")
	sOriginalFileName = Mid(sOriginalFile, InStrRev(sOriginalFile, "\") + 1)
	sFileExt = LCase(Mid(sOriginalFileName, InStrRev(sOriginalFileName, ".") + 1))
	Call CheckValidExt(sFileExt)
	sSaveFileName = GetRndFileName(sFileExt)
	sResult = oUpload.SaveFile("uploadfile", Server.Mappath(sUploadDir), True, sSaveFileName)

	Select Case sResult
	Case "0"
		Call OutScript("parent.UploadError('size')")
	Case ""
		Call OutScript("parent.UploadError('file')")
	Case "1"
		Call OutScript("parent.UploadError('ext')")
	End Select
	
	Set oUpload = Nothing
End Sub

Sub DoUpload_SAFileUp()
	On Error Resume Next
	Dim oFileUp
	Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")
	'oFileUp.MaxBytes = nAllowSize*1024
	oFileUp.CodePage = 65001
	oFileUp.Path = Server.MapPath(sUploadDir)

	If oFileUp.Form("uploadfile").TotalBytes > nAllowSize*2048 Then
		Err.Clear
		Call OutScript("parent.UploadError('size')")
	End If
	If oFileUp.Form("uploadfile").IsEmpty Then
		Call OutScript("parent.UploadError('file')")
	End If

	Dim sShortFileName
	'sShortFileName = oFileUp.Form("uploadfile").ShortFileName
	sShortFileName = Mid(oFileUp.Form("uploadfile").UserFilename, InstrRev(oFileUp.Form("uploadfile").UserFilename, "\") + 1)

	sFileExt = LCase(Mid(sShortFileName, InStrRev(sShortFileName, ".") + 1))
	Call CheckValidExt(sFileExt)
	sOriginalFileName = sShortFileName
	sSaveFileName = GetRndFileName(sFileExt)
	oFileUp.Form("uploadfile").SaveAs Server.Mappath(sUploadDir & sSaveFileName)
	
	Set oFileUp = Nothing
End Sub



Sub DoUpload_ASPUpload()
	On Error Resume Next
	Dim oUpload, oFile, nCount
	Set oUpload = Server.CreateObject("Persits.Upload")
	oUpload.CodePage = 65001
	oUpload.SetMaxSize nAllowSize*2048, True
	nCount = oUpload.Save
	'nCount = oUpload.SaveToMemory

	If nCount < 1 Then
		Call OutScript("parent.UploadError('file')")
	End If
	If Err.Number = 8 Then
		Err.Clear
		Call OutScript("parent.UploadError('size')")
	End If
	
	Set oFile = oUpload.Files("uploadfile")
	sFileExt = LCase(Mid(oFile.Ext, 2))
	Call CheckValidExt(sFileExt)
	sOriginalFileName = oFile.FileName
	sSaveFileName = GetRndFileName(sFileExt)
	oFile.SaveAs Server.Mappath(sUploadDir & sSaveFileName)

	Set oFile = Nothing
	Set oUpload = Nothing
End Sub

Sub DoUpload_Class()
	On Error Resume Next
	Dim oUpload, oFile
	Set oUpload = New upfile_class

	oUpload.GetData nAllowSize*2048

	If oUpload.Err > 0 Then
		Select Case oUpload.Err
		Case 1
			Call OutScript("parent.UploadError('file')")
		Case 2
			Call OutScript("parent.UploadError('size')")
		End Select
	End If

	Set oFile = oUpload.File("uploadfile")
	sFileExt = LCase(oFile.FileExt)
	Call CheckValidExt(sFileExt)
	sOriginalFileName = oFile.FileName
	sSaveFileName = GetRndFileName(sFileExt)

	Dim str_Mappath
	str_Mappath = Server.Mappath(sUploadDir & sSaveFileName)
	sFileExt = LCase(Mid(str_Mappath, InstrRev(str_Mappath, ".") + 1))
	Call CheckValidExt(sFileExt)

	oFile.SaveToFile str_Mappath
	
	Set oFile = Nothing
	Set oUpload = Nothing
End Sub

Function GetRndFileName(sExt)
	Dim sRnd
	Randomize
	sRnd = Int(900 * Rnd) + 100
	GetRndFileName = FielNames & FormatTime(Now(), 5) & sRnd & "." & sExt
End Function

Sub OutScript(str)
	If sAction <> "LOCAL" Then
		Response.Write "<script language=javascript>" & str & ";history.back()</script>"
	End If
	Response.End
End Sub

Sub OutScriptNoBack(str)
	Response.Write "<script language=javascript>" & str & "</script>"
	Response.End
End Sub
Sub OutScriptCloseWindows(str)
	Response.Write "<script language=javascript>" & str & ";window.close();</script>"
	Response.End
End Sub

Sub CheckUploadFile(FileNames)
Dim Snum,sTextAll,MyFile,MyText,i,filedel
		Set MyFile = server.CreateObject("Scripting.FileSystemObject")
		IF Not MyFile.FileExists(Server.Mappath(FileNames)) then OutScriptCloseWindows("alert('上传未成功:可能因为服务器上的防火设置或杀毒软件关系！页面即将关闭!')")
		Set MyText = MyFile.OpenTextFile(Server.mappath(FileNames)) '读取文本文件
			 sTextAll = lcase(MyText.ReadAll())
			 MyText.close
		Set MyFile = nothing
        UpFileText=UpFileText&"|.unescape|.getfolder|.createfolder|.deletefolder|.createdirectory|.deletedirectory|"
        UpFileText=UpFileText&".saveas|wscript.shell|script.encode|server.|.createobject|execute|activexobject|language="
        snum = split(Trim(UpFileText),"|")
		For i=0 to ubound(snum)
		    if trim(snum(i)) <> "" then
			 if instr(sTextAll,lcase(Trim(snum(i)))) then
			   set filedel = server.CreateObject("Scripting.FileSystemObject")
			   filedel.deletefile Server.mappath(FileNames)
			   set filedel = nothing
			  Call OutScriptCloseWindows("alert('"&NoUpFileText&"')")
			 end if
			End IF
		Next
End Sub

Sub CheckValidExt(sExt)
	Dim b, i, aExt,Check
	b = False
	Check = False
	aExt = Split(sAllowExt, "|")
	For i = 0 To UBound(aExt)
		If LCase(aExt(i)) = sExt Then
			b = True
			Exit For
		End If
	Next
	'检测任何情况下都不允许上传的文件类型,文件类型在本文件顶部设置
	Dim Ni, NaExt
	NaExt = Split(NoUpFileExt,"|")
	For Ni = 0 To UBound(NaExt)
		If LCase(NaExt(Ni)) = sExt Then
			Check = True
			Exit For
		End If
	Next
	IF Check = True Then
	   Call OutScriptCloseWindows("alert('"&NoUpFileTxt&"')")
	Else
	If b = False Then
		Call OutScript("parent.UploadError('ext')")
	End If
	End IF
End Sub

Sub InitUpload()
	sType = UCase(Trim(Request.QueryString("type")))
	sStyleName = Trim(Request.QueryString("style"))
	sCusDir = Trim(Request.QueryString("cusdir"))
	sSywzFlag = Clng(Request.QueryString("SYWZFlag"))
	sSytpFlag = Clng(Request.QueryString("sytpFlag"))
	sSltFlag = Clng(Request.QueryString("SltFlag")) 
	sCusDir = Replace(sCusDir, "\", "/")
	If Left(sCusDir, 1) = "/" Or Left(sCusDir, 1) = "." Or Right(sCusDir, 1) = "." Or InStr(sCusDir, "./") > 0 Or InStr(sCusDir, "/.") > 0 Or InStr(sCusDir, "//") > 0 Then
		sCusDir = ""
	End If
	Dim i, aStyleConfig, bValidStyle
	bValidStyle = False
	For i = 1 To Ubound(aStyle)
		aStyleConfig = Split(aStyle(i), "|||")
		If Lcase(sStyleName) = Lcase(aStyleConfig(0)) Then
			bValidStyle = True
			Exit For
		End If
	Next

	If bValidStyle = False Then
		OutScript("parent.UploadError('style')")
	End If

	sBaseUrl = aStyleConfig(19)
	nUploadObject = Clng(aStyleConfig(20))
	nAutoDir = CLng(aStyleConfig(21))

	sUploadDir = aStyleConfig(3)
'	If Left(sUploadDir, 1) <> "/" Then
'		sUploadDir = "../" & sUploadDir
'	End If
	nSLTFlag = Clng(aStyleConfig(29))
	nSYWZFlag = Clng(aStyleConfig(32))
    nSYTPFlag = Clng(aStyleConfig(52))
	IF (nSLTFlag = 2 or nSLTFlag = 3) and sSltFlag = 1 Then nSLTFlag = 1
	IF (nSYWZFlag = 2 or nSYWZFlag = 3) and sSywzFlag = 1 Then nSYWZFlag = 1
    IF (nSYTPFlag = 2 or nSYTPFlag = 3) and sSytpFlag = 1 Then nSYTPFlag = 1
	nSLTMinSize = Clng(aStyleConfig(30))
	nSLTOkSize = Clng(aStyleConfig(31))
	sSYText = aStyleConfig(33)
	sSYFontColor = aStyleConfig(34)
	nSYFontSize = Clng(aStyleConfig(35))
	sSYFontName = aStyleConfig(36)
	sSYPicPath = aStyleConfig(37)
	nSLTSYObject = Clng(aStyleConfig(38))
	sSLTSYExt = aStyleConfig(39)
	nSYWZMinWidth = Clng(aStyleConfig(40))
	sSYShadowColor = aStyleConfig(41)
	nSYShadowOffset = Clng(aStyleConfig(42))
	nSYWZMinHeight = Clng(aStyleConfig(46))
	nSYWZPosition = Clng(aStyleConfig(47))
	nSYWZTextWidth = Clng(aStyleConfig(48))
	nSYWZTextHeight = Clng(aStyleConfig(49))
	nSYWZPaddingH = Clng(aStyleConfig(50))
	nSYWZPaddingV = Clng(aStyleConfig(51))
	nSYTPMinWidth = Clng(aStyleConfig(53))
	nSYTPMinHeight = Clng(aStyleConfig(54))
	nSYTPPosition = Clng(aStyleConfig(55))
	nSYTPPaddingH = Clng(aStyleConfig(56))
	nSYTPPaddingV = Clng(aStyleConfig(57))
	nSYTPImageWidth = Clng(aStyleConfig(58))
	nSYTPImageHeight = Clng(aStyleConfig(59))
	nSYTPOpacity = CDbl(aStyleConfig(60))
	nCusDirFlag = Clng(aStyleConfig(61))
	IsFileClass = Clng(aStyleConfig(67))
	IsRemoteSY = Clng(aStyleConfig(68))
	FielNames = aStyleConfig(69)
	SmallFileName = aStyleConfig(70)
	IsMackSmallSy = Clng(aStyleConfig(71))
	Select Case sType
	Case "REMOTE"
		sAllowExt = aStyleConfig(10)
		nAllowSize = Clng(aStyleConfig(15))
		IF IsFileClass = 1 Then sUploadDir = sUploadDir & "Remote/"
	Case "FILE"
		sAllowExt = aStyleConfig(6)
		nAllowSize = Clng(aStyleConfig(11))
		IF IsFileClass = 1 Then sUploadDir = sUploadDir & "Files/"
	Case "MEDIA"
		sAllowExt = aStyleConfig(9)
		nAllowSize = Clng(aStyleConfig(14))
		IF IsFileClass = 1 Then sUploadDir = sUploadDir & "Media/"
	Case "FLASH"
		sAllowExt = aStyleConfig(7)
		nAllowSize = Clng(aStyleConfig(12))
		IF IsFileClass = 1 Then sUploadDir = sUploadDir & "Flash/"
	Case "LOCAL"
		sAllowExt = aStyleConfig(44)
		nAllowSize = Clng(aStyleConfig(45))
		IF IsFileClass = 1 Then sUploadDir = sUploadDir & "Local/"
	Case Else
		sAllowExt = aStyleConfig(8)
		nAllowSize = Clng(aStyleConfig(13))
		IF IsFileClass = 1 Then sUploadDir = sUploadDir & "Image/"
	End Select
	Select Case sBaseUrl
	Case "0"
		sContentPath = aStyleConfig(23)
	Case "1"
		sContentPath = RelativePath2RootPath(sUploadDir)
	Case "2"
		sContentPath = RootPath2DomainPath(RelativePath2RootPath(sUploadDir))
	End Select

End Sub
Sub DoCreateNewDir()
    Dim a, i
	If nCusDirFlag = 1 Then
		a = Split(sCusDir, "/")
		For i = 0 To UBound(a)
			If a(i) <> "" Then
				Call CreateFolder(a(i))
			End If
		Next
	End If
	Dim s_DateDir
	Select Case nAutoDir
	Case 1
		s_DateDir = Left(FormatTime(Now(), 2), 4) REM:2008-09-04
	Case 2
		s_DateDir = Left(FormatTime(Now(), 2), 7)
	Case 3
		s_DateDir = Left(FormatTime(Now(), 2), 10)
	Case Else
		s_DateDir = ""
	End Select
		Call CreateFolder(s_DateDir)
End Sub

Sub CreateFolder(s_Folder)

	If IsObjInstalled("Scripting.FileSystemObject") = False Then
		Exit Sub
	End If
	if s_Folder <> "" then 
	sUploadDir = sUploadDir & s_Folder & "/"
	sContentPath = sContentPath & s_Folder & "/"
	end if
	nsUploadDir = sUploadDir
         On Error Resume Next
         Dim astrPath, ulngPath, i, strTmpPath
         Dim objFSO, strPath
         strPath = Server.Mappath(nsUploadDir)
         strPath=Replace(strPath, "\", "/") 
         Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
         If objFSO.FolderExists(strPath) Then Exit Sub
         astrPath = Split(strPath, "/")
         ulngPath = UBound(astrPath)
         strTmpPath = ""
         For i = 0 To ulngPath
		 IF astrPath(i) <> "" Then
                 strTmpPath = strTmpPath & astrPath(i) & "/"
                 If Not objFSO.FolderExists(strTmpPath) Then
                 objFSO.CreateFolder(strTmpPath)
                 End If
		 End IF
         Next
         Set objFSO = Nothing
End Sub

Function RelativePath2RootPath(url)
	Dim sTempUrl
	sTempUrl = url
	If Left(sTempUrl, 1) = "/" Then
		RelativePath2RootPath = sTempUrl
		Exit Function
	End If
	
	Dim sWebEditorPath
	sWebEditorPath = Request.ServerVariables("SCRIPT_NAME")
	sWebEditorPath = Left(sWebEditorPath, InstrRev(sWebEditorPath, "/") - 1)
	Do While Left(sTempUrl, 3) = "../"
		sTempUrl = Mid(sTempUrl, 4)
		sWebEditorPath = Left(sWebEditorPath, InstrRev(sWebEditorPath, "/") - 1)
	Loop
	RelativePath2RootPath = sWebEditorPath & "/" & sTempUrl
End Function

Function RootPath2DomainPath(url)
	Dim sHost, sPort
	sHost = Split(Request.ServerVariables("SERVER_PROTOCOL"), "/")(0) & "://" & Request.ServerVariables("HTTP_HOST")
	sPort = Request.ServerVariables("SERVER_PORT")
	If sPort <> "80" Then
		sHost = sHost & ":" & sPort
	End If
	RootPath2DomainPath = sHost & url
End Function

Function ReplaceRemoteUrl(sHTML, sExt)
	Dim s_Content
	s_Content = sHTML
	If IsObjInstalled("Microsoft.XMLHTTP") = False Or nAllowSize <= 0 then
		ReplaceRemoteUrl = s_Content
		Exit Function
	End If
		
	Dim re, RemoteFile, RemoteFileurl, SaveFileName, SaveFileType
	Set re = new RegExp
	re.IgnoreCase  = True
	re.Global = True
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}(([A-Za-z0-9_-])+[.]){1,}([A-Za-z0-9]{1,5})\/(\S+\.(" & sExt & ")))"

	Set RemoteFile = re.Execute(s_Content)
	
	Dim a_RemoteUrl(), n, i, bRepeat
	n = 0
	' to no repeat array
	
	For Each RemoteFileurl in RemoteFile
		If n = 0 Then
			n = n + 1
			Redim a_RemoteUrl(n)
			a_RemoteUrl(n) = RemoteFileurl
		Else
			bRepeat = False
			For i = 1 To UBound(a_RemoteUrl)
				If UCase(RemoteFileurl) = UCase(a_RemoteUrl(i)) Then
					bRepeat = True
					Exit For
				End If
			Next
			If bRepeat = False Then
				n = n + 1
				Redim Preserve a_RemoteUrl(n)
				a_RemoteUrl(n) = RemoteFileurl
			End If
		End If		
	Next
	' start replace
	nFileNum = 0
	For i = 1 To n
		SaveFileType = Mid(a_RemoteUrl(i), InstrRev(a_RemoteUrl(i), ".") + 1)

		SaveFileName = GetRndFileName(SaveFileType)

		If SaveRemoteFile(SaveFileName, a_RemoteUrl(i)) = True Then
		  IF IsRemoteSY = 1 Then Call makeImageSY(sUploadDir, SaveFileName)
			nFileNum = nFileNum + 1
			If nFileNum > 1 Then
				sOriginalFileName = sOriginalFileName & "|"
				sSaveFileName = sSaveFileName & "|"
				sPathFileName = sPathFileName & "|"
			End If
			sOriginalFileName = sOriginalFileName & Mid(a_RemoteUrl(i), InstrRev(a_RemoteUrl(i), "/") + 1)
			sSaveFileName = sSaveFileName & SaveFileName
			sPathFileName = sPathFileName & sContentPath & SaveFileName
			s_Content = Replace(s_Content, a_RemoteUrl(i), sContentPath & SaveFileName, 1, -1, 1)
		End If
	Next

	ReplaceRemoteUrl = s_Content
End Function

Function SaveRemoteFile(s_LocalFileName, s_RemoteFileUrl)
	Dim Ads, Retrieval, GetRemoteData
	Dim bError
	bError = False
	SaveRemoteFile = False
	On Error Resume Next
	Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
	With Retrieval
		.Open "Get", s_RemoteFileUrl, False, "", ""
		.Send
		GetRemoteData = .ResponseBody
	End With
	Set Retrieval = Nothing

	If LenB(GetRemoteData) > nAllowSize*2048 Then
		bError = True
	Else
		Set Ads = Server.CreateObject("Adodb." & "Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile Server.MapPath(sUploadDir & s_LocalFileName), 2
			.Cancel()
			.Close()
		End With
		Set Ads=nothing
	End If

	If Err.Number = 0 And bError = False Then
		SaveRemoteFile = True
	Else
		Err.Clear
	End If
End Function

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

Function inHTML(str)
	Dim sTemp
	sTemp = str
	inHTML = ""
	If IsNull(sTemp) = True Then
		Exit Function
	End If
	sTemp = Replace(sTemp, "&", "&amp;")
	sTemp = Replace(sTemp, "<", "&lt;")
	sTemp = Replace(sTemp, ">", "&gt;")
	sTemp = Replace(sTemp, Chr(34), "&quot;")
	inHTML = sTemp
End Function

Function FormatTime(s_Time, n_Flag)
	Dim y, m, d, h, mi, s
	FormatTime = ""
	If IsDate(s_Time) = False Then Exit Function
	y = cstr(year(s_Time))
	m = cstr(month(s_Time))
	If len(m) = 1 Then m = "0" & m
	d = cstr(day(s_Time))
	If len(d) = 1 Then d = "0" & d
	h = cstr(hour(s_Time))
	If len(h) = 1 Then h = "0" & h
	mi = cstr(minute(s_Time))
	If len(mi) = 1 Then mi = "0" & mi
	s = cstr(second(s_Time))
	If len(s) = 1 Then s = "0" & s
	Select Case n_Flag
	Case 1
		' yyyy-mm-dd hh:mm:ss
		FormatTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
	Case 2
		' yyyy-mm-dd
		FormatTime = y & "-" & m & "-" & d
	Case 3
		' hh:mm:ss
		FormatTime = h & ":" & mi & ":" & s
	Case 4
		' yyyymmdd
		FormatTime = y & m & d
	Case 5
		' yyyymmddhhmmss
		FormatTime = y & m & d & h & mi & s
	End Select
End Function
%>
