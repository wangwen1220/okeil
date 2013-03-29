<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>CreateEasy</title>
<script type="text/javascript" src="SkinS/js/buttons.js"></script>
<script type="text/javascript" src="SkinS/js/ewebeditor.js"></script>

<script type="text/javascript">
LoadScript("Skins/style/" + myParam.StyleName + ".js");
//alert(myParam.StyleName);
</script>

<script type="text/javascript">
lang.Init();
LoadScript("Skins/Language/" + lang.ActiveLanguage + ".js");
//alert(lang.ActiveLanguage);
</script>

<script type="text/javascript">
myEditor_Init();
document.write ("<link href='" + myEditor.RootPath + "/Skins/Language/" + lang.ActiveLanguage + ".css' type='text/css' rel='stylesheet'>");
document.write ("<link href='" + myEditor.RootPath + "/Skins/Images_Css/" + config.Skin + "/editor.css' type='text/css' rel='stylesheet'>");
//alert(myEditor.RootPath);
//alert(config.Skin);
</script>

</head>
<body scrolling=no>

<script type="text/javascript">
	document.write("<table id=eWebEditor_Layout border=0 cellpadding=0 cellspacing=0 width='100%' height='100%'>");document.write("<tr id='eWebEditor_ToolarTR' style='display:none'><td>");showToolbar();document.write("</td></tr>");document.write("<tr><td height='100%'>");document.write("	<input type='hidden' ID='ContentEdit' value=''>");document.write("	<input type='hidden' ID='ModeEdit' value=''>");document.write("	<input type='hidden' ID='ContentLoad' value=''>");document.write("	<input type='hidden' ID='ContentFlag' value='0'>");document.write("	<iframe class='Composition' ID='eWebEditor' style='width:100%;height:100%' marginwidth=1 marginheight=1 scrolling='yes' frameborder='0' src='dialog/blank.htm'></iframe>");document.write("</td></tr>");if (config.StateFlag){document.write("<tr><td class=SB>");document.write("	<TABLE border='0' cellPadding='0' cellSpacing='0' width='100%' class=SB>");document.write("	<TR valign=middle>");document.write("	<td>");document.write("		<table border=0 cellpadding=0 cellspacing=0 class=SB_Mode>");document.write("		<tr>");document.write("		<td class=SB_Mode_Left></td>");if (config.SBCode){document.write("		<td class=SB_Mode_BtnOff id=eWebEditor_CODE onclick=\"setMode('CODE')\" unselectable=on><table border=0 cellpadding=0 cellspacing=0><tr><td class=SB_Mode_Btn_Img>" + getBtnImgHTML("ModeCode") + "</td><td class=SB_Mode_Btn_Text>" + lang["StatusModeCode"] + "</td></tr></table></td>");document.write("		<td class=SB_Mode_Sep></td>");}if (config.SBEdit){document.write("		<td class=SB_Mode_BtnOff id=eWebEditor_EDIT onclick=\"setMode('EDIT')\" unselectable=on><table border=0 cellpadding=0 cellspacing=0><tr><td class=SB_Mode_Btn_Img>" + getBtnImgHTML("ModeEdit") + "</td><td class=SB_Mode_Btn_Text>" + lang["StatusModeEdit"] + "</td></tr></table></td>");document.write("		<td class=SB_Mode_Sep></td>");}if (config.SBText){document.write("		<td class=SB_Mode_BtnOff id=eWebEditor_TEXT onclick=\"setMode('TEXT')\" unselectable=on><table border=0 cellpadding=0 cellspacing=0><tr><td class=SB_Mode_Btn_Img>" + getBtnImgHTML("ModeText") + "</td><td class=SB_Mode_Btn_Text>" + lang["StatusModeText"] + "</td></tr></table></td>");document.write("		<td class=SB_Mode_Sep></td>");}if (config.SBView){document.write("		<td class=SB_Mode_BtnOff id=eWebEditor_VIEW onclick=\"setMode('VIEW')\" unselectable=on><table border=0 cellpadding=0 cellspacing=0><tr><td class=SB_Mode_Btn_Img>" + getBtnImgHTML("ModeView") + "</td><td class=SB_Mode_Btn_Text>" + lang["StatusModeView"] + "</td></tr></table></td>");}document.write("		</tr>");document.write("		</table>");document.write("	</td>");if (myParam.FullScreen!="1"){document.write("	<td align=right>");document.write("		<table border=0 cellpadding=0 cellspacing=0 class=SB_Size>");document.write("		<tr>");document.write("		<td class=SB_Size_Btn onclick='sizeChange(300)' title='"+lang["SizePlus"]+"'>"+getBtnImgHTML("SizePlus")+"</td>");document.write("		<td class=SB_Size_Sep></td>");document.write("		<td class=SB_Size_Btn onclick='sizeChange(-300)' title='"+lang["SizeMinus"]+"'>"+getBtnImgHTML("SizeMinus")+"</td>");document.write("		<td class=SB_Size_Right></td>");document.write("		</tr>");document.write("		</table>");document.write("	</td>");}document.write("	</TR>");document.write("	</Table>");document.write("</td></tr>");}document.write("</table>");document.write("<div id='eWebEditor_Temp_HTML' style='VISIBILITY: hidden; OVERFLOW: hidden; POSITION: absolute; WIDTH: 1px; HEIGHT: 1px'></div>");document.write("<div style='position:absolute;display:none'>");document.write("<form id='eWebEditor_UploadForm' action='upload." + config.ServerExt + "?action=remote&type=remote&style=" + myParam.StyleName + "&language=" + lang.ActiveLanguage + "&cusdir=" + myParam.CusDir + "' method='post' target='eWebEditor_UploadTarget'>");document.write("<input type='hidden' name='eWebEditor_UploadText'>");document.write("</form>");document.write("<iframe name='eWebEditor_UploadTarget' width=0 height=0 src='dialog/blank.htm'></iframe>");document.write("</div>");document.write("<div id=divProcessing style='width:200px;height:30px;position:absolute;display:none'>");document.write("<table border=0 cellpadding=0 cellspacing=1 bgcolor='#000000' width='100%' height='100%'><tr><td bgcolor=#3A6EA5><marquee id='msgProcessing' align='middle' behavior='alternate' scrollamount='5' style='font-size:9pt;color:#ffffff'></marquee></td></tr></table>");document.write("</div>");
</script>

</body>
</html>