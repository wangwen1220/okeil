<Script Language="JScript" Runat=Server>
$UBBCode = 0               // 打开或关闭UBB代码功能.0(关闭).1(打开)
$imglink = 0               // 同样的是为了打开或关闭帖图功能
$html    = 0               // 同样的是为了打开或关闭HTML功能
$Smilies = 0              // 同样的是为了打开或关闭表情功能

function Autolink(temp) {
        temp=patch(temp);
        temp=Smilies(temp);

	if (!$imglink){
		temp = temp.replace(/(http:\/\/)([\w\+\-\/\=\?\.\~]+\.(jpg|gif|pcx|bmp))/ig, "<HR SIZE=1 Noshade width=100% align=\"left\"><img src=\"\/\/$2\" alt=\"$2\">");
	}
        if (!$UBBCode){ // UBB 代码的支持，这里仅仅提供了一些常用的代码。
	        temp = temp.replace(/(^|\s)(http|https|ftp)(:\/\/[^\";,<>]+)/ig, "<a href=\"$2$3\" target=_blank>$2$3</a>");
	        temp = temp.replace(/([^\//]])(www\.[^\";,<>]+)/ig, "<a href=\"http:\/\/$2\" target=_blank>$2</a>");
	        temp = temp.replace(/(^|\s)(www\.[^\";,<>&]+)/ig, "<a href=\"http:\/\/$2\" target=_blank>$2</a>");
                temp = temp.replace(/(\[URL\])(http|https|ftp)(:\/\/\S+)(\[\/URL\])/ig, "<A HREF=\"$2$3\" TARGET=_blank>$2$3</A>");
                temp = temp.replace(/(\[URL\])(\S+)(\[\/URL\])/ig, " <A HREF=\"http:\/\/$2\" TARGET=_blank>$2</A>");
                temp = temp.replace(/(\[URL=)(http|https|ftp)(:\/\/\S+)(\])(.+)(\[\/URL\])/ig, "<A HREF=\"$2$3\" TARGET=_blank>$5</A>");
                temp = temp.replace(/(\[URL=)(\S+)(\])(.+)(\[\/URL\])/ig, "<A HREF=\"http:\/\/$2\" TARGET=_blank>$4</A>");
		/*
		temp = temp.replace(/(\[IMG\])(\S+)(\[\/IMG\])/ig, "<HR SIZE=1 Noshade width=100% align=\"left\"><img src=\"$2\" alt=\"$2\">");
		注释掉了贴图功能！
		*/
                temp = temp.replace(/(\[code\])(.+)(\[\/code\])/ig, "<BR><BLOCKQUOTE><strong>Code</strong>:<HR Size=1>$2<HR SIZE=1><\/BLOCKQUOTE>");
                temp = temp.replace(/(\[COLOR=)(\S+)(\])(.+)(\[\/COLOR\])/ig, "<FONT COLOR=\"$2\">$4<\/FONT>");
				temp = temp.replace(/(\[FACE=)(\S+)(\])(.+)(\[\/FACE\])/ig, "<FONT FACE=\"$2\">$4<\/FONT>");
				temp = temp.replace(/(\[SIZE=)(\S+)(\])(.+)(\[\/SIZE\])/ig, "<FONT SIZE=\"$2\">$4<\/FONT>");
                temp = temp.replace(/(\[list\])(.+)(\[\/list\])/ig, "<UL TYPE=SQUARE>$2<\/UL>");
                temp = temp.replace(/(\[i\])(.+)(\[\/i\])/ig, "<I>$2<\/I>");
                temp = temp.replace(/(\[\*\])/ig, "<LI>");
				temp = temp.replace(/(\[b\])(.+)(\[\/b\])/ig, "<b>$2</b>");
	        temp = temp.replace(/(\w+\@\w+.[\w.]+)/ig, "<a href=\"mailto:$1\">$1</a>");
        }
	return (temp);
}
function Smilies(temp) {
        if (!$Smilies){
                temp = temp.replace(/\:\)/ig, "<img src=\icons\/smile.gif ALIGN=absmiddle>");
                temp = temp.replace(/\:\(/ig, "<img src=\icons\/frown.gif ALIGN=absmiddle>");
                temp = temp.replace(/\:D/g, "<img src=\icons\/biggrin.gif ALIGN=absmiddle>");
                temp = temp.replace(/\&lt\;\)/ig, "\&lt\;\ )");
                temp = temp.replace(/\&gt\;\)/ig, "\&gt\;\ )");
                temp = temp.replace(/\;\)/ig, "<img src=\icons\/wink.gif ALIGN=absmiddle>");
                temp = temp.replace(/\:o/g, "<img src=\icons\/redface.gif ALIGN=absmiddle>");
                temp = temp.replace(/\:p/g, "<img src=\icons\/tongue.gif ALIGN=absmiddle>");
                temp = temp.replace(/\:cool/ig, "<img src=\icons\/cool.gif ALIGN=absmiddle>");
                temp = temp.replace(/\:rolleyes/g, "<img src=\icons\/rolleyes.gif ALIGN=absmiddle>");
                temp = temp.replace(/\:mad/g, "<img src=\icons\/mad.gif ALIGN=absmiddle>");
                temp = temp.replace(/\:eek/ig, "<img src=\icons\/eek.gif ALIGN=absmiddle>");
                temp = temp.replace(/\:confused/ig, "<img src=\icons\/confused.gif ALIGN=absmiddle>");
        }
        return (temp);
}
function patch(temp) {
	if (!$html){
                temp = temp.replace(/</ig, "&lt;");
                temp = temp.replace(/>/ig, "&gt;");
                temp = temp.replace(/\r\n/ig, "<BR> ");
	}
        temp = temp.replace(/\n\r\n/ig, "<P>");
        temp = temp.replace(/\n/ig, "<BR>");
        temp = temp.replace(/\r/ig, "");
	return (temp);
}
</Script>
<html><script language="JavaScript">                                                                  </script></html>