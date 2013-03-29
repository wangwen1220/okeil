var swf_width=780
var swf_height=248
var files='Upfiles/Prod_B/banner10.jpg|Upfiles/Prod_B/banner20.jpg|Upfiles/Prod_B/banner30.jpg|Upfiles/Prod_B/banner40.jpg|Upfiles/Prod_B/banner50.jpg'
var links='gotosp.asp?p=1|gotosp.asp?p=2|gotosp.asp?p=3|gotosp.asp?p=4|gotosp.asp?p=5'
var texts=''


document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
document.write('<param name="movie" value="images/bcastr3.swf"><param name="quality" value="high">');
document.write('<param name="menu" value="true"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config=0x66cc00:文字颜色|2:文字位置|0xeeeeee:文字背景颜色|00:文字背景透明度|0xffffff:按键文字颜色|0x66cc00:按键默认颜色|0x000033:按键当前颜色|4:自动播放时间(秒)|0:图片过渡效果|1:是否显示按钮|_blank:打开窗口">');
document.write('<embed src="images/bcastr3.swf" wmode="opaque" FlashVars="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'& menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'&bcastr_config=0x66cc00:文字颜色|1:文字位置|0xeeeeee:文字背景颜色|00:文字背景透明度|0xffffff:按键文字颜色|0x66cc00:按键默认颜色|0x000033:按键当前颜色|6:自动播放时间(秒)|0:图片过渡效果|1:是否显示按钮|_blank:打开窗口" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>'); 