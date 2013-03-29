//全局配置信息
var ERROR_REG_username  = "用户名不能为空";
var ERROR_REG_password1 = "密码必须大于6个字符小于20个字符";
var ERROR_REG_password2 = "密码不能为空或者两次输入密码不一致";
var ERROR_REG_passwd 	= "密码钥匙很重要，请认真填写并牢记";
var ERROR_REG_email 	= "电子邮件地址输入不正确";
var ERROR_REG_nname 	= "真实姓名不能留空";
var ERROR_REG_mobile 	= "联系电话不能留空";
var ERROR_REG_checkbox 	= "请认真的阅读会员协议";

var ERROR_ORDER_bi_fname  	= "联系人姓名不能为空";
var ERROR_ORDER_bi_fmobile 	= "联系电话不能为空";
var ERROR_ORDER_bi_userid	= "入住人姓名不能为空";
var ERROR_ORDER_bi_faddress = "收件人地址不能留空";
// JavaScript Document
String.prototype.trim = function(){	return this.replace(/(^\s*)|(\s*$)/g, "");}  //去掉字符串空间调用方式 字符串.trim()
String.prototype.lengthw = function(){	return this.replace(/[^\x00-\xff]/g,"**").length;}  //求字符穿真实长度汉字２个字节　字符串.lengthw() 
String.prototype.isEmail = function(){	return /^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/.test(this);} //判断是否email  
String.prototype.existChinese = function(){return /^[\x00-\xff]*$/.test(this);}  
String.prototype.isUrl = function(){	return /^http[s]?:\/\/([\w-]+\.)+[\w-]+([\w-./?%&=]*)?$/i.test(this);} //检查url
String.prototype.isPhoneCall = function(){	return /(^[0-9]{3,4}\-[0-9]{3,8}$)|(^[0-9]{3,8}$)|(^\([0-9]{3,4}\)[0-9]{3,8}$)|(^0{0,1}13[0-9]{9}$)/.test(this);} //检查电话号码
String.prototype.isNumber=function(){return /^[0-9]+$/.test(this);} //检查整数
String.prototype.toNumber = function(def){return isNaN(parseInt(this, 10)) ? def : parseInt(this, 10);} // 整数转换
String.prototype.toMoney = function(def){return isNaN(parseFloat(this)) ? def : parseFloat(this);} // 小数转换

// UTF-8
String.prototype.encodeUTF8 = function ()
{
	var tmp = escape(this);
    var regex = /\+/g;	// s 中含有+号
	if(tmp.indexOf("+")>0)
	{
		tmp = tmp.replace(regex,"%2b");
	}
	
	regex = /\=/g			// s 中含有=号
	if(tmp.indexOf("=")>0)
	{
		tmp = tmp.replace(regex,"%3d");
	}
	return tmp;
}

//动态加载js类  
function IncludeJS(sSrc){ document.write("<script type='text/javascript' src='" + sSrc + "'><\/script>");} //插入js代码

//得到id,name  先判断是否有id,如果没有就读name不过只返回一个对象
function $(id) {
 if (document.getElementById != null)
 {
 if(document.getElementById(id)) return document.getElementById(id);
 else{
	 if(document.getElementsByName(id)) return document.getElementsByName(id)[0];}
 }
 else if (document.all != null) {
 return document.all[id];
 }
 else {
 return null;
 }
}

function $$(s){	return document.frames?document.frames[s]:$(s).contentWindow;}  //获得iframe对象
function $c(s){	return document.createElement(s);}  //开创节点，返回object对象
function exist(s){	return $(s)!=null;}  //判断指定ID是否存在
function dw(s){	document.write(s);}   //输出一行数据
function hide(s){	$(s).style.display=$(s).style.display=="none"?"":"none";}
function removeNode(s){	if(exist(s))	{		$(s).innerHTML = '';$(s).removeNode?$(s).removeNode():$(s).parentNode.removeChild($(s));	}}   //移走节点
function getEvent(ee){return ee.srcElement ? ee.srcElement : ee.target;}       //得到Event属性
function InnerText(obj){if(!obj){return ""; }else {return document.all ? obj.innerText : obj.textContent;}}  //得到innerText属性

//添加事件  使用方式　addEvent(对象,"事件","函数",false)  addEvent(btn,"click","abc()",false);给btn对象加入点击　事件激发函数abc();
function addEvent(obj,evType,fn,useCapture ){if (obj.addEventListener){obj.addEventListener( evType, fn, useCapture );return true;}if (obj.attachEvent) return obj.attachEvent( "on" + evType, fn );alert( "Unable to add event listener for " + evType + " to " + obj.tagName );}
function removeEvent( obj, type, fn,useCapture ) { if ( obj.detachEvent ) { obj.detachEvent( 'on'+type, obj[type+fn] );      obj[type+fn] = null;    } else      obj.removeEventListener( type, fn, false ); } 
function HtmlEncode(text){var re = {'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'};for (i in re) text = text.replace(new RegExp(i,'g'), re[i]);return text;}
function HtmlDecode(text){var re = {'&lt;':'<','&gt;':'>','&amp;':'&','&quot;':'"'};for (i in re) text = text.replace(new RegExp(i,'g'), re[i]);return text;}
function setHome(){try{window.external.AddFavorite(window.document.location,window.document.title)}catch(e){};}
function setCopy(_sTxt){try{clipboardData.setData('Text',_sTxt)}catch(e){};alert('所选数据已经复制到剪贴板中！');}
var Browser={};
Browser.isMozilla=(typeof document.implementation!='undefined')&&(typeof document.implementation.createDocument!='undefined')&&(typeof HTMLDocument!='undefined');
Browser.isIE=window.ActiveXObject ? true : false;
Browser.isFirefox=(navigator.userAgent.toLowerCase().indexOf("firefox")!=-1);
Browser.isSafari=(navigator.userAgent.toLowerCase().indexOf("safari")!=-1);
Browser.isOpera=(navigator.userAgent.toLowerCase().indexOf("opera")!=-1);

IncludeJS("/fn_inc/js/object/XmlObject.js");
//gb2312转转utf-8函数 再xmlhttp获得数据有中文可以用到
window.gb2utf8=function(data)
{
　　　　var glbEncode=[],t,i,j,len
　　　　gb2utf8_data=data
　　　　execScript("gb2utf8_data = MidB(gb2utf8_data, 1)+' '", "vbscript")
　　　　t=escape(gb2utf8_data).replace(/%u/g,"").replace(/(.{2})(.{2})/g,"%$2%$1").replace(/%([A-Z].)%(.{2})/g,"@$1$2")
　　　　t=t.split("@")
　　　　i=0
　　　　len=t.length
　　　　while(++i<len){
　　　　　　j=t[i].substring(0,4)
　　　　　　if(!glbEncode[j]) {
　　　　　　　　gb2utf8_char = eval("0x"+j)
　　　　　　　　execScript("gb2utf8_char=Chr(gb2utf8_char)","vbscript")
　　　　　　　　glbEncode[j]=escape(gb2utf8_char).substring(1,6)
　　　　　　}
　　　　　　t[i]=glbEncode[j]+t[i].substring(4)
　　　　}
　　　　gb2utf8_data=gb2utf8_char=null
　　　　return unescape(t.join("%")).slice(0,-1)
}


///以下是iframe get数据类
function IframePost()
{
	var me = this;
	this.url;
	this.HTML;
	this.CallBack = function() {}
    this.create=function()
    {
    	var obj = document.createElement("iframe");
		addEvent(obj,"load",function(){ me.CallBack(); obj.removeNode(); me.win = null },"");
		obj.src = this.url;
		obj.style.display = 'none';
		document.body.appendChild(obj);
		this.HTML = obj.contentWindow;//.document.body.innerHTML;//.replace(/&lt;/g,"<")..replace(/&gt;/g,">");
	}
}

///获得文件。显示到id里面  这种方式兼容多浏览器　
function getfile2id(file,id)  //这里file 为文件地址，id这里可以是字符串　，也可以是函数名　这样可以将结果返回一个函数处理
{
	   var iframe=new IframePost();
	   iframe.url = file;  //文件地址
	   iframe.CallBack = function()
	   {
            var tdata=iframe.HTML.document.body.innerHTML;
			if(typeof(id)=="function"){id(HtmlDecode(tdata));return;};
			var obj=typeof(id)=="string"?$(id):id;
			obj.innerHTML=HtmlDecode(tdata);
	   }
	   iframe.create();
}

// AJAX类
function AJAXRequest() {
	var xmlObj = false;
	var CBfunc,ObjSelf;
	ObjSelf=this;
	try { xmlObj=new XMLHttpRequest; }
	catch(e) {
		try { xmlObj=new ActiveXObject("MSXML2.XMLHTTP"); }
		catch(e2) {
			try { xmlObj=new ActiveXObject("Microsoft.XMLHTTP"); }
			catch(e3) { xmlObj=false; }
		}
	}
	if (!xmlObj) return false;
	this.method="POST";
	this.url;
	this.async=true;
	this.content="";

	this.callback=function(cbobj) {return;}
	this.send=function() {
		if(!this.method||!this.url) return false;

		xmlObj.open (this.method, this.url, this.async);
		if(this.method=="POST") xmlObj.setRequestHeader("Content-Type","application/x-www-form-urlencoded;charset=UTF-8");
    	if(xmlObj.overrideMimeType) xmlObj.overrideMimeType("text/html;charset=gb2312");//设定以gb2312编码识别数据  
		if(this.async)
		{
			xmlObj.onreadystatechange=function()
			{
				if(xmlObj.readyState==4) {
					if(xmlObj.status==200) {
						ObjSelf.callback(xmlObj);
					}
				}
			}
		}
		if(this.method=="POST") xmlObj.send(this.content);
		else xmlObj.send(null);
		if(!this.async)ObjSelf.callback(xmlObj);
	}
}