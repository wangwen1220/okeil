//ȫ��������Ϣ
var ERROR_REG_username  = "�û�������Ϊ��";
var ERROR_REG_password1 = "����������6���ַ�С��20���ַ�";
var ERROR_REG_password2 = "���벻��Ϊ�ջ��������������벻һ��";
var ERROR_REG_passwd 	= "����Կ�׺���Ҫ����������д���μ�";
var ERROR_REG_email 	= "�����ʼ���ַ���벻��ȷ";
var ERROR_REG_nname 	= "��ʵ������������";
var ERROR_REG_mobile 	= "��ϵ�绰��������";
var ERROR_REG_checkbox 	= "��������Ķ���ԱЭ��";

var ERROR_ORDER_bi_fname  	= "��ϵ����������Ϊ��";
var ERROR_ORDER_bi_fmobile 	= "��ϵ�绰����Ϊ��";
var ERROR_ORDER_bi_userid	= "��ס����������Ϊ��";
var ERROR_ORDER_bi_faddress = "�ռ��˵�ַ��������";
// JavaScript Document
String.prototype.trim = function(){	return this.replace(/(^\s*)|(\s*$)/g, "");}  //ȥ���ַ����ռ���÷�ʽ �ַ���.trim()
String.prototype.lengthw = function(){	return this.replace(/[^\x00-\xff]/g,"**").length;}  //���ַ�����ʵ���Ⱥ��֣����ֽڡ��ַ���.lengthw() 
String.prototype.isEmail = function(){	return /^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/.test(this);} //�ж��Ƿ�email  
String.prototype.existChinese = function(){return /^[\x00-\xff]*$/.test(this);}  
String.prototype.isUrl = function(){	return /^http[s]?:\/\/([\w-]+\.)+[\w-]+([\w-./?%&=]*)?$/i.test(this);} //���url
String.prototype.isPhoneCall = function(){	return /(^[0-9]{3,4}\-[0-9]{3,8}$)|(^[0-9]{3,8}$)|(^\([0-9]{3,4}\)[0-9]{3,8}$)|(^0{0,1}13[0-9]{9}$)/.test(this);} //���绰����
String.prototype.isNumber=function(){return /^[0-9]+$/.test(this);} //�������
String.prototype.toNumber = function(def){return isNaN(parseInt(this, 10)) ? def : parseInt(this, 10);} // ����ת��
String.prototype.toMoney = function(def){return isNaN(parseFloat(this)) ? def : parseFloat(this);} // С��ת��

// UTF-8
String.prototype.encodeUTF8 = function ()
{
	var tmp = escape(this);
    var regex = /\+/g;	// s �к���+��
	if(tmp.indexOf("+")>0)
	{
		tmp = tmp.replace(regex,"%2b");
	}
	
	regex = /\=/g			// s �к���=��
	if(tmp.indexOf("=")>0)
	{
		tmp = tmp.replace(regex,"%3d");
	}
	return tmp;
}

//��̬����js��  
function IncludeJS(sSrc){ document.write("<script type='text/javascript' src='" + sSrc + "'><\/script>");} //����js����

//�õ�id,name  ���ж��Ƿ���id,���û�оͶ�name����ֻ����һ������
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

function $$(s){	return document.frames?document.frames[s]:$(s).contentWindow;}  //���iframe����
function $c(s){	return document.createElement(s);}  //�����ڵ㣬����object����
function exist(s){	return $(s)!=null;}  //�ж�ָ��ID�Ƿ����
function dw(s){	document.write(s);}   //���һ������
function hide(s){	$(s).style.display=$(s).style.display=="none"?"":"none";}
function removeNode(s){	if(exist(s))	{		$(s).innerHTML = '';$(s).removeNode?$(s).removeNode():$(s).parentNode.removeChild($(s));	}}   //���߽ڵ�
function getEvent(ee){return ee.srcElement ? ee.srcElement : ee.target;}       //�õ�Event����
function InnerText(obj){if(!obj){return ""; }else {return document.all ? obj.innerText : obj.textContent;}}  //�õ�innerText����

//����¼�  ʹ�÷�ʽ��addEvent(����,"�¼�","����",false)  addEvent(btn,"click","abc()",false);��btn������������¼���������abc();
function addEvent(obj,evType,fn,useCapture ){if (obj.addEventListener){obj.addEventListener( evType, fn, useCapture );return true;}if (obj.attachEvent) return obj.attachEvent( "on" + evType, fn );alert( "Unable to add event listener for " + evType + " to " + obj.tagName );}
function removeEvent( obj, type, fn,useCapture ) { if ( obj.detachEvent ) { obj.detachEvent( 'on'+type, obj[type+fn] );      obj[type+fn] = null;    } else      obj.removeEventListener( type, fn, false ); } 
function HtmlEncode(text){var re = {'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'};for (i in re) text = text.replace(new RegExp(i,'g'), re[i]);return text;}
function HtmlDecode(text){var re = {'&lt;':'<','&gt;':'>','&amp;':'&','&quot;':'"'};for (i in re) text = text.replace(new RegExp(i,'g'), re[i]);return text;}
function setHome(){try{window.external.AddFavorite(window.document.location,window.document.title)}catch(e){};}
function setCopy(_sTxt){try{clipboardData.setData('Text',_sTxt)}catch(e){};alert('��ѡ�����Ѿ����Ƶ��������У�');}
var Browser={};
Browser.isMozilla=(typeof document.implementation!='undefined')&&(typeof document.implementation.createDocument!='undefined')&&(typeof HTMLDocument!='undefined');
Browser.isIE=window.ActiveXObject ? true : false;
Browser.isFirefox=(navigator.userAgent.toLowerCase().indexOf("firefox")!=-1);
Browser.isSafari=(navigator.userAgent.toLowerCase().indexOf("safari")!=-1);
Browser.isOpera=(navigator.userAgent.toLowerCase().indexOf("opera")!=-1);

IncludeJS("/fn_inc/js/object/XmlObject.js");
//gb2312תתutf-8���� ��xmlhttp������������Ŀ����õ�
window.gb2utf8=function(data)
{
��������var glbEncode=[],t,i,j,len
��������gb2utf8_data=data
��������execScript("gb2utf8_data = MidB(gb2utf8_data, 1)+' '", "vbscript")
��������t=escape(gb2utf8_data).replace(/%u/g,"").replace(/(.{2})(.{2})/g,"%$2%$1").replace(/%([A-Z].)%(.{2})/g,"@$1$2")
��������t=t.split("@")
��������i=0
��������len=t.length
��������while(++i<len){
������������j=t[i].substring(0,4)
������������if(!glbEncode[j]) {
����������������gb2utf8_char = eval("0x"+j)
����������������execScript("gb2utf8_char=Chr(gb2utf8_char)","vbscript")
����������������glbEncode[j]=escape(gb2utf8_char).substring(1,6)
������������}
������������t[i]=glbEncode[j]+t[i].substring(4)
��������}
��������gb2utf8_data=gb2utf8_char=null
��������return unescape(t.join("%")).slice(0,-1)
}


///������iframe get������
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

///����ļ�����ʾ��id����  ���ַ�ʽ���ݶ��������
function getfile2id(file,id)  //����file Ϊ�ļ���ַ��id����������ַ�������Ҳ�����Ǻ��������������Խ��������һ����������
{
	   var iframe=new IframePost();
	   iframe.url = file;  //�ļ���ַ
	   iframe.CallBack = function()
	   {
            var tdata=iframe.HTML.document.body.innerHTML;
			if(typeof(id)=="function"){id(HtmlDecode(tdata));return;};
			var obj=typeof(id)=="string"?$(id):id;
			obj.innerHTML=HtmlDecode(tdata);
	   }
	   iframe.create();
}

// AJAX��
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
    	if(xmlObj.overrideMimeType) xmlObj.overrideMimeType("text/html;charset=gb2312");//�趨��gb2312����ʶ������  
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