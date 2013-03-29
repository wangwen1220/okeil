String.prototype.trim = function(){	return this.replace(/(^\s*)|(\s*$)/g, "");}  //去掉字符串空间调用方式 字符串.trim()
var SelectDateObj;
function getobjectx(e){
var l=e.offsetLeft;
while(e=e.offsetParent){
l+=e.offsetLeft;
	}
return l;

}

function getobjecty(e){
var t=e.offsetTop;
while(e=e.offsetParent){
t+=e.offsetTop;
	}
return t;
}

function hidedate()
{
 
	if(document.all.SelectDateList.style.display=='block')
	document.all.SelectDateList.style.display='none';

}
function showdate(txtid){
SelectDateObj=eval("document.all."+txtid);

if(document.all.SelectDateList.style.display=='block')
document.all.SelectDateList.style.display='none';
else{
TableFunction().JumpToRun("start");
//posX = event.clientX ; 
//posY = event.clientY ;
var x,y;
x=getobjectx(SelectDateObj)	;
y=getobjecty(SelectDateObj)	;

document.all.SelectDateList.style.left = x+SelectDateObj.offsetWidth ;
document.all.SelectDateList.style.top = y+SelectDateObj.offsetHeight;
document.all.SelectDateList.style.display='block';
}
}
function TableFunction(){
    this.GetDateStr=function(y,m){
        this.DayArray=[];
        for(var i=0;i<42;i++)this.DayArray[i]=" ";
        for(var i=0;i<new Date(y,m,0).getDate();i++)this.DayArray[i+new Date(y,m-1,1).getDay()]=i+1;
        return this.DayArray;
        }
    this.GetTableStr=function(y,m){
        this.DateArray=["日","一","二","三","四","五","六"];
    this.DStr="<div id=SelectDateList name=SelectDateList style='display:none;position:absolute;'>"
        this.DStr=this.DStr+"<table oncontextmenu='return false' onselectstart='return false' style='width:160;cursor:default;border:1 solid #9c9c9c;border-right:0;border-bottom:0;filter:progid:dximagetransform.microsoft.dropshadow(color=#e3e3e3,offx=3,offy=3,positive=true)' border='0' cellpadding='0' cellspacing='0'>\n"+
        "<tr><td colspan='7' class='TrOut'>"+
        "<table width='100%' height='100%'border='0' cellpadding='0' cellspacing='0'><tr align='center'>\n"+
        "<td width='20' style='font-family:\"webdings\";font-size:9pt' onclick='TableFunction().JumpToRun(\"b\")' onmouseover='this.style.color=\"#ff9900\"' onmouseout='this.style.color=\"\"'>3</td>\n"+
        "<td id='YearTD' width='70' onmouseover='this.style.background=\"#cccccc\"' onmouseout='this.style.background=\"\"' onclick='TableFunction().WriteSelect(this,this.innerText.split(\" \")[0],\"y\",false)'>"+y+" 年</td>\n"+
        "<td id='MonthTD' width='47' onmouseover='this.style.background=\"#cccccc\"' onmouseout='this.style.background=\"\"' onclick='TableFunction().WriteSelect(this,this.innerText.split(\" \")[0],\"m\",false)'>"+m+" 月</td>\n"+
        "<td width='20' style='font-family:\"webdings\";font-size:9pt' onclick='TableFunction().JumpToRun(\"n\")' onmouseover='this.style.color=\"#ff9900\"' onmouseout='this.style.color=\"\"'>4</td></tr></table>\n"+
        "</td></tr>\n"+
        "<tr align='center'>\n";
        for(var i=0;i<7;i++)
        this.DStr+="<td class='TrOut'>"+DateArray[i]+"</td>\n";
        this.DStr+="</tr>\n";
        for(var i=0;i<6;i++){
        this.DStr+="<tr align='center'>\n";
        for(var j=0;j<7;j++){
            var CS=new Date().getDate()==this.GetDateStr(y,m)[i*7+j]?"TdOver":"TdOut";
            this.DStr+="<td id='TD' class='"+CS+"' cs='"+CS+"' onmouseover='this.className=\"TdOver\"' onmouseout='if(this.cs!=\"TdOver\")this.className=\"TdOut\"' onclick='TableFunction().AlertDay()'>"+this.GetDateStr(y,m)[i*7+j]+" </td>\n";
            }
        this.DStr+="</tr>\n";
        }
        this.DStr+="</table></div>";
        return this.DStr;
        }
    this.WriteSelect=function(obj,values,action,getobj){
        if(values=="")return;
        if(getobj){
            obj.innerHTML=values+(action=="y"?" 年":" 月");
            this.RewriteTableStr(YearTD.innerText.split(" ")[0],MonthTD.innerText.split(" ")[0]);
            return false;
            }
        var StrArray=[];
        if(action=="y"){
            for(var i=0;i<15;i++){
                var year=values-7+i;
                StrArray[i]="<option value='"+year+"' "+(values==year?"selected":"")+"> "+year+"年</option>\n";
                }
            obj.innerHTML="<select id='select1' style='width:67' onchange='TableFunction().WriteSelect(parentElement,this.value,\"y\",true)' onblur='YearTD.innerText=this.value+\" 年\"'>\n"+StrArray.join("")+"</select>";
            select1.focus();
            }
        if(action=="m"){
            for(var i=1;i<13;i++)
                StrArray[i]="<option value='"+i+"' "+(i==values?"selected":"")+"> "+i+"月</option>\n";
            obj.innerHTML="<select id='select2' style='width:47' onchange='TableFunction().WriteSelect(parentElement,this.value,\"m\",true)' onblur='MonthTD.innerText=this.value+\" 月\"'>\n"+StrArray.join("")+"</select>";
            select2.focus();
            }
        }
    this.RewriteTableStr=function(y,m){
        var TArray=this.GetDateStr(y,m);
        var len=TArray.length;
        for(var i=0;i<len;i++){
            TD[i].innerHTML=TArray[i]+" ";
            TD[i].className="TdOut";
            TD[i].cs="TdOut";
            if(new Date().getYear()==y&&new Date().getMonth()+1==m&&new Date().getDate()==TArray[i]){
                TD[i].className="TdOver";
                TD[i].cs="TdOver";
                }
            }
        }
    this.JumpToRun=function(action){
        var YearNO=YearTD.innerText.split(' ')[0];
        var MonthNO=MonthTD.innerText.split(' ')[0];
        if(action=="b"){
            if(MonthNO=="1"){
                MonthNO=13;
                YearNO=YearNO-1;
                }
            MonthTD.innerText=MonthNO-1+" 月";
            YearTD.innerText=YearNO+" 年";
            this.RewriteTableStr(YearNO,MonthNO-1);
            }
        if(action=="n"){
            if(MonthNO=="12"){
                MonthNO=0;
                YearNO=YearNO-(-1);
                }
            YearTD.innerText=YearNO+" 年";
            MonthTD.innerText=MonthNO-(-1)+" 月";
            this.RewriteTableStr(YearNO,MonthNO-(-1));
            }
    if(action=="start"){
                MonthNO=new Date().getMonth();
                YearNO=new Date().getYear();
            YearTD.innerText=YearNO+" 年";
            MonthTD.innerText=MonthNO-(-1)+" 月";
            this.RewriteTableStr(YearNO,MonthNO-(-1));
            }
        }
    this.AlertDay=function(){
        //if(event.srcElement.innerText!=" ")

//        alert(YearTD.innerText.split(' ')[0]+"年"+MonthTD.innerText.split(' ')[0]+"月"+event.srcElement.innerText+"日");
		var oldvalue=  SelectDateObj.value;
		if (oldvalue.length==0 || oldvalue ==null)
		{
			//SelectDateObj.value=YearTD.innerText.split(' ')[0]+"-"+MonthTD.innerText.split(' ')[0]+"-"+event.srcElement.innerText + new Date().toTimeString().substring(0,5);
			SelectDateObj.value=YearTD.innerText.split(' ')[0]+"-"+MonthTD.innerText.split(' ')[0]+"-"+event.srcElement.innerText.trim();
		}
		else
		{
		  if (oldvalue.lastIndexOf(":")>=0)
		  {
			SelectDateObj.value =YearTD.innerText.split(' ')[0]+"-"+MonthTD.innerText.split(' ')[0]+"-"+event.srcElement.innerText+oldvalue.split(" ")[1];	
		  }
		  else
		  {
			//SelectDateObj.value=YearTD.innerText.split(' ')[0]+"-"+MonthTD.innerText.split(' ')[0]+"-"+event.srcElement.innerText + new Date().toTimeString().substring(0,5);
			SelectDateObj.value=YearTD.innerText.split(' ')[0]+"-"+MonthTD.innerText.split(' ')[0]+"-"+event.srcElement.innerText.trim();
		  }
		}
	
 //   SelectDateObj.disabled=true;
		document.all.SelectDateList.style.display="none";
        }
    return this;
    }
document.write(TableFunction().GetTableStr(new Date().getYear(),new Date().getMonth()+1));
