<script language="javascript">
<!--
var bsYear; 
var bsDate; 
var bsWeek; 
var arrLen=8; //数组长度 
var sValue=0; //当年的秒数 
var dayiy=0; //当年第几天 
var miy=0; //月份的下标 
var iyear=0; //年份标记 
var dayim=0; //当月第几天 
var spd=86400; //每天的秒数 

var year1999="30;29;29;30;29;29;30;29;30;30;30;29"; //354 
var year2000="30;30;29;29;30;29;29;30;29;30;30;29"; //354 
var year2001="30;30;29;30;29;30;29;29;30;29;30;29;30"; //384 
var year2002="30;30;29;30;29;30;29;29;30;29;30;29"; //354 
var year2003="30;30;29;30;30;29;30;29;29;30;29;30"; //355 
var year2004="29;30;29;30;30;29;30;29;30;29;30;29;30"; //384 
var year2005="29;30;29;30;29;30;30;29;30;29;30;29"; //354 
var year2006="30;29;30;29;30;30;29;29;30;30;29;29;30"; 

var month1999="正月;二月;三月;四月;五月;六月;七月;八月;九月;十月;十一月;十二月" 
var month2001="正月;二月;三月;四月;闰四月;五月;六月;七月;八月;九月;十月;十一月;十二月" 
var month2004="正月;二月;闰二月;三月;四月;五月;六月;七月;八月;九月;十月;十一月;十二月" 
var month2006="正月;二月;三月;四月;五月;六月;七月;闰七月;八月;九月;十月;十一月;十二月" 
var Dn="初一;初二;初三;初四;初五;初六;初七;初八;初九;初十;十一;十二;十三;十四;十五;十六;十七;十八;十九;二十;廿一;廿二;廿三;廿四;廿五;廿六;廿七;廿八;廿九;三十"; 

var Ys=new Array(arrLen); 
Ys[0]=919094400;Ys[1]=949680000;Ys[2]=980265600; 
Ys[3]=1013443200;Ys[4]=1044028800;Ys[5]=1074700800; 
Ys[6]=1107878400;Ys[7]=1138464000; 

var Yn=new Array(arrLen); //农历年的名称 
Yn[0]="己卯年";Yn[1]="庚辰年";Yn[2]="辛巳年"; 
Yn[3]="壬午年";Yn[4]="癸未年";Yn[5]="甲申年"; 
Yn[6]="乙酉年";Yn[7]="丙戌年"; 
var D=new Date(); 
var yy=D.getYear(); 
var mm=D.getMonth()+1; 
var dd=D.getDate(); 
var ww=D.getDay(); 
if (ww==0) ww="<font color=RED>星期日</font>"; 
if (ww==1) ww="<font color=#008040>星期一</font>"; 
if (ww==2) ww="<font color=#008040>星期二</font>"; 
if (ww==3) ww="<font color=#008040>星期三</font>"; 
if (ww==4) ww="<font color=#008040>星期四</font>"; 
if (ww==5) ww="<font color=#008040>星期五</font>"; 
if (ww==6) ww="<font color=RED>星期六</font>";
ww=ww; 
var ss=parseInt(D.getTime() / 1000); 
if (yy<100) yy="19"+yy; 

for (i=0;i<arrLen;i++) 
if (ss>=Ys[i]){ 
iyear=i; 
sValue=ss-Ys[i]; //当年的秒数 
} 
dayiy=parseInt(sValue/spd)+1; //当年的天数 

var dpm=year1999; 
if (iyear==1) dpm=year2000; 
if (iyear==2) dpm=year2001; 
if (iyear==3) dpm=year2002; 
if (iyear==4) dpm=year2003; 
if (iyear==5) dpm=year2004; 
if (iyear==6) dpm=year2005; 
if (iyear==7) dpm=year2006; 
dpm=dpm.split(";"); 

var Mn=month1999; 
if (iyear==2) Mn=month2001; 
if (iyear==5) Mn=month2004; 
if (iyear==7) Mn=month2006; 
Mn=Mn.split(";"); 

var Dn="初一;初二;初三;初四;初五;初六;初七;初八;初九;初十;十一;十二;十三;十四;十五;十六;十七;十八;十九;二十;廿一;廿二;廿三;廿四;廿五;廿六;廿七;廿八;廿九;三十"; 
Dn=Dn.split(";"); 

dayim=dayiy; 

var total=new Array(13); 
total[0]=parseInt(dpm[0]); 
for (i=1;i<dpm.length-1;i++) total[i]=parseInt(dpm[i])+total[i-1]; 
for (i=dpm.length-1;i>0;i--) 
if (dayim>total[i-1]){ 
dayim=dayim-total[i-1]; 
miy=i; 
} 
bsWeek=ww; 
bsDate=yy+"年"+mm+"月"+dd+"日"; 
bsYear=Yn[iyear]; 
bsYear2=Mn[miy]+Dn[dayim-1]; 
if (ss>=Ys[7]||ss<Ys[0]) bsYear=Yn[7]; 

function CAL(){ 
	document.write("<font color=navy>"+bsDate+"&nbsp</font>"+bsWeek+"<font color=navy>&nbsp农历:"+bsYear+"&nbsp"+bsYear2+"</font>"); 
} 

function Chen_CAL(){ 
	
	c1=new Image(); c1.src="images/clock/c1.gif"
	c2=new Image(); c2.src="images/clock/c2.gif"
	c3=new Image(); c3.src="images/clock/c3.gif"
	c4=new Image(); c4.src="images/clock/c4.gif"
	c5=new Image(); c5.src="images/clock/c5.gif"
	c6=new Image(); c6.src="images/clock/c6.gif"
	c7=new Image(); c7.src="images/clock/c7.gif"
	c8=new Image(); c8.src="images/clock/c8.gif"
	c9=new Image(); c9.src="images/clock/c9.gif"
	c0=new Image(); c0.src="images/clock/c0.gif"
	cb=new Image(); cb.src="images/clock/cb.gif"
	
	showtime()
	tdyear.innerHTML="<font color=navy>"+bsDate+"</font>"
	tdweek.innerHTML=bsWeek
	tdcyear.innerHTML="<font color=navy>"+bsYear+"&nbsp"+bsYear2+"</font>"
	//document.write("<font color=navy>"+bsDate+"</font><br>"+bsWeek+"<br><font color=navy>"+bsYear+"&nbsp"+bsYear2+"</font>"); 
}

function showtime(){
	if (!document.images)
	return
	var Digital=new Date()
	var h=Digital.getHours()
	var m=Digital.getMinutes()
	var s=Digital.getSeconds()

	if (h<=9){
		document.images.a.src=cb.src
		document.images.b.src=eval("c"+h+".src")
	}
	else {
	document.images.a.src=eval("c"+Math.floor(h/10)+".src")
	document.images.b.src=eval("c"+(h%10)+".src")
	}
	if (m<=9){
		document.images.d.src=c0.src
		document.images.e.src=eval("c"+m+".src")
	}
	else {
		document.images.d.src=eval("c"+Math.floor(m/10)+".src")
		document.images.e.src=eval("c"+(m%10)+".src")
	}
	if (s<=9){
		document.g.src=c0.src
		document.images.h.src=eval("c"+s+".src")
	}
	else {
	document.images.g.src=eval("c"+Math.floor(s/10)+".src")
	document.images.h.src=eval("c"+(s%10)+".src")
	}	
	setTimeout("showtime()",1000)
}
//-->
</script>