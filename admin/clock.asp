<script language="javascript">
<!--
var bsYear; 
var bsDate; 
var bsWeek; 
var arrLen=8; //���鳤�� 
var sValue=0; //��������� 
var dayiy=0; //����ڼ��� 
var miy=0; //�·ݵ��±� 
var iyear=0; //��ݱ�� 
var dayim=0; //���µڼ��� 
var spd=86400; //ÿ������� 

var year1999="30;29;29;30;29;29;30;29;30;30;30;29"; //354 
var year2000="30;30;29;29;30;29;29;30;29;30;30;29"; //354 
var year2001="30;30;29;30;29;30;29;29;30;29;30;29;30"; //384 
var year2002="30;30;29;30;29;30;29;29;30;29;30;29"; //354 
var year2003="30;30;29;30;30;29;30;29;29;30;29;30"; //355 
var year2004="29;30;29;30;30;29;30;29;30;29;30;29;30"; //384 
var year2005="29;30;29;30;29;30;30;29;30;29;30;29"; //354 
var year2006="30;29;30;29;30;30;29;29;30;30;29;29;30"; 

var month1999="����;����;����;����;����;����;����;����;����;ʮ��;ʮһ��;ʮ����" 
var month2001="����;����;����;����;������;����;����;����;����;����;ʮ��;ʮһ��;ʮ����" 
var month2004="����;����;�����;����;����;����;����;����;����;����;ʮ��;ʮһ��;ʮ����" 
var month2006="����;����;����;����;����;����;����;������;����;����;ʮ��;ʮһ��;ʮ����" 
var Dn="��һ;����;����;����;����;����;����;����;����;��ʮ;ʮһ;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;��ʮ;إһ;إ��;إ��;إ��;إ��;إ��;إ��;إ��;إ��;��ʮ"; 

var Ys=new Array(arrLen); 
Ys[0]=919094400;Ys[1]=949680000;Ys[2]=980265600; 
Ys[3]=1013443200;Ys[4]=1044028800;Ys[5]=1074700800; 
Ys[6]=1107878400;Ys[7]=1138464000; 

var Yn=new Array(arrLen); //ũ��������� 
Yn[0]="��î��";Yn[1]="������";Yn[2]="������"; 
Yn[3]="������";Yn[4]="��δ��";Yn[5]="������"; 
Yn[6]="������";Yn[7]="������"; 
var D=new Date(); 
var yy=D.getYear(); 
var mm=D.getMonth()+1; 
var dd=D.getDate(); 
var ww=D.getDay(); 
if (ww==0) ww="<font color=RED>������</font>"; 
if (ww==1) ww="<font color=#008040>����һ</font>"; 
if (ww==2) ww="<font color=#008040>���ڶ�</font>"; 
if (ww==3) ww="<font color=#008040>������</font>"; 
if (ww==4) ww="<font color=#008040>������</font>"; 
if (ww==5) ww="<font color=#008040>������</font>"; 
if (ww==6) ww="<font color=RED>������</font>";
ww=ww; 
var ss=parseInt(D.getTime() / 1000); 
if (yy<100) yy="19"+yy; 

for (i=0;i<arrLen;i++) 
if (ss>=Ys[i]){ 
iyear=i; 
sValue=ss-Ys[i]; //��������� 
} 
dayiy=parseInt(sValue/spd)+1; //��������� 

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

var Dn="��һ;����;����;����;����;����;����;����;����;��ʮ;ʮһ;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;ʮ��;��ʮ;إһ;إ��;إ��;إ��;إ��;إ��;إ��;إ��;إ��;��ʮ"; 
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
bsDate=yy+"��"+mm+"��"+dd+"��"; 
bsYear=Yn[iyear]; 
bsYear2=Mn[miy]+Dn[dayim-1]; 
if (ss>=Ys[7]||ss<Ys[0]) bsYear=Yn[7]; 

function CAL(){ 
	document.write("<font color=navy>"+bsDate+"&nbsp</font>"+bsWeek+"<font color=navy>&nbspũ��:"+bsYear+"&nbsp"+bsYear2+"</font>"); 
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