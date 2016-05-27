<%
Dim UserPointInfo(4)
'UBB代码勘套循环的最多次数，避免死循环加入此变量
Const MaxLoopcount=100
%>
<script language=vbscript runat=server>
Dim Ubblists
'[/img]编号:1.[/upload]编号:2.[/dir]编号:3.[/qt]编号:4.[/mp]编号:5.
'[/rm]编号:6.[/sound]编号:7.[/flash]编号:8.[/money]编号:9.[/point]编号:10.
'[/usercp]编号:11.[/power]编号:12.[/post]编号:13.[/replyview]编号:14.[/usemoney]编号:15.
'[/url]编号:16.[/email]编号:17.http编号:18.https编号:19.ftp编号:20.rtsp编号:21.
'mms编号:22.[/html]编号:23.[/code]编号:24.[/color]编号:25.[/face]编号:26.[/align]编号:27.
'[/quote]编号:28.[/fly]编号:29.[/move]编号:30.[/shadow]编号:31.[/glow]编号:32.[/size]编号:33.
'[/i]编号:34.[/b]编号:35.[/u]编号:36.[em编号:37.www.编号:38. 
Class Dvbbs_UbbCode
	Public Re,reed,isgetreed
	'论坛内容部分UBBCODE，入口：内容、用户组ID、模式(1=帖子/2=公告、短信等)、模式2(0=新版/1=老版)
	Public Function Dv_UbbCode(s,PostUserGroup,PostType,sType)
		Dim ii,ranNum
		If PostType=2 Then
			Dvbbs.Board_Setting=Split("1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1",",")
			Dvbbs.Board_Setting(6)=1
			Dvbbs.Board_Setting(5)=0:Dvbbs.Board_Setting(7)=1
			Dvbbs.Board_Setting(8)=1:Dvbbs.Board_Setting(9)=1
			Dvbbs.Board_Setting(10)=0:Dvbbs.Board_Setting(11)=0
			Dvbbs.Board_Setting(12)=0:Dvbbs.Board_Setting(13)=0
			Dvbbs.Board_Setting(14)=0:Dvbbs.Board_Setting(15)=0
			Dvbbs.Board_Setting(23)=0:Dvbbs.Board_Setting(44)=0
		End If
		If Dvbbs.UserID=0 Then
			UserPointInfo(0)=0:UserPointInfo(1)=0:UserPointInfo(2)=0:UserPointInfo(3)=0:UserPointInfo(4)=0
		Else
			UserPointInfo(0)=Dvbbs.MyUserInfo(21):UserPointInfo(1)=Dvbbs.MyUserInfo(22):UserPointInfo(2)=Dvbbs.MyUserInfo(23):UserPointInfo(3)=Dvbbs.MyUserInfo(24):UserPointInfo(4)=Dvbbs.MyUserInfo(8)
		End If
		Dim po
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		If (Not InStr(Ubblists,",39,")>0) Or Ubblists="" Or IsNull(Ubblists) Then'老贴子
			s=server.htmlencode(s)
		End If
		re.Pattern="(<br>)"
		s=re.Replace(s,"[br]")
		re.Pattern="(<s+cript(.[^>]*)>)"
		s=re.Replace(s,"&lt;&#83cript$2&gt;")
		re.Pattern="(<\/s+cript>)"
		s=re.Replace(s,"&lt;/&#83cript&gt;")
		'如果论坛没开放HTML脚本,拦截所有标记和脚本
		If Dvbbs.Board_Setting(5)="0" Then
			re.Pattern="(<(i|b|p)>)"
			s=re.Replace(s,"[$2]")
			re.Pattern="(<(\/i|\/b|\/p)>)"
			s=re.Replace(s,"[$2]")
			re.Pattern="(<DIV class=quote>)((.|\n)*)(<\/div>)"
			s=re.Replace(s,"[quote]$2[/quote]")
			'先去掉标记中的换行
			re.Pattern="(>)("&vbNewLine&")(<)"
			s=re.Replace(s,"$1$3") 
			re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
			s=re.Replace(s,"$1$3")
			re.Pattern="<(.[^>]*)>"
			s=re.Replace(s,"")
			re.Pattern="(\[(i|b|p)\])"
			S=re.Replace(S,"<$2>")
			re.Pattern="(\[(\/i|\/b|\/p)\])"
			S=re.Replace(s,"<$2>")
			If Dv_FilterJS2(s)Then
				re.Pattern="\[(br)\]"
				s=re.Replace(s,"<$1>")
				re.Pattern = "(&nbsp;)"
				s = re.Replace(s,Chr(9))
				re.Pattern = "(<br>)"
				s = re.Replace(s,vbNewLine)
				re.Pattern = "(<p>)"
				s = re.Replace(s,"")
				re.Pattern = "(<\/p>)"
				s = re.Replace(s,vbNewLine)
				s=server.htmlencode(s)
				s="<form name=""scode"&replyid&""" method=""post"" action=""""><TABLE class=tableborder2 cellSpacing=1 cellPadding=3 width=""100%"" align=center border=0><TR><TH height=22>以下内容含脚本,或可能导致页面不正常的代码</TH></TR><TR><TD class=tablebody1 align=middle width=""98%""><TEXTAREA id=CodeText style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDTH: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=20 cols=120>"&s&"</TEXTAREA></TD></TR><TR><TD class=tablebody2 align=middle width=""98%""><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行.</td></tr><TR><TD class=tablebody1 align=middle width=""98%""><input type=""button"" name=""run"" value=""运行代码"" onclick=""Dvbbs_ViewCode("&replyid&");""></TD></TR></TABLE></form>"
					s = Replace(s, vbNewLine, "")
					s = Replace(s, CHR(10), "")
					s = Replace(s, CHR(13), "")
				Dv_UbbCode=s
				Exit Function
			End If
		Else
			If Dv_FilterJS(s)Then
				re.Pattern="\[(br)\]"
				s=re.Replace(s,"<$1>")
				re.Pattern = "(&nbsp;)"
				s = re.Replace(s,Chr(9))
				re.Pattern = "(<br>)"
				s = re.Replace(s,vbNewLine)
				re.Pattern = "(<p>)"
				s = re.Replace(s,"")
				re.Pattern = "(<\/p>)"
				s = re.Replace(s,vbNewLine)
				s=server.htmlencode(s)
				s="<form name=""scode"&replyid&""" method=""post"" action=""""><TABLE class=tableborder2 cellSpacing=1 cellPadding=3 width=""100%"" align=center border=0><TR><TH height=22>以下内容含脚本,或可能导致页面不正常的代码</TH></TR><TR><TD class=tablebody1 align=middle width=""98%""><TEXTAREA id=CodeText style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDTH: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=20 cols=120>"&s&"</TEXTAREA></TD></TR><TR><TD class=tablebody2 align=middle width=""98%""><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行.</td></tr><TR><TD class=tablebody1 align=middle width=""98%""><input type=""button"" name=""run"" value=""运行代码"" onclick=""Dvbbs_ViewCode("&replyid&");""></TD></TR></TABLE></form>"
					s = Replace(s, vbNewLine, "")
					s = Replace(s, CHR(10), "")
					s = Replace(s, CHR(13), "")
				Dv_UbbCode=s
				Exit Function
			End If
			re.Pattern="<((asp|\!|%))"
			s=re.Replace(s,"&lt;$1")
			re.Pattern="(>)("&vbNewLine&")(<)"
			s=re.Replace(s,"$1$3") 
			re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
			s=re.Replace(s,"$1$3")
		End If
		s = Replace(s, "	", "&nbsp;")
		s = Replace(s, "  ", "&nbsp;&nbsp;")
		re.Pattern="<(\w+)(&nbsp;)+([^>]*)>"
		s = re.Replace(s,"<$1 $3>")
		s = Replace(s, vbNewLine, "<br>")
		s = Replace(s, CHR(10), "")
		s = Replace(s, CHR(13), "")
		s = Replace(s, "[br]", "<br>")
		s=dv_fixHTML(s)
		'去掉图片中的脚本代码
		re.Pattern="<IMG.[^>]*SRC(=| )(.[^>]*)>"
		s=re.replace(s,"<IMG SRC=$2 onclick=""javascript:window.open(this.src);"" style=""CURSOR: pointer"">")
		If (Trim(UbbLists)=",39," Or Trim(UbbLists)=",39,40,") And Not InStr(Lcase(s),"[username")>0 Then 
			Dv_UbbCode=s
			Exit Function
		End If
		'img code
		If InStr(Ubblists,",1,")>0 Or sType=1 Then s=Dv_UbbCode_S2(s,"\[IMG\]","\[\/IMG\]","IMG","<a onfocus=this.blur() href=""$1"" target=_blank title=新窗口打开><IMG SRC=""$1"" border=0 ></a>","<IMG SRC=""skins/default/filetype/gif.gif"" border=0><a onfocus=this.blur() href=""$1"" target=_blank>$1</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(7)))
		'upload code
		If InStr(Ubblists,",2,")>0 Or sType=1 Then s=Dv_UbbCode_U(s,PostUserGroup,Cint(Dvbbs.Board_Setting(7)))

		'media code
		If InStr(Ubblists,",3,")>0 Or sType=1 Then s=Dv_UbbCode_iS2(s,"\[DIR=(.[^\[]*)\]","\[\/DIR\]","DIR","<object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=$1 height=$2><param name=src value=$3><embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed></object>","<a href=$3 target=_blank>$3</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(9)),"=*([0-9]*),*([0-9]*)")

		If InStr(Ubblists,",4,")>0 Or sType=1 Then s=Dv_UbbCode_iS2(s,"\[QT=(.[^\[]*)\]","\[\/QT\]","QT","<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>","<a href=$3 target=_blank>$3</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(9)),"=*([0-9]*),*([0-9]*)")

		If InStr(Ubblists,",5,")>0 Or sType=1 Then
		s=Dv_UbbCode_iS2(s,"\[MP=(.[^\[]*)\]","\[\/MP\]","MP","<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=""$3"" width=$1 height=$2></embed></object>","<a href=""$3"" target=_blank>$3</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(9)),"=*([0-9]*),*([0-9]*)")
		'Dv7 MediaPlayer自定义播放模式；
		s=Dv_UbbCode_iS2(s,"\[MP=(.[^\[]*)\]","\[\/MP\]","MP","<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><PARAM NAME=AUTOSTART VALUE=$3 ><param name=ShowStatusBar value=-1><param name=Filename value=$4><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=""$4"" width=$1 height=$2></embed></object>","<a href=""$4"" target=_blank>$4</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(9)),"=*([0-9]*),*([0-9]*),*([0|1|true|false]*)")
		End If

		If InStr(Ubblists,",6,")>0 Or sType=1 Then
		s=Dv_UbbCode_iS2(s,"\[RM=(.[^\[]*)\]","\[\/RM\]","RM","<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=""$3""><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=""$3""><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>","<a href=$3 target=_blank>$3</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(9)),"=*([0-9]*),*([0-9]*)")
		'Dv7 RealPlayer自定义播放模式；
		s=Dv_UbbCode_iS2(s,"\[RM=(.[^\[]*)\]","\[\/RM\]","RM","<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=""$4""><PARAM NAME=CONSOLE VALUE=""$4""><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=$3 ></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=""video"" width=$1><PARAM NAME=SRC VALUE=""$4""><PARAM NAME=AUTOSTART VALUE=$3><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=""$4""></OBJECT>","<a href=$4 target=_blank>$4</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(9)),"=*([0-9]*),*([0-9]*),*([0|1|true|false]*)")
		End If

		If InStr(Ubblists,",7,")>0 Or sType=1 Then s=Dv_UbbCode_S2(s,"\[sound\]","\[\/sound\]","sound","<a href=""$1"" target=_blank><IMG SRC=skins/default/filetype/mid.gif border=0 alt=""背景音乐""></a><bgsound src=""$1"" loop=""-1"">","<a href=$1 target=_blank>$1</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(9)))

		'flash code
		If InStr(Ubblists,",8,")>0 Or sType=1 Then
			s=Dv_UbbCode_S2(s,"\[FLASH\]","\[\/FLASH\]","FLASH","<a href=""$1"" TARGET=_blank><IMG SRC=skins/default/filetype/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$1""><PARAM NAME=quality VALUE=high><embed src=""$1"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$1</embed></OBJECT>","<IMG SRC="&Dvbbs.Forum_info(7)&"swf.gif border=0><a href=$1 target=_blank>$1</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(44)))
			s=Dv_UbbCode_iS2(s,"\[FLASH=(.[^\[]*)\]","\[\/FLASH\]","FLASH","<a href=""$3"" TARGET=_blank><IMG SRC=skins/default/filetype/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$1 height=$2><PARAM NAME=movie VALUE=""$3""><PARAM NAME=quality VALUE=high><embed src=""$3"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$1 height=$2>$3</embed></OBJECT>","<a href=$3 target=_blank>$3</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(44)),"=*([0-9]*),*([0-9]*)")
		End If
		'point view
		If InStr(Ubblists,",9,")>0 Or sType=1 Then  s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"\[money=*([0-9]*)\]","\[\/money\]","money","$1<hr noshade size=1><font color=gray>以下内容需要金钱数达到<B>$3</B>才可以浏览</font><BR>$4<hr noshade size=1>$6","$1<hr noshade size=1><font color="&Dvbbs.Mainsetting(1)&">以下内容需要金钱数达到<B>$3</B>才可以浏览</font><hr noshade size=1>$6",UserPointInfo(0),Cint(Dvbbs.Board_Setting(10)))
		If InStr(Ubblists,",10,")>0 Or sType=1 Then s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"\[point=*([0-9]*)\]","\[\/point\]","point","$1<hr noshade size=1><font color=gray>以下内容需要积分达到<B>$3</B>才可以浏览</font><BR>$4<hr noshade size=1>$6","$1<hr noshade size=1><font color="&Dvbbs.Mainsetting(1)&">以下内容需要积分达到<B>$3</B>才可以浏览</font><hr noshade size=1>$6",UserPointInfo(1),Cint(Dvbbs.Board_Setting(11)))
		If InStr(Ubblists,",11,")>0 Or sType=1 Then s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"\[UserCP=*([0-9]*)\]","\[\/UserCP\]","UserCP","$1<hr noshade size=1><font color=gray>以下内容需要魅力达到<B>$3</B>才可以浏览</font><BR>$4<hr noshade size=1>$6","$1<hr noshade size=1><font color="&Dvbbs.Mainsetting(1)&">以下内容需要魅力达到<B>$3</B>才可以浏览</font><hr noshade size=1>$6",UserPointInfo(2),Cint(Dvbbs.Board_Setting(12)))
		If InStr(Ubblists,",12,")>0 Or sType=1 Then s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"\[Power=*([0-9]*)\]","\[\/Power\]","Power","$1<hr noshade size=1><font color=gray>以下内容需要威望达到<B>$3</B>才可以浏览</font><BR>$4<hr noshade size=1>$6","$1<hr noshade size=1><font color="&Dvbbs.Mainsetting(1)&">以下内容需要威望达到<B>$3</B>才可以浏览</font><hr noshade size=1>$6",UserPointInfo(3),Cint(Dvbbs.Board_Setting(13)))
		If InStr(Ubblists,",13,")>0 Or sType=1 Then s=Dv_UbbCode_Get(s,PostUserGroup,PostType,"\[Post=*([0-9]*)\]","\[\/Post\]","Post","$1<hr noshade size=1><font color=gray>以下内容需要帖子数达到<B>$3</B>才可以浏览</font><BR>$4<hr noshade size=1>$6","$1<hr noshade size=1><font color="&Dvbbs.Mainsetting(1)&">以下内容需要帖子数达到<B>$3</B>才可以浏览</font><hr noshade size=1>$6",UserPointInfo(4),Cint(Dvbbs.Board_Setting(14)))
		If InStr(Ubblists,",14,")>0 Or sType=1 Then s=UBB_REPLYVIEW(s,PostUserGroup,PostType)
		If InStr(Ubblists,",15,")>0 Or sType=1 Then s=UBB_USEMONEY(s,PostUserGroup,PostType)
		'url code
		If InStr(Ubblists,",16,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"\[URL\]","\[\/URL\]","URL","<A HREF=""$1"" TARGET=_blank>$1</A>")
			're.Pattern="(\[URL=(.[^:\/\/|\[]*)\])(.[^\[]*)(\[\/URL\])"
			's= re.Replace(s,"<A HREF=""http://$2"" TARGET=_blank>$3</A>")
			re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
			s= re.Replace(s,"<A HREF=""$2"" TARGET=_blank>$3</A>")
		End If
		'email code
		If InStr(Ubblists,",17,")>0 Or sType=1 Then
			s=Dv_UbbCode_S1(s,"\[EMAIL\]","\[\/EMAIL\]","EMAIL","<img align=absmiddle src=skins/default/email1.gif ><A HREF=""mailto:$1"">$1</A>")
			re.Pattern="(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
			s= re.Replace(s,"<img align=absmiddle src=skins/default/email1.gif ><A HREF=""mailto:$2"" TARGET=_blank>$3</A>")
		End If
		If InStr(Ubblists,",37,")>0 Or sType=1 Then
			If (Cint(Dvbbs.Board_Setting(8)) = 1 Or PostUserGroup<4) And InStr(Lcase(s),"[em")>0 Then
				re.Pattern="\[em(.[^\[]*)\]"
				s=re.Replace(s,"<img src="&EmotPath&"em$1.gif border=0 align=middle>")
			Else
				re.Pattern="\[em(.[^\[]*)\]"
				s=re.Replace(s,"")
			End If
		End If
		If InStr(Ubblists,",23,")>0 Or sType=1 Then s=Dv_UbbCode_S1(s,"\[HTML\]","\[\/HTML\]","HTML","<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""6"" class="""&abgcolor&"""><td><b>以下内容为程序代码:</b><br>$1</td></table>")
		If InStr(Ubblists,",24,")>0 Or sType=1 Then s=Dv_UbbCode_S1(s,"\[code\]","\[\/code\]","code","<div class=htmlcode><b>以下内容为程序代码:</b><br>$1</div>")
		If InStr(Ubblists,",25,")>0 Or sType=1 Then s=Dv_UbbCode_C(s)
		If InStr(Ubblists,",26,")>0 Or sType=1 Then s=Dv_UbbCode_F(s)
		If InStr(Ubblists,",27,")>0 Or sType=1 Then s=Dv_UbbCode_Align(s)
		If InStr(Lcase(s),"center]")>0 Or sType=1 Then s=Dv_UbbCode_S1(s,"\[center\]","\[\/center\]","center","<div align=center>$1</div>")
		If InStr(Ubblists,",28,")>0 Or sType=1 Then s=Dv_UbbCode_Q(s)
		If InStr(Ubblists,",29,")>0 Or sType=1 Then s=Dv_UbbCode_S1(s,"\[fly\]","\[\/fly\]","fly","<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
		If InStr(Ubblists,",30,")>0 Or sType=1 Then s=Dv_UbbCode_S1(s,"\[move\]","\[\/move\]","move","<MARQUEE scrollamount=3>$1</marquee>")
		If InStr(Ubblists,",31,")>0 Or sType=1 Then s=Dv_UbbCode_iS1(s,"\[SHADOW=(.[^\[]*)\]","\[\/SHADOW\]","SHADOW","<div style=""width:$1px;filter:shadow(color=$2, strength=$3)"">$4</div>","=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)")
		If InStr(Ubblists,",32,")>0 Or sType=1 Then s=Dv_UbbCode_iS1(s,"\[GLOW=(.[^\[]*)\]","\[\/GLOW\]","GLOW","<div style=""width:$1px;filter:glow(color=$2, strength=$3)"">$4</div>","=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)")
		If InStr(Ubblists,",33,")>0 Or sType=1 Then s=Dv_UbbCode_S(s)
		If InStr(Ubblists,",34,")>0 Or sType=1 Then s=Dv_UbbCode_S1(s,"\[i\]","\[\/i\]","i","<i>$1</i>")
		If InStr(Ubblists,",35,")>0 Or sType=1 Then s=Dv_UbbCode_S1(s,"\[b\]","\[\/b\]","b","<b>$1</b>")
		If InStr(Ubblists,",36,")>0 Or sType=1 Then s=Dv_UbbCode_S1(s,"\[u\]","\[\/u\]","u","<u>$1</u>")
		If InStr(Lcase(s),"[username")>0 Then s= Dv_UbbCode_name(s)
		'不开放HTML支持，不转换HREF
		If Dvbbs.Board_Setting(5)="1" Then
			'自动识别网址
			If InStr(Ubblists,",18,")>0 Or InStr(Ubblists,",19,")>0 Or InStr(Ubblists,",20,")>0 Or InStr(Ubblists,",21,")>0 Or InStr(Ubblists,",22,")>0 Or sType=1 Then
				re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+([^<>""])+)"
				s = re.Replace(s,"<a target=_blank href=$1>$1</a>")
				re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+([^<>""])+)$([^\[]*)"
				s = re.Replace(s,"<a target=_blank href=$1>$1</a>")
				re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+([^<>""])+)"
				s = re.Replace(s,"$1<a target=_blank href=$2>$2</a>")
			End If
			'自动识别www等开头的网址
			If InStr(Ubblists,",38,")>0 Or sType=1 Then
				re.Pattern = "([\s])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
				s = re.Replace(s,"<a target=_blank href=""http://$2"">$2</a>")
			End If
		End If
		Dv_UbbCode=bbimg(s,500)
		Set Re=Nothing
	End Function
	Private Function bbimg(strText,ssize)
		Dim s
		s=strText
		re.Pattern="<img(.[^>]*)>"
		If ssize=500 Then
			s=re.replace(s,"<img$1 onmousewheel=""return bbimg(this)"" onload=""javascript:if(this.width>screen.width-"&ssize&")this.style.width=screen.width-"&ssize&";"">")
		Else
			s=re.replace(s,"<img$1 onmousewheel=""return bbimg(this)"" onload=""javascript:if(this.width>screen.width-"&ssize&")this.style.width=screen.width-"&ssize&";if(this.height>250)this.style.width=(this.width*250)/this.height;"">")
		End If
		bbimg=s
	End Function
	'签名UBB转换
	Public Function Dv_SignUbbCode(s,PostUserGroup)
		Dim ii
		Dim po
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		If Dvbbs.forum_setting(66)="0" Then 
			s= server.htmlEncode(s)
			If Dv_FilterJS2(s)Then
				re.Pattern="\[(br)\]"
				s=re.Replace(s,"<$1>")
				re.Pattern = "(&nbsp;)"
				s = re.Replace(s,Chr(9))
				re.Pattern = "(<br>)"
				s = re.Replace(s,vbNewLine)
				re.Pattern = "(<p>)"
				s = re.Replace(s,"")
				re.Pattern = "(<\/p>)"
				s = re.Replace(s,vbNewLine)
				s=server.htmlencode(s)
				s="<form name=""scode"&replyid&""" method=""post"" action=""""><TABLE class=tableborder2 cellSpacing=1 cellPadding=3 width=""100%"" align=center border=0><TR><TH height=22>以下内容含脚本,或可能导致页面不正常的代码</TH></TR><TR><TD class=tablebody1 align=middle width=""98%""><TEXTAREA id=CodeText style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDTH: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=20 cols=120>"&s&"</TEXTAREA></TD></TR><TR><TD class=tablebody2 align=middle width=""98%""><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行.</td></tr><TR><TD class=tablebody1 align=middle width=""98%""><input type=""button"" name=""run"" value=""运行代码"" onclick=""Dvbbs_ViewCode("&replyid&");""></TD></TR></TABLE></form>"
					s = Replace(s, vbNewLine, "")
					s = Replace(s, CHR(10), "")
					s = Replace(s, CHR(13), "")
				Dv_SignUbbCode=s
				Exit Function
			End If
		Else
			If Dv_FilterJS(s) Then
				re.Pattern="\[(br)\]"
				s=re.Replace(s,"<$1>")
				re.Pattern = "(&nbsp;)"
				s = re.Replace(s,Chr(9))
				re.Pattern = "(<br>)"
				s = re.Replace(s,vbNewLine)
				re.Pattern = "(<p>)"
				s = re.Replace(s,"")
				re.Pattern = "(<\/p>)"
				s = re.Replace(s,vbNewLine)
				s=server.htmlencode(s)
				s="<form name=""scode"&replyid&""" method=""post"" action=""""><TABLE class=tableborder2 cellSpacing=1 cellPadding=3 width=""100%"" align=center border=0><TR><TH height=22>以下内容含脚本,或可能导致页面不正常的代码</TH></TR><TR><TD class=tablebody1 align=middle width=""98%""><TEXTAREA id=CodeText style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDTH: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=20 cols=120>"&s&"</TEXTAREA></TD></TR><TR><TD class=tablebody2 align=middle width=""98%""><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行.</td></tr><TR><TD class=tablebody1 align=middle width=""98%""><input type=""button"" name=""run"" value=""运行代码"" onclick=""Dvbbs_ViewCode("&replyid&");""></TD></TR></TABLE></form>"
				Dv_SignUbbCode=s
				Exit Function
			End If
			re.Pattern="<((asp|\!|%))"
			s=re.Replace(s,"&lt;$1")
			re.Pattern="(>)("&vbNewLine&")(<)"
			s=re.Replace(s,"$1$3") 
			re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
			s=re.Replace(s,"$1$3") 
		End If
		s = Replace(s, "  ", "&nbsp;&nbsp;")
		s = Replace(s, vbNewLine, "<br>")
		s = Replace(s, CHR(10), "")
		s = Replace(s, CHR(13), "")
		
		'常规设置不支持UBB代码，则退出
		If Cint(Dvbbs.Forum_setting(65))=0 Then 
			Dv_SignUbbCode=s
			Exit Function
		End If 
		'img code
		If InStr(Lcase(s),"[/img]")>0 Then s=Dv_UbbCode_S2(s,"\[IMG\]","\[\/IMG\]","IMG","<IMG SRC=""$1"" border=0 >","<IMG SRC=""skins/default/filetype/gif.gif"" border=0><a onfocus=this.blur() href=""$1"" target=_blank>$1</a>",PostUserGroup,Cint(Dvbbs.forum_setting(67)))
		'media code
		If InStr(Lcase(s),"[/sound]")>0 Then s=Dv_UbbCode_S2(s,"\[sound\]","\[\/sound\]","sound","<a href=""$1"" target=_blank><IMG SRC=skins/default/filetype/mid.gif border=0 alt=""背景音乐""></a><bgsound src=""$1"" loop=""-1"">","<a href=$1 target=_blank>$1</a>",PostUserGroup,Cint(Dvbbs.Board_Setting(9)))
		'flash code
		If InStr(Lcase(s),"[/flash]")>0 Then
			s=Dv_UbbCode_S2(s,"\[FLASH\]","\[\/FLASH\]","FLASH","<a href=""$1"" TARGET=_blank><IMG SRC=skins/default/filetype/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$1""><PARAM NAME=quality VALUE=high><embed src=""$1"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$1</embed></OBJECT>","<IMG SRC="&Dvbbs.Forum_info(7)&"swf.gif border=0><a href=$1 target=_blank>$1</a>（注意：Flash内容可能含有恶意代码）",PostUserGroup,Cint(Dvbbs.forum_setting(71)))
			s=Dv_UbbCode_iS2(s,"\[FLASH=(.[^\[]*)\]","\[\/FLASH\]","FLASH","<a href=""$3"" TARGET=_blank><IMG SRC=skins/default/filetype/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$1 height=$2><PARAM NAME=movie VALUE=""$3""><PARAM NAME=quality VALUE=high><embed src=""$3"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$1 height=$2>$3</embed></OBJECT>","<a href=$3 target=_blank>$3</a>（注意：Flash内容可能含有恶意代码）",PostUserGroup,Cint(Dvbbs.forum_setting(71)),"=*([0-9]*),*([0-9]*)")
		End If
		'url code
		If InStr(Lcase(s),"[/url]")>0 Then
			s=Dv_UbbCode_S1(s,"\[URL\]","\[\/URL\]","URL","<A HREF=""$1"" TARGET=_blank>$1</A>")
			re.Pattern="(\[URL=(.[^:\/\/|\[]*)\])(.[^\[]*)(\[\/URL\])"
			s= re.Replace(s,"<A HREF=""http://$2"" TARGET=_blank>$3</A>")
			re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
			s= re.Replace(s,"<A HREF=""$2"" TARGET=_blank>$3</A>")
		End If
		'email code
		If InStr(Lcase(s),"[/email]")>0 Then
			s=Dv_UbbCode_S1(s,"\[EMAIL\]","\[\/EMAIL\]","EMAIL","<img align=absmiddle src=skins/default/filetype/email1.gif ><A HREF=""mailto:$1"">$1</A>")
			s=Dv_UbbCode_iS1(s,"\[EMAIL=(.[^\[]*)\]","\[\/EMAIL\]","EMAIL","<img align=absmiddle src=skins/default/filetype/email1.gif ><A HREF=""mailto:$1"">$2</A>","=(.[^\[]*)")
		End If
		If InStr(Lcase(s),"[/html]")>0 Then s=Dv_UbbCode_S1(s,"\[HTML\]","\[\/HTML\]","HTML","<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""6"" class="""&abgcolor&"""><td><b>以下内容为程序代码:</b><br>$1</td></table>")
		If InStr(Lcase(s),"[/color]")>0 Then s=Dv_UbbCode_C(s)
		If InStr(Lcase(s),"[/face]")>0 Then s=Dv_UbbCode_F(s)
		If InStr(Lcase(s),"[/align]")>0 Then s=Dv_UbbCode_Align(s)

		If InStr(Lcase(s),"[/shadow]")>0 Then s=Dv_UbbCode_iS1(s,"\[SHADOW=(.[^\[]*)\]","\[\/SHADOW\]","SHADOW","<table width=$1 ><tr><td style=""filter:shadow(color=$2, strength=$3)"">$4</td></tr></table>","=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)")
		If InStr(Lcase(s),"[/glow]")>0 Then s=Dv_UbbCode_iS1(s,"\[GLOW=(.[^\[]*)\]","\[\/GLOW\]","GLOW","<table width=$1 ><tr><td style=""filter:glow(color=$2, strength=$3)"">$4</td></tr></table>","=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)")
		If InStr(Lcase(s),"[/i]")>0 Then s=Dv_UbbCode_S1(s,"\[i\]","\[\/i\]","i","<i>$1</i>")
		If InStr(Lcase(s),"[/b]")>0 Then s=Dv_UbbCode_S1(s,"\[b\]","\[\/b\]","b","<b>$1</b>")
		If InStr(Lcase(s),"[/u]")>0 Then s=Dv_UbbCode_S1(s,"\[u\]","\[\/u\]","u","<u>$1</u>")
		If InStr(Lcase(s),"[/size]")>0 Then s=Dv_UbbCode_S(s)
		REM ：签名移动(如需使用则把以下屏蔽去掉)
		'If InStr(Lcase(s),"[/fly]")>0 Then s=Dv_UbbCode_S1(s,"\[fly\]","\[\/fly\]","fly","<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
		'If InStr(Lcase(s),"[/move]")>0 Then s=Dv_UbbCode_S1(s,"\[move\]","\[\/move\]","move","<MARQUEE scrollamount=3>$1</marquee>")
		REM 不开放HTML支持，不转换HREF
		REM 加上签名是否开放HTML判断 2004-5-6 Dvbbs.YangZheng
		If Dvbbs.Board_Setting(5)="1" And Dvbbs.Forum_Setting(66) = "1" Then
			'自动识别网址
			If InStr(Lcase(s),"http://")>0 Then
				re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+([^<>""])+)"
				s = re.Replace(s,"<a target=_blank href=$1>$1</a>")
				re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+([^<>""])+)$"
				s = re.Replace(s,"<a target=_blank href=$1>$1</a>")
				re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+([^<>""])+)"
				s = re.Replace(s,"$1<a target=_blank href=$2>$2</a>")
			End If
			'自动识别www等开头的网址
			If InStr(Lcase(s),"www.")>0 Then
				re.Pattern = "([^@|^\.|^\/\/])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
				s = re.Replace(s,"<a target=_blank href=http://$2>$2</a>")
			End If
		End If
		s=bbimg(s,600)
		Dv_SignUbbCode=s
		Set Re=Nothing
	End Function
	Private Function Dv_UbbCode_S1(strText,uCodeL,uCodeR,uCodeC,tCode)
		Dim s
		s=strText
		re.Pattern=uCodeL
		s=re.replace(s, chr(1) & uCodeC & chr(2))
		re.Pattern=uCodeR
		s=re.replace(s, chr(1) & "/" & uCodeC & chr(2))
		re.Pattern="\x01"&uCodeC&"\x02\x01\/"&uCodeC&"\x02"
		s=re.Replace(s,"")
		re.Pattern="\x01"&uCodeC&"\x02(.[^\x01]*)\x01\/"&uCodeC&"\x02"
		s=re.Replace(s,tCode)
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		Dv_UbbCode_S1=s
	End Function
	Private Function Dv_UbbCode_iS1(strText,uCodeL,uCodeR,uCodeC,tCode,iCode)
		Dim s
		s=strText
		re.Pattern=uCodeL
		s=re.replace(s, chr(1) & uCodeC & "=$1" & chr(2))
		re.Pattern=uCodeR
		s=re.replace(s, chr(1) & "/" & uCodeC & chr(2))
		re.Pattern="\x01"&uCodeC&iCode&"\x02\x01\/"&uCodeC&"\x02"
		s=re.Replace(s,"")
		re.Pattern="\x01"&uCodeC&iCode&"\x02(.[^\x01]*)\x01\/"&uCodeC&"\x02"
		s=re.Replace(s,tCode)
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		Dv_UbbCode_iS1=s
	End Function

	Private Function Dv_UbbCode_S2(strText,uCodeL,uCodeR,uCodeC,tCode1,tCode2,PostUserGroup,Flag)
		Dim s
	
		s=strText
		If Dvbbs.Forum_Setting(76)="" Or  Dvbbs.Forum_Setting(76)="" Then
			re.Pattern="UploadFile/"
			s=re.replace(s,Dvbbs.Forum_Setting(76))
		End If
		re.Pattern=uCodeL
		s=re.replace(s, chr(1) & uCodeC & chr(2))
		re.Pattern=uCodeR
		s=re.replace(s, chr(1) & "/" & uCodeC & chr(2))
		re.Pattern="\x01"&uCodeC&"\x02(.[^\x01]*)\x01\/"&uCodeC&"\x02"
		If Flag = 1 or PostUserGroup<4 Then
			s=re.Replace(s,tCode1)
		Else
			s=re.Replace(s,tCode2)
		End If 
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		Dv_UbbCode_S2=s
	End Function
	Private Function Dv_UbbCode_iS2(strText,uCodeL,uCodeR,uCodeC,tCode1,tCode2,PostUserGroup,Flag,iCode)
		Dim s
		s=strText
		re.Pattern=uCodeL
		s=re.replace(s, chr(1) & uCodeC & "=$1" & chr(2))
		re.Pattern=uCodeR
		s=re.replace(s, chr(1) & "/" & uCodeC & chr(2))
		Rem 过滤媒体标签中的FLASH 2004-4-29 Dvbbs.YangZheng
		If Instr(uCodeC,"FLASH") = 0 Then
			re.Pattern="(.(swf|swi))"
			s=re.Replace(s,"")
		End If
		re.Pattern="\x01"&uCodeC&iCode&"\x02(.[^\x01]*)\x01\/"&uCodeC&"\x02"
		If Flag = 1 or PostUserGroup<4 Then
			s=re.Replace(s,tCode1)
		Else
			s=re.Replace(s,tCode2)
		End If 
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		Dv_UbbCode_iS2=s
	End Function

	Private Function Dv_UbbCode_Align(strText)
		Dim s
		s=strText
		re.Pattern="\[ALIGN=(center|left|right)\]"
		s=re.replace(s, chr(1) & "ALIGN=$1" & chr(2))
		re.Pattern="\[\/ALIGN\]"
		s=re.replace(s, chr(1) & "/ALIGN" & chr(2))
		re.Pattern="\x01ALIGN=(center|left|right)\x02\x01\/ALIGN\x02"
		s=re.Replace(s,"")
		re.Pattern="\x01ALIGN=(center|left|right)\x02(.[^\x01]*)\x01\/ALIGN\x02"
		s=re.Replace(s,"<div align=$1>$2</div>")
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		Dv_UbbCode_Align=s
	End Function

	Private Function Dv_UbbCode_U(strText,PostUserGroup,Flag)	'(帖子内容，用户组，是否开放图片标签)
		Dim s
		If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
		If right(Dvbbs.Forum_Setting(76),1)<>"/" Then Dvbbs.Forum_Setting(76)=Dvbbs.Forum_Setting(76)&"/"
		s=strText
		're.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp|png)\]UploadFile/"
		re.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp|png|swf|swi)\]UploadFile/"
		s=re.replace(s, chr(1) & "UPLOAD=$1" & chr(2))
		re.Pattern="\[\/UPLOAD\]"
		s=re.replace(s, chr(1) & "/UPLOAD" & chr(2))
		re.Pattern="\x01UPLOAD=(gif|jpg|jpeg|bmp|png)\x02\x01\/UPLOAD\x02"
		s=re.Replace(s,"")
		re.Pattern="\x01UPLOAD=(gif|jpg|jpeg|bmp|png)\x02(.[^\x01]*)\x01\/UPLOAD\x02"
		If Dvbbs.Forum_Setting(75)="0" Then 
			If Flag = 1 or PostUserGroup<4 Then
				s= re.Replace(s,"<br><IMG SRC=""skins/default/filetype/$1.gif"" border=0 >此主题相关图片如下：<br><A HREF="""&Dvbbs.Forum_Setting(76)&"$2"" TARGET=_blank id=""ImgSpan""><IMG SRC="""&Dvbbs.Forum_Setting(76)&"$2"" border=0 alt=按此在新窗口浏览图片 ></A>")
			Else
				s= re.Replace(s,"<br><IMG SRC=""skins/default/filetype/$1.gif"" border=0 ><A HREF="""&Dvbbs.Forum_Setting(76)&"$2"" TARGET=_blank>"&Dvbbs.Forum_Setting(76)&"$2</A>")
			End If 
		Else
			If Flag = 1 or PostUserGroup<4 Then
				s= re.Replace(s,"<br><IMG SRC=""skins/default/filetype/$1.gif"" border=0 >此主题相关图片如下：<br><A HREF=""showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2"" TARGET=_blank id=""ImgSpan"" ><IMG SRC=""showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2"" border=0 alt=按此在新窗口浏览图片 ></A>")
			Else
				s= re.Replace(s,"<br><IMG SRC=""skins/default/filetype/$1.gif"" border=0 ><A HREF=""showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2"" TARGET=_blank>showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2</A>")
			End If
		End If
		re.Pattern="\x01UPLOAD=(swf|swi)\x02(.[^\x01]*)\x01\/UPLOAD\x02"
		If Dvbbs.Forum_Setting(75)="0" Then 
			If Dvbbs.Board_Setting(44) = 1 or PostUserGroup<4 Then
				s= re.Replace(s,"<br><IMG SRC=""skins/default/filetype/swf.gif"" border=0 ><A HREF="""&Dvbbs.Forum_Setting(76)&"$2"" TARGET=_blank>点击浏览该FLASH文件</A>：<br><embed src="""&Dvbbs.Forum_Setting(76)&"$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=300></embed>")
			Else
				s= re.Replace(s,"<br><IMG SRC=""skins/default/filetype/swf.gif"" border=0 ><A HREF="""&Dvbbs.Forum_Setting(76)&"$2"" TARGET=_blank>"&Dvbbs.Forum_Setting(76)&"$2</A>")
			End If 
		Else
			If Flag = 1 or PostUserGroup<4 Then
				s= re.Replace(s,"<br><IMG SRC=""skins/default/filetype/swf.gif"" border=0 >点击浏览该FLASH文件：<br><A HREF=""showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2"" TARGET=_blank id=""ImgSpan"" ><IMG SRC=""showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2"" border=0 alt=按此在新窗口浏览图片 ></A>")
			Else
				s= re.Replace(s,"<br><IMG SRC=""skins/default/filetype/$1.gif"" border=0 ><A HREF=""showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2"" TARGET=_blank>showimg.asp?BoardID="&Dvbbs.BoardID&"&filename=$2</A>")
			End If
		End If
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		re.Pattern="\[UPLOAD=(.[^\[]*)\]"
		s=re.replace(s, chr(1) & "UPLOAD=$1" & chr(2))
		re.Pattern="\[\/UPLOAD\]"
		s=re.replace(s, chr(1) & "/UPLOAD" & chr(2))
		re.Pattern="\x01UPLOAD=(.[^\x01]*)\x02\x01\/UPLOAD\x02"
		s=re.Replace(s,"")
		re.Pattern="\x01UPLOAD=(.[^\x01]*)\x02(viewFile\.asp.[^\x01]*)\x01\/UPLOAD\x02"
		s= re.Replace(s,"<br><IMG SRC=""skins/default/filetype/$1.gif"" border=0> <a href=""$2"" target=_blank>点击浏览该文件</a>")
		re.Pattern="viewFile.asp\?"
		s= re.Replace(s,"viewFile.asp?Boardid="&Dvbbs.Boardid&"&")
		re.Pattern="\x01UPLOAD=(.[^\x01]*)\x02(.[^\x01]*)\x01\/UPLOAD\x02"
		s= re.Replace(s,"<br><IMG SRC=""skins/default/filetype/$1.gif"" border=0> <a href=""$2"" target=_blank>点击浏览该文件</a><br><IMG src=""$2"" border=0 >")
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		Dv_UbbCode_U=s
	End Function
	Private Function Dv_UbbCode_S(strText)
		Dim s
		Dim Test
		Dim LoopCount
		LoopCount=0
		s=strText
		Do While True
			re.Pattern="\[SIZE=([1-7])\]"
			Test=re.Test(s)
			If Test Then
				re.Pattern="\[\/SIZE\]"
				Test=re.Test(s)
				If Test Then
					re.Pattern="\[SIZE=([1-7])\]"
					s=re.replace(s, chr(1) & "SIZE=$1" & chr(2))
					re.Pattern="\[\/SIZE\]"
					s=re.replace(s, chr(1) & "/SIZE" & chr(2))
					re.Pattern="\x01SIZE=([1-7])\x02\x01\/SIZE\x02"
					s=re.Replace(s,"")
					re.Pattern="\x01SIZE=([1-7])\x02(.[^\x01]*)\x01\/SIZE\x02"
					s=re.Replace(s,"<font size=$1>$2</font>")
					re.Pattern="\x02"
					s=re.replace(s, "]")
					re.Pattern="\x01"
					s=re.replace(s, "[")
				Else
					Exit Do
				End If 
			Else
				Exit Do
			End If
			 LoopCount=LoopCount+1
			 If LoopCount>MaxLoopCount Then Exit Do
		Loop
		Dv_UbbCode_S=s
	End Function

	Private Function Dv_UbbCode_Q(strText)
		Dim s
		Dim Test
		Dim LoopCount
		LoopCount=0
		s=strText
		re.Pattern="\[QUOTE\]"
		Test=re.Test(s)
		If Test Then
			re.Pattern="\[\/QUOTE\]"
			Test=re.Test(s)
			If Test Then
				re.Pattern="\[QUOTE\]"
				s=re.replace(s, chr(1) & "QUOTE" & chr(2))
				re.Pattern="\[\/QUOTE\]"
				s=re.replace(s, chr(1) & "/QUOTE" & chr(2))
				Do
					re.Pattern="\x01QUOTE\x02\x01\/QUOTE\x02"
					s=re.Replace(s,"")
					re.Pattern="\x01QUOTE\x02(.[^\x01]*)\x01\/QUOTE\x02"
					s=re.Replace(s,"<DIV class=quote>$1</div><br>")
					Test=re.Test(s)
					LoopCount=LoopCount+1
					If LoopCount>MaxLoopCount Then Exit Do
				Loop While(Test)
				re.Pattern="\x02"
				s=re.replace(s, "]")
				re.Pattern="\x01"
				s=re.replace(s, "[")
			End If 
		End If
		Dv_UbbCode_Q=s
	End Function

	Private Function Dv_UbbCode_C(strText)
		Dim s
		Dim Test
		Dim LoopCount
		LoopCount=0
		s=strText
		Do While True
			re.Pattern="\[COLOR=(.[^\[]*)\]"
			Test=re.Test(s)
			If Test Then
				re.Pattern="\[\/COLOR\]"
				Test=re.Test(s)
				If Test Then
					re.Pattern="\[COLOR=(.[^\[]*)\]"
					s=re.replace(s, chr(1) & "COLOR=$1" & chr(2))
					re.Pattern="\[\/COLOR\]"
					s=re.replace(s, chr(1) & "/COLOR" & chr(2))
					re.Pattern="\x01COLOR=(.[^\x01]*)\x02\x01\/COLOR\x02"
					s=re.Replace(s,"")
					re.Pattern="\x01COLOR=(.[^\x01]*)\x02(.[^\x01]*)\x01\/COLOR\x02"
					s=re.Replace(s,"<font color=$1>$2</font>")
					re.Pattern="\x02"
					s=re.replace(s, "]")
					re.Pattern="\x01"
					s=re.replace(s, "[")
				Else
					Exit Do
				End If
			Else
				Exit Do
			End If
			LoopCount=LoopCount+1
			If LoopCount>MaxLoopCount Then Exit Do
		Loop
		Dv_UbbCode_C=s
	End Function

	Private Function Dv_UbbCode_F(strText)
		Dim s
		Dim Test
		Dim LoopCount
		LoopCount=0
		s=strText
		Do While True
			re.Pattern="\[FACE=(.[^\[]*)\]"
			Test=re.Test(s)
			If Test Then
				re.Pattern="\[\/FACE\]"
				Test=re.Test(s)
				If Test Then
					re.Pattern="\[FACE=(.[^\[]*)\]"
					s=re.replace(s, chr(1) & "FACE=$1" & chr(2))
					re.Pattern="\[\/FACE\]"
					s=re.replace(s, chr(1) & "/FACE" & chr(2))
					re.Pattern="\x01FACE=(.[^\x01]*)\x02\x01\/FACE\x02"
					s=re.Replace(s,"")
					re.Pattern="\x01FACE=(.[^\x01]*)\x02(.[^\x01]*)\x01\/FACE\x02"
					s=re.Replace(s,"<font face=$1>$2</font>")
					re.Pattern="\x02"
					s=re.replace(s, "]")
					re.Pattern="\x01"
					s=re.replace(s, "[")
				Else
					Exit Do
				End If
			Else
				Exit Do
			End If
			LoopCount=LoopCount+1
			If LoopCount>MaxLoopCount Then Exit Do
		Loop
		Dv_UbbCode_F=s
	End Function
	Private Function Dv_UbbCode_name(strText)
		Dim s
		Dim Test,po
		s=strText
		re.Pattern="\[username=(.[^\[]*)](.[^\[]*)\[\/username\]"
		If Cint(Dvbbs.Board_Setting(56))=1 Then
			po=re.Replace(s,"[,$1,]")
			If  Dvbbs.Membername<>"" and (Dvbbs.Membername=UserName or InStr(po,","&Dvbbs.Membername&",")>0  or Dvbbs.master) Then
				s=re.Replace(s,"<hr noshade size=1><font color=red>以下内容是专门发给<B>$1</B>浏览</font><BR>$2<hr noshade size=1>")
			Else
				s=re.Replace(s,"<hr noshade size=1><font color=gray>以下内容是专门发给<B>$1</B>浏览</font><BR><hr noshade size=1>")
			End If 
		Else
			s=re.Replace(s,"$2")
		End If
		Dv_UbbCode_name=s
	End Function
	Private Function Dv_UbbCode_Get(strText,PostUserGroup,PostType,uCodeL,uCodeR,uCodeC,tCode1,tCode2,UsePoint,Flag)
		Dim s
		Dim Test
		Dim po,ii
		Dim LoopCount
		LoopCount=0
		s=strText
		UsePoint=CLng(UsePoint)
		Do While True
			re.Pattern=uCodeL
			Test=re.Test(s)
			If Test Then
				re.Pattern=uCodeR
				Test=re.Test(s)
				If Test Then
					re.Pattern=uCodeL
					s=re.replace(s, chr(1) & ""&uCodeC&"=$1" & chr(2))
					re.Pattern=uCodeR
					s=re.replace(s, chr(1) & "/"&uCodeC&"" & chr(2))
					re.Pattern="(\x01"&uCodeC&"=*([0-9]*)\x02)(\x01\/"&uCodeC&"\x02)"
					s=re.Replace(s,"")
					If (Flag=1 or PostUserGroup<4) and PostType=1 Then
						re.Pattern="(^.*)(\x01"&uCodeC&"=*([0-9]*)\x02)(.[^\x01]*)(\x01\/"&uCodeC&"\x02)(.*)"
						po=re.Replace(s,"$3")
						If  IsNumeric(po) Then
							ii=int(po) 
						Else
							ii=0
						End If 
						If  Dvbbs.Membername<>"" and (Dvbbs.Membername=UserName or UsePoint>=ii or Dvbbs.master) Then
							s=re.Replace(s,tCode1)
						Else
							s=re.Replace(s,tCode2)
						End If
					Else
						re.Pattern="(\x01"&uCodeC&"=*([0-9]*)\x02)(.[^\x01]*)(\x01\/"&uCodeC&"\x02)"
						s=re.Replace(s,"$3")
					End If 
					re.Pattern="\x02"
					s=re.replace(s, "]")
					re.Pattern="\x01"
					s=re.replace(s, "[")
				Else
					Exit Do
				End If 
			Else
				Exit Do
			End If
			LoopCount=LoopCount + 1
			If LoopCount>MaxLoopCount Then Exit Do
		Loop
		Dv_UbbCode_Get=s
	End Function

	Private Function UBB_REPLYVIEW(strText,PostUserGroup,PostType)
		Dim s
		Dim Test
		Dim vrs
		
		s=strText
		re.Pattern="\[REPLYVIEW\]"
		s=re.replace(s, chr(1) & "REPLYVIEW" & chr(2))
		re.Pattern="\[\/REPLYVIEW\]"
		s=re.replace(s, chr(1) & "/REPLYVIEW" & chr(2))
		re.Pattern="(\x01REPLYVIEW\x02)(\x01\/REPLYVIEW\x02)"
		s=re.Replace(s,"")
		re.Pattern="(\x01REPLYVIEW\x02)(.[^\x01]*)(\x01\/REPLYVIEW\x02)"
		If (Dvbbs.Board_Setting(15)="1" or PostUserGroup<4) and PostType=1  Then
			If isgetreed<>1 Then 
				Set vrs=dvbbs.execute("select AnnounceID from "&TotalUseTable&" where rootid="&Announceid&" and PostUserID="&Dvbbs.UserID)
				isgetreed=1
				If Not vRs.eof Then
					reed=1 
				Else
					reed=0
				End If
				Set vrs=Nothing
			End If 
			If Dvbbs.Membername<>"" and (reed=1 or Dvbbs.master) Then
				s=re.Replace(s,"<hr noshade size=1><font color=gray>以下内容只有<B>回复</B>后才可以浏览</font><BR>$2<hr noshade size=1>")
			Else
				s=re.Replace(s,"<hr noshade size=1><font color="&Dvbbs.Mainsetting(1)&">以下内容只有<B>回复</B>后才可以浏览</font><hr noshade size=1>")
			End If 
		Else
			s=re.Replace(s,"$2")
		End If 
		re.Pattern="\x02"
		s=re.replace(s, "]")
		re.Pattern="\x01"
		s=re.replace(s, "[")
		UBB_REPLYVIEW=s
	End Function

	Private Function UBB_USEMONEY(strText,PostUserGroup,PostType)
		Dim s
		Dim Test
		Dim po,ii,iii
		Dim SplitBuyUser,iPostBuyUser
		Dim LoopCount
		LoopCount=0
		s=strText
		Do While True
			re.Pattern="\[USEMONEY=*([0-9]*)\]"
			Test=re.Test(s)
			If Test Then
				re.Pattern="\[\/USEMONEY\]"
				Test=re.Test(s)
				If Test Then
					re.Pattern="\[USEMONEY=*([0-9]*)\]"
					s=re.replace(s, chr(1) & "USEMONEY=$1" & chr(2))
					re.Pattern="\[\/USEMONEY\]"
					s=re.replace(s, chr(1) & "/USEMONEY" & chr(2))
					re.Pattern="(\x01USEMONEY=*([0-9]*)\x02)(\x01\/USEMONEY\x02)"
					s=re.Replace(s,"")
					If (Cint(Dvbbs.Board_Setting(23))=1 or PostUserGroup<4) and PostType=1 Then
						re.Pattern="(^.*)(\x01USEMONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)(.*)"
						po=re.Replace(s,"$3")
						If  IsNumeric(po) Then
							ii=int(po) 
						Else
							ii=0
						End If 
						If  Dvbbs.Membername<>"" and (Dvbbs.Membername=UserName or Dvbbs.master) Then
							If (Not IsNull(PostBuyUser)) And PostBuyUser<>"" Then
								SplitBuyUser=split(PostBuyUser,"|")
								iPostBuyUser="<option value=0>已购买用户</option>"
								for iii=0 to ubound(SplitBuyUser)
									iPostBuyUser=iPostBuyUser & "<option value="&iii&">"&SplitBuyUser(iii)&"</option>"
								next
							Else
								iPostBuyUser="<option value=0>还没有用户购买</option>"
							End If 
							s=re.Replace(s,"$1<hr noshade size=1><font color=gray>以下内容需要花费现金<B>$3</B>才可以浏览</font>&nbsp;&nbsp;<select size=1 name=buyuser>"&iPostBuyUser&"</select><BR>$4<hr noshade size=1>$6")
						Else
							If (Not IsNull(PostBuyUser)) and PostBuyUser<>"" Then
								If  Instr("|"&PostBuyUser&"|","|"&Dvbbs.Membername&"|")>0 Then
									s=re.Replace(s,"$1<hr noshade size=1><font color=gray>以下内容需要花费现金<B>$3</B>才可以浏览，您已经购买本帖</font><BR>$4<hr noshade size=1>$6")
								Else
									If UserPointInfo(0)>=ii Then
										s=re.Replace(s,"$1<Form action=""BuyPost.asp"" mothod=post><font color="&Dvbbs.Mainsetting(1)&">以下内容需要花费现金<B>$3</B>才可以浏览&nbsp;&nbsp;<input type=hidden name=boardid value="&Dvbbs.boardid&"><input type=hidden value="&replyid&" name=replyid><input type=hidden value="&AnnounceID&" name=id><input type=hidden value="&totalusetable&" name=posttable><input type=submit name=submit value=好黑啊…我…我买了！>&nbsp;&nbsp;</font></form>$6")
									Else
										s=re.Replace(s,"$1<hr noshade size=1><font color="&Dvbbs.Mainsetting(1)&">以下内容需要花费现金<B>$3</B>才可以浏览，您没有这么多现金</font><hr noshade size=1>$6")
									End If 
								End If 
							Else
								If UserPointInfo(0)>=ii Then
									s=re.Replace(s,"$1<Form action=""BuyPost.asp"" mothod=post><font color="&Dvbbs.Mainsetting(1)&">以下内容需要花费现金<B>$3</B>才可以浏览&nbsp;&nbsp;<input type=hidden name=boardid value="&Dvbbs.boardid&"><input type=hidden value="&replyid&" name=replyid><input type=hidden value="&AnnounceID&" name=id><input type=hidden value="&totalusetable&" name=posttable><input type=submit name=submit value=好黑啊…我…我买了！>&nbsp;&nbsp;</font></form>$6")
								Else
									s=re.Replace(s,"$1<hr noshade size=1><font color="&Dvbbs.Mainsetting(1)&">以下内容需要花费现金<B>$3</B>才可以浏览，您没有这么多现金</font><hr noshade size=1>$6")
								End If 
							End If 
						End If 
						re.Pattern="([^>=""])(\x01USEMONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)(.*)"
						po=re.Replace(s,"$3")
						If  IsNumeric(po) Then
							ii=int(po) 
						Else
							ii=0
						End If 
						If  Dvbbs.Membername<>"" and (Dvbbs.Membername=UserName or Dvbbs.master) Then
							If (Not IsNull(PostBuyUser)) And PostBuyUser<>"" Then
								SplitBuyUser=split(PostBuyUser,"|")
								iPostBuyUser="<option value=0>已购买用户</option>"
								for iii=0 to ubound(SplitBuyUser)
									iPostBuyUser=iPostBuyUser & "<option value="&iii&">"&SplitBuyUser(iii)&"</option>"
								next
							Else
								iPostBuyUser="<option value=0>还没有用户购买</option>"
							End If 
							s=re.Replace(s,"$1<hr noshade size=1><font color=gray>以下内容需要花费现金<B>$3</B>才可以浏览</font>&nbsp;&nbsp;<select size=1 name=buyuser>"&iPostBuyUser&"</select><BR>$4<hr noshade size=1>$6")
						Else
							If (Not IsNull(PostBuyUser)) and PostBuyUser<>"" Then
								If  Instr("|"&PostBuyUser&"|","|"&Dvbbs.Membername&"|")>0 Then
									s=re.Replace(s,"$1<hr noshade size=1><font color=gray>以下内容需要花费现金<B>$3</B>才可以浏览，您已经购买本帖</font><BR>$4<hr noshade size=1>$6")
								Else
									If UserPointInfo(0)>=ii Then
										s=re.Replace(s,"$1<Form action=""BuyPost.asp"" mothod=post><font color="&Dvbbs.Mainsetting(1)&">以下内容需要花费现金<B>$3</B>才可以浏览&nbsp;&nbsp;<input type=hidden name=boardid value="&Dvbbs.boardid&"><input type=hidden value="&replyid&" name=replyid><input type=hidden value="&AnnounceID&" name=id><input type=hidden value="&totalusetable&" name=posttable><input type=submit name=submit value=好黑啊…我…我买了！>&nbsp;&nbsp;</font></form>$6")
									Else
										s=re.Replace(s,"$1<hr noshade size=1><font color="&Dvbbs.Mainsetting(1)&">以下内容需要花费现金<B>$3</B>才可以浏览，您没有这么多现金</font><hr noshade size=1>$6")
									End If 
								End If 
							Else
								If UserPointInfo(0)>=ii Then
									s=re.Replace(s,"$1<Form action=""BuyPost.asp"" mothod=post><font color="&Dvbbs.Mainsetting(1)&">以下内容需要花费现金<B>$3</B>才可以浏览&nbsp;&nbsp;<input type=hidden name=boardid value="&Dvbbs.boardid&"><input type=hidden value="&replyid&" name=replyid><input type=hidden value="&AnnounceID&" name=id><input type=hidden value="&totalusetable&" name=posttable><input type=submit name=submit value=好黑啊…我…我买了！>&nbsp;&nbsp;</font></form>$6")
								Else
									s=re.Replace(s,"$1<hr noshade size=1><font color="&Dvbbs.Mainsetting(1)&">以下内容需要花费现金<B>$3</B>才可以浏览，您没有这么多现金</font><hr noshade size=1>$6")
								End If 
							End If 
						End If 
					Else
						re.Pattern="(\x01USEMONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)"
						s=re.Replace(s,"$3")
					End If 
					re.Pattern="\x02"
					s=re.replace(s, "]")
					re.Pattern="\x01"
					s=re.replace(s, "[")
				Else
					Exit Do
				End If 
			Else
				Exit Do
			End If 
			LoopCount=LoopCount+1
			If LoopCount>MaxLoopCount Then Exit Do
		Loop
		UBB_USEMONEY=s
	End Function
	Private Function dv_fixHTML(strText)
		Dim s
		s=strText
		If InStr(Ubblists,",39,")>0 And (InStr(Ubblists,",table,")>0 Or InStr(Ubblists,",td,")>0 Or InStr(Ubblists,",th,")>0 Or InStr(Ubblists,",tr,")>0 ) Then
			s = server.htmlencode(s)
			s="<form name=""scode"&replyid&""" method=""post"" action=""""><TABLE class=tableborder2 cellSpacing=1 cellPadding=3 width=""100%"" align=center border=0><TR><TH height=22>以下内容含错误标记</TH></TR><TR><TD class=tablebody1 align=middle width=""98%""><TEXTAREA id=CodeText style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDTH: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=20 cols=120>"&s&"</TEXTAREA></TD></TR><TR><TD class=tablebody2 align=middle width=""98%""><input type=""button"" name=""run"" value=""在新窗口中查看"" onclick=""Dvbbs_ViewCode("&replyid&");""></TD></TR></TABLE></form>"
		End If
		dv_fixHTML=s
	End Function
	Public Function Dv_FilterJS(v)
		If Not Isnull(V) Then
			Dim t,test,Replacelist,t1
			t=v
			t1=v
			re.Pattern="&#36;"
			t1=re.Replace(t1,"$")
			re.Pattern="&#36"
			t1=re.Replace(t1,"$")
			re.Pattern="&#39;"
			t1=re.Replace(t1,"'")
			re.Pattern="&#39"
			t1=re.Replace(t1,"'")
			If InStr(Dvbbs.forum_setting(77),"|")=0 Then 
				Replacelist="(&#([0-9][0-9]*)|function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"
			Else
				Replacelist="("&Dvbbs.forum_setting(77)&"&#([0-9][0-9]*)|function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit|error|click|key|load|focus|Blur))"
			End If
			re.Pattern="<((.[^>]*"&Replacelist&"[^>]*)|"&Replacelist&")>"
			Test=re.Test(t1)
			If Test=false Then
				re.Pattern="(\[(.[^\]]*)\])((.[^\]]*"&Replacelist&"[^\]]*)|"&Replacelist&")(\[\/(.[^\]]*)\])"
				Test=re.Test(t1)
			End If
			Dv_FilterJS=test
		End If
	End Function
	Public Function Dv_FilterJS2(v)
		If Not Isnull(V) Then
			Dim t,test,Replacelist,t1
			t=v
			t1=v
			re.Pattern="&#36;"
			t1=re.Replace(t1,"$")
			re.Pattern="&#36"
			t1=re.Replace(t1,"$")
			re.Pattern="&#39;"
			t1=re.Replace(t1,"'")
			re.Pattern="&#39"
			t1=re.Replace(t1,"'")
			If InStr(Dvbbs.forum_setting(77),"|")=0 Then 
				Replacelist="(&#([0-9][0-9]*)|function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"
			Else
				Replacelist="("&Dvbbs.forum_setting(77)&"&#([0-9][0-9]*)|function|meta|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit|error|click|key|load|focus|Blur))"
			End If
			re.Pattern="(\[(.[^\]]*)\])((.[^\]]*"&Replacelist&"[^\]]*)|"&Replacelist&")(\[\/(.[^\]]*)\])"
			Test=re.Test(t1)
			Dv_FilterJS2=test
		End If
	End Function
End Class
</script>