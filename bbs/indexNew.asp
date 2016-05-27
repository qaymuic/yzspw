<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dim TempArray
Dvbbs.BoardID=0
Dvbbs.LoadTemplates("index")
Dvbbs.Stats=template.Strings(0)
Dvbbs.Nav()
Dim BrowserType,TempStr
Set BrowserType=New Cls_Browser
If Dvbbs.Forum_ads(2)="1" or Dvbbs.Forum_ads(13)="1"  Then Response.Write template.html(13)
Dvbbs.ActiveOnline()
TempArray = Split(template.html(3),"||")
Show_Index_Top
GetsmBbsList()
'GetBbsList()
Response.Write Replace(template.html(9),"{$Getlink}",Getlink())
If Dvbbs.Forum_setting(29)="1" Then Call birthuser()
Show_Index_Footer
Set BrowserType=Nothing
Dvbbs.Footer()
Sub Show_Index_Top
	Dim newsstr,TempStr,TopArray
	newsstr=news
	If newsstr(1)="" Or Not IsDate(newsstr(1)) Then newsstr(1)=Now()
	TempStr = template.html(0)
	TopArray = Split(template.html(2),"||")
	TempStr=Replace(TempStr,"{$news}",newsstr(0))
	TempStr=Replace(TempStr,"{$newstime}",newsstr(1))
	TempStr=Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	TempStr=Replace(TempStr,"{$UserNum}",Dvbbs.CacheData(10,0))
	TempStr=Replace(TempStr,"{$lastUser}",Dvbbs.HtmlEncode(Dvbbs.CacheData(14,0)))
	TempStr=Replace(TempStr,"{$TodayNum}",Dvbbs.CacheData(9,0))
	TempStr=Replace(TempStr,"{$TopicNum}",Dvbbs.CacheData(7,0))
	TempStr=Replace(TempStr,"{$YesTerdayNum}",Dvbbs.CacheData(11,0))
	TempStr=Replace(TempStr,"{$PostNum}",Dvbbs.CacheData(8,0))
	TempStr=Replace(TempStr,"{$MaxPostNum}",Dvbbs.CacheData(12,0))
	TempStr=Replace(TempStr,"{$MaxPostDate}",Dvbbs.CacheData(13,0))
	If Dvbbs.UserID=0 Then
		TempStr=Replace(TempStr,"{$myinfo}",Replace(TopArray(0),"{$forumname}",Dvbbs.Forum_Info(0)))
		If Dvbbs.Forum_ChanSetting(0)="1" Then TempStr=Replace(TempStr,"{$isray}",TopArray(1))
		TempStr=Replace(TempStr,"{$isray}","")
	Else
		TopArray = Split(Dvbbs.mainhtml(12),"||")
		If Clng(Dvbbs.SendMsgNum)>0 Then
			Dim UserMsg
			UserMsg = TopArray(0)
			If Dvbbs.Forum_Setting(10)="1" Then
				UserMsg = UserMsg & TopArray(1) & TopArray(2)
			Else
				UserMsg = UserMsg & TopArray(2)
			End If
			UserMsg = Replace(UserMsg,"{$smsid}",Dvbbs.sendmsgid)
			UserMsg = Replace(UserMsg,"{$sender}",Dvbbs.sendmsguser)
			UserMsg = Replace(UserMsg,"{$newmsgnum}",Dvbbs.sendmsgnum)
			template.html(1) = Replace(template.html(1),"{$umsg}",UserMsg)
		Else
			template.html(1) = Replace(template.html(1),"{$umsg}",TopArray(3))
		End If
		If Dvbbs.Forum_ChanSetting(0)="1" Then template.html(1)=Replace(template.html(1),"{$sysmsg}",Replace(TempArray(0),"{$raypic}",Dvbbs.mainpic(14)))
		template.html(1)=Replace(template.html(1),"{$sysmsg}","")
		TempStr=Replace(TempStr,"{$myinfo}",template.html(1))
		TempStr=Replace(TempStr,"{$UserID}",Dvbbs.Userid)
		If IsNumeric(Dvbbs.MyUserInfo(12)) And IsNumeric(Dvbbs.MyUserInfo(13)) And Dvbbs.MyUserInfo(13)<>"" And Dvbbs.MyUserInfo(12)<>"" Then
			If Clng(Dvbbs.MyUserInfo(13))=Clng(Dvbbs.Forum_Setting(39)) And Clng(Dvbbs.MyUserInfo(12))=Clng(Dvbbs.Forum_Setting(38)) Then
			TempStr=Replace(TempStr,"{$userlogo}","<img src="&Dvbbs.MyUserInfo(11)&">")
			Else
			TempStr=Replace(TempStr,"{$userlogo}","<img src="&Dvbbs.MyUserInfo(11)&" width=60 height=60>")
			End If
		Else
			TempStr=Replace(TempStr,"{$userlogo}","<img src=images/logo_2.gif>")
		End If
	End If
	TempStr=Replace(TempStr,"{$bgcolor}",Dvbbs.mainsetting(12))
	TempStr=Replace(TempStr,"{$alertcolor}",Dvbbs.mainsetting(1))
	Response.Write TempStr
End Sub 

Function news()
	Dvbbs.Name="news"&Dvbbs.boardid
	If Dvbbs.ObjIsEmpty() Then 
		Dim tmpstr,bgs
		Dim Rs,SQL
		SQL="select top 1 title,addtime,bgs from Dv_bbsnews where boardid="&Dvbbs.boardid&" order by id desc"
		Set Rs=DVbbs.Execute(sql)
		If Rs.BOF And Rs. EOF Then
			tmpstr=template.Strings(8)&"|||"
		Else
			bgs=Rs(2)
			If bgs="" or isnull(bgs) then
				tmpstr=Rs(0)&"|||"&Rs(1)
			Else
				tmpstr="<img src=Skins/Default/filetype/mid.gif border=0><bgsound src="&bgs&" border=0>"&Rs(0)&"|||"&Rs(1)
			End if
		End If
		Set Rs=Nothing 
		Dvbbs.Value=tmpstr		 
	End If
	news=split(Dvbbs.Value,"|||")
End Function 
Sub GetsmBbsList()
	Dim Chachedata,ishidden,ShowMasters,j
	Dim Forum_Boards,i,BoardID,Board_Data,ClassID
	Dim setings,lastposttime,depth,lastpost,BoardType,BoardReadme,htmlstr
	Dvbbs.name="BbsListTop"&Dvbbs.skinid
	If Dvbbs.ObjIsEmpty() Then	
		Dim TempListArray,havenew,loadboard
		TempListArray = Split(template.html(8),"||")
		Chachedata= Chachedata& Replace(Replace(template.html(7),"{$follow}",Dvbbs.mainpic(11)),"{$nofollow}",Dvbbs.mainpic(10))
		Chachedata= Chachedata& "<script language=""javascript"">"
		Chachedata= Chachedata& vbNewLine
		'传送图片变量到JS
		For i=0 to UBound(template.pic)-1
			Chachedata= Chachedata& "piclist["&i&"]='"&template.pic(i)&"';"
			Chachedata= Chachedata& vbNewLine		
		Next
		'传递论坛主设置数据到JS
		For i=0 to UBound(Dvbbs.mainsetting)
			Chachedata= Chachedata& "mainsetting["&i&"]='"&Dvbbs.mainsetting(i)&"';"
			Chachedata= Chachedata& vbNewLine	
		Next 
		'传送模板数据到JS以备调用
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(4),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(5),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(1),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(6),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(10),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		'传送字符串变量到JS
		For i=0 to 10
			Chachedata= Chachedata& "Strings[Strings.length]='"& template.Strings(i)&"';"		
		Next
	
		Dvbbs.value = Chachedata
	End If
	Response.Write Dvbbs.value
	Response.Write "</script>"
	template.html(8)=Split(template.html(8),"||")
	ClassID=""
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 to UBound(Forum_Boards)
		Dvbbs.Name="BoardInfo_" & Forum_Boards(i)
  		If Dvbbs.ObjIsEmpty() Then Dvbbs.ReloadBoardInfo(Forum_Boards(i))
		Board_Data=Dvbbs.Value
		If Board_Data(2,0)="0" Then
			BoardType=Board_Data(1,0)&""
			If ClassID<>"" Then 
				Response.Write template.html(8)(1)		
				Response.Write "<br>"
			End If
			ClassID=Forum_Boards(i)
			htmlstr=template.html(14)
			htmlstr=Replace(htmlstr,"{$boardid}",Board_Data(0,0))
			htmlstr=Replace(htmlstr,"{$pic}","")
			htmlstr=Replace(htmlstr,"{$BoardType}",BoardType)
			Response.Write 	htmlstr
			Response.Write template.html(8)(0)		
		Else
			havenew=0
			loadboard=True
			ishidden=false
			depth=CInt(Board_Data(4,0))
			If depth > Cint(Dvbbs.forum_setting(5)) Then
			Else
				ShowMasters=""			
				Board_Data(8,0)=split(Board_Data(8,0)&"","|")
				For j=0 to UBound(Board_Data(8,0))
					If j>5 Then 
						ShowMasters=ShowMasters&"<font color=gray>More</font>"
						Exit For
					End If
					ShowMasters=ShowMasters&"&nbsp;<a href=dispuser.asp?name="&Board_Data(8,0)(j)&" target=_blank >"&Board_Data(8,0)(j)&"</a>"
				Next 
				If ShowMasters="" Then ShowMasters="&nbsp;暂无"
				setings=split(Board_Data(16,0),",")
				lastpost=Board_Data(14,0)
				lastposttime=split(Board_Data(14,0),"$")(2)
				If Not IsDate(lastposttime) Then lastposttime=Now()
				If datediff("h",Dvbbs.Lastlogin,lastposttime)=0 Then havenew=1
				If CInt(setings(1))=1 And Dvbbs.GroupSetting(37)<>"1" Then loadboard=False
				If loadboard  Then
					BoardReadme=Board_Data(7,0)&""
					BoardType=Board_Data(1,0)&""
					htmlstr= template.html(5)
					htmlstr=Replace(htmlstr,"{$boardid}",Board_Data(0,0))
					htmlstr=Replace(htmlstr,"{$readme}",BoardReadme)
					htmlstr=Replace(htmlstr,"{$BoardType}",BoardType)
					If Board_Data(6,0)="0" Then 
						Board_Data(6,0)="" 
					Else
						Board_Data(6,0)=Replace(template.Strings(1),"{$child}",Board_Data(6,0))
					End If
					If Trim(Board_Data(11,0))<>"" Then
						Board_Data(11,0)="<table align=""left""><tr><td><a href=""list.asp?boardid="&Board_Data(0,0)&"""><img src="""&Board_Data(11,0)&""" align=""top"" border=""0""></a></td><td width=""20""></td></tr></table>"
					End If
					htmlstr=Replace(htmlstr,"{$indexIMG}",Board_Data(11,0)&"")
					htmlstr=Replace(htmlstr,"{$child}",Board_Data(6,0))
					htmlstr=Replace(htmlstr,"{$alertcolor}",Dvbbs.mainsetting(1))
					htmlstr=Replace(htmlstr,"{$blinkcolor}",Dvbbs.mainsetting(3))
					htmlstr=Replace(htmlstr,"{$PostNum}",Board_Data(9,0))
					htmlstr=Replace(htmlstr,"{$TopicNum}",Board_Data(10,0))
					htmlstr=Replace(htmlstr,"{$todayNum}",Board_Data(12,0))
					If setings(2)="1" And Not Dvbbs.Master Then
						 htmlstr=Replace(htmlstr,"{$LastPost}",template.Strings(2))
					Else
						htmlstr=Replace(htmlstr,"{$LastPost}",showlastpost(lastpost))
					End If 
					htmlstr=Replace(htmlstr,"{$statuspic}",showpic(havenew,setings(0),setings(2)))
					htmlstr=Replace(htmlstr,"{$ShowMasters}",ShowMasters)
					Response.Write htmlstr		
				End If
			End If
		End If	
		
	Next
	Response.Write template.html(8)(1)
	Response.Write  "<br>"
End Sub
Function showpic(havenew,Board_Setting,Board_Setting1)
	Dim pic,Str,Str1
	Str="无新贴"
	Str1="开放的版面"
	pic=template.pic(0)
	If havenew=1 Then
		Str="有新贴"
		pic=template.pic(1)
	End If 
	If Board_Setting =1 Then 
		pic=template.pic(2)
		Str1="锁定的版面"
	End If 
	If Board_Setting1=1 Then
		pic=template.pic(2)
		Str1="认证论坛"
	End If
	showpic="<img src="""&pic&""" alt="""&Str1&","&Str&""">"
End Function 
Function showlastpost(lastpoststr)
	lastpoststr=replace(lastpoststr,"<","&lt;")
	if lastpoststr="$$$$" Or lastpoststr="" Then 
		showlastpost="主题：无<br>作者：无<br>日期：无"
	Else
		 Dim Str
		lastpoststr=split(lastpoststr,"$")
		Str=Str&"主题：<a href=""Dispbbs.asp?boardid="&lastpoststr(7)&"&ID="&lastpoststr(6)&"&replyID="&lastpoststr(1)&"&skin=1"" title=""转到："&lastpoststr(3)&""">"
		Str=Str&Left(lastpoststr(3),10)
		Str=Str&"</a>"
		Str=Str&"<br>作者："
		Str=Str&"<a href=""dispuser.asp?id="&lastpoststr(5)&""" target=""_blank"">"&lastpoststr(0)&"</a>"
		Str=Str&"<br>日期："
		Str=Str&lastpoststr(2)&"&nbsp;<a href=""dispbbs.asp?Boardid="&lastpoststr(7)&"&ID="& lastpoststr(6) &"&replyID="& lastpoststr(1) &"&skin=1""><IMG border=0 src=""Skins/Default/lastpost.gif"" title=""主题："&lastpoststr(3)&"""></a>"
		showlastpost=Str
	End If
End Function
Sub GetBbsList()
	Dim Chachedata,ishidden
	Dvbbs.name="BbsListTop"&Dvbbs.skinid'长时间的缓存数据
	If Dvbbs.ObjIsEmpty() Then	
		Dim TempListArray,havenew,loadboard
		TempListArray = Split(template.html(8),"||")
		Chachedata= Chachedata& Replace(Replace(template.html(7),"{$follow}",Dvbbs.mainpic(11)),"{$nofollow}",Dvbbs.mainpic(10))
		Chachedata= Chachedata& "<script language=""javascript"">"
		Chachedata= Chachedata& vbNewLine
		'传送图片变量到JS
		For i=0 to UBound(template.pic)-1
			Chachedata= Chachedata& "piclist["&i&"]='"&template.pic(i)&"';"
			Chachedata= Chachedata& vbNewLine		
		Next
		'传递论坛主设置数据到JS
		For i=0 to UBound(Dvbbs.mainsetting)
			Chachedata= Chachedata& "mainsetting["&i&"]='"&Dvbbs.mainsetting(i)&"';"
			Chachedata= Chachedata& vbNewLine	
		Next 
		'传送模板数据到JS以备调用
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(4),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(5),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(1),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(6),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		Chachedata= Chachedata& "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(10),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
		Chachedata= Chachedata& vbNewLine
		'传送字符串变量到JS
		For i=0 to 10
			Chachedata= Chachedata& "Strings[Strings.length]='"& template.Strings(i)&"';"		
		Next
	
		Dvbbs.value = Chachedata
	End If
	Response.Write Dvbbs.value
	Dim Forum_Boards,i,BoardID,Board_Data,ClassID
	Dim setings,lastposttime,depth,lastpost,BoardType,BoardReadme
	ClassID=""
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 to UBound(Forum_Boards)
		Dvbbs.Name="BoardInfo_" & Forum_Boards(i)
  		If Dvbbs.ObjIsEmpty() Then Dvbbs.ReloadBoardInfo(Forum_Boards(i))
		Board_Data=Dvbbs.Value
		If Board_Data(2,0)="0" Then
			BoardType=Board_Data(1,0)&""
			BoardType=Replace(BoardType,"\","\\")
			BoardType=Replace(BoardType,"'","\'")
			If ClassID<>"" Then 
				Response.Write "classfooter();"		
			End If
			ClassID=Forum_Boards(i)
			Response.Write "showclass("
			Response.Write Forum_Boards(i)
			Response.Write ",'"
			Response.Write BoardType
			Response.Write "','"
			Response.Write Board_Data(16,0)
			Response.Write "','"
			Response.Write Request.Cookies("List")("list"&Forum_Boards(i))
			Response.Write "',"
			Response.Write Board_Data(6,0)
			Response.Write ");"		
		Else
			havenew=0
			loadboard=True
			ishidden=false
			depth=CInt(Board_Data(4,0))
			If depth > Cint(Dvbbs.forum_setting(5)) Then
			Else
				setings=split(Board_Data(16,0),",")(1)
				lastpost=Board_Data(14,0)
				lastpost=Server.HTMLEnCode(lastpost)
				lastpost=Replace(lastpost,"\","\\")
				lastpost=Replace(lastpost,CHR(10) & CHR(10),"\r")
				lastpost=Replace(lastpost,Chr(10),"\n")
				lastpost=Replace(lastpost,Chr(13),"")
				lastposttime=split(Board_Data(14,0),"$")(2)
				If Not IsDate(lastposttime) Then lastposttime=Now()
				If datediff("h",Dvbbs.Lastlogin,lastposttime)=0 Then havenew=1
				If CInt(setings)=1 And Dvbbs.GroupSetting(37)<>"1" Then loadboard=False
				If loadboard  Then
					BoardReadme=Board_Data(7,0)&""
					BoardType=Board_Data(1,0)&""
					BoardType=Replace(BoardType,"\","\\")
					BoardType=Replace(BoardType,"'","\'")
					Response.Write "showboard("
					Response.Write Forum_Boards(i)
					Response.Write ",'"
					Response.Write BoardType
					Response.Write "',"
					Response.Write Board_Data(6,0)
					Response.Write ",'"
					Response.Write BoardReadme
					Response.Write "','"
					Response.Write Board_Data(8,0)
					Response.Write "',"
					Response.Write Board_Data(9,0)
					Response.Write ","
					Response.Write Board_Data(10,0)
					Response.Write ",'"
					Response.Write Board_Data(11,0)
					Response.Write "',"
					Response.Write Board_Data(12,0)
					Response.Write ",'"		
					Response.Write lastpost
					Response.Write "','"
					Response.Write Board_Data(16,0)
					Response.Write "',"
					Response.Write havenew
					Response.Write ");"
				Else	
					Response.Write "Child=(Child-1);"	
					Response.Write "boardcount++;"			
					Response.Write "showcode('','');"
				End If
			End If
		End If	
		Response.Write  vbNewLine
	Next
	If ClassID<>"" Then 
		Response.Write "classfooter();"		
	End If
	Response.Write vbNewLine
	Response.Write "</script>"
	Forum_Boards = Null
End Sub

Function Getlink()
	Dvbbs.Name="link"
	If Dvbbs.ObjIsEmpty() Then
		Dim Rs,SQl
		SQL="select boardname,readme,url,logo,islogo from [Dv_bbslink] Order by islogo,id"
		Set Rs=Dvbbs.Execute(SQL)
		If Not rs.eof Then
			Dvbbs.Value=RS.GetString (,,"!@#%|","$?&!@","")
		Else
			Dvbbs.Value=""
		End If
	End If
	Getlink=Dvbbs.Value
End Function

Sub Show_Index_Footer()
	TempStr = template.html(11)
	TempStr = Replace(TempStr,"{$userip}",Dvbbs.UserTrueIP)
	TempStr = Replace(TempStr,"{$system}",BrowserType.platform)
	TempStr = Replace(TempStr,"{$brw}",BrowserType.Browser & BrowserType.version)
	TempStr = Replace(TempStr,"{$showstr}",template.Strings(6))
	TempStr = Replace(TempStr,"{$onlinenum}",MyBoardOnline.Forum_Online)
	TempStr = Replace(TempStr,"{$ousernum}",MyBoardOnline.Forum_UserOnline)
	TempStr = Replace(TempStr,"{$gusernum}",MyBoardOnline.Forum_GuestOnline)
	TempStr = Replace(TempStr,"{$maxuser}",Dvbbs.Maxonline)
	TempStr = Replace(TempStr,"{$maxusertime}",Dvbbs.CacheData(6,0))
	TempStr = Replace(TempStr,"{$piclist}",GetGroupTitle())
	TempStr = Replace(TempStr,"{$BuildDate}",FormatDateTime(Dvbbs.Forum_Setting(74),1))
	TempStr = Replace(TempStr,"{$nonewpic}",template.pic(0))
	TempStr = Replace(TempStr,"{$isnewpic}",template.pic(1))
	TempStr = Replace(TempStr,"{$islockpic}",template.pic(2))
	Response.Write TempStr
	If Dvbbs.forum_setting(14)="1" Or Dvbbs.forum_setting(15)="1" Then 
		Response.Write "<iframe width=""0"" height=""0"" src=""Online.asp?action=1&Boardid=0"" name=""hiddenframe""></iframe>"
	Else
		Response.Write "<iframe width=""0"" height=""0"" src="""" name=""hiddenframe""></iframe>"
	End If
	TempStr = ""
	Response.Write "<script language=""javascript"">"
	Response.Write Chr(10)
	
	If Dvbbs.Forum_ads(2)="1" Then
		 Response.Write "move_ad('"&Dvbbs.Forum_ads(3)&"','"&Dvbbs.Forum_ads(4)&"','"&Dvbbs.Forum_ads(5)&"','"&Dvbbs.Forum_ads(6)&"');"
	End If
	If Dvbbs.Forum_ads(13)="1" Then
		Response.Write "fix_up_ad('"& Dvbbs.Forum_ads(8) & "','" & Dvbbs.Forum_ads(10) & "','" & Dvbbs.Forum_ads(11) & "','" & Dvbbs.Forum_ads(9) & "');"		
	End If 
	Response.Write Chr(10)
	Response.Write "</script>"
End Sub

Function GetGroupTitle()
	Dvbbs.Name="GroupTitle"
	If Dvbbs.ObjIsEmpty() Then
		Dim Rs,SQl
		SQL="select TitlePic,title from [Dv_UserGroups] where IsDisp=1 Order by Orders "
		Set Rs=Dvbbs.Execute(SQL)
		SQL="<img src="""&RS.GetString (,,"""> ","    ‖ <img src=""","")
		SQl=Left(SQL,Len(SQL)-Len("    ‖ <img src="""))
		If Dvbbs.Forum_ChanSetting(0)="1" Then 
			SQl= SQL & "    ‖ <img src="""&Dvbbs.mainpic(14)&""">  "&Dvbbs.lanStr(6)
		End If
		Dvbbs.Value = SQL
		Set rs=Nothing 
	End If
	GetGroupTitle=Dvbbs.Value
End Function

Sub birthuser()
	Dim Strings
	Strings=Dvbbs.CacheData(16,0)
	Strings=split(Strings,"$$")
	If Not IsDate(Strings(0)) Then Strings(0)=Now()-1
	If CDate(Strings(0)) <> Date() Then 
		Dim Rs,SQL,NowMonth,NowDate,TMPDATA,birthNum,tmpstr,i,todaystr0,todaystr1
		NowMonth=Month(Date())
		NowDate=Day(Date())
		If NowMonth< 10 Then
			todaystr0="0"&NowMonth
		Else
			todaystr0=CStr(NowMonth)
		End If
		If NowDate < 10 Then
			todaystr0=todaystr0&"-"&"0"&NowDate
		Else
			todaystr0=todaystr0&"-"&NowDate
		End If
		todaystr1=NowMonth&"-"&NowDate
		If todaystr0=todaystr1 Then
			SQL="select username,Userbirthday from [Dv_user] where Userbirthday like '%"&todaystr1&"' Order by UserID"
		Else
			SQL="select username,Userbirthday from [Dv_user] where Userbirthday like '%"&todaystr1&"' Or Userbirthday like '%"&todaystr0&"' Order by UserID"
		End If
		birthNum=0
		Set Rs=Dvbbs.Execute(SQL)
		i=0
		If Not Rs.EOF Then
			Do while Not Rs.EOF
				If IsDate(Rs(1)) Then 
					If Month(Rs(1))=NowMonth And Day(Rs(1)) Then
						i=i+1
						tmpstr=template.Strings(10)
						birthNum=birthNum+1
						tmpstr=Replace(tmpstr,"{$username}",rs(0))
						tmpstr=Replace(tmpstr,"{$age}",datediff("yyyy",rs(1),Now()))
						If i=1  Then
							TMPDATA=TMPDATA&"<tr>"
						End If
						TMPDATA=TMPDATA&"<td>"&tmpstr&"</td>"
						If i=5 Then
							TMPDATA=TMPDATA&"</tr>"
							i=0
						End If
					End If
				End If
				Rs.MoveNext
			Loop
		End If
		If birthNum mod 5 <> 0 Then TMPDATA=TMPDATA&"</tr>"
		TMPDATA="<TABLE cellSpacing=2 cellPadding=2 width=100% border=0>"&TMPDATA&"</table>"
		Set Rs=Nothing
		template.html(12)=Replace(template.html(12),"{$birthNum}",birthNum)
		If  TMPDATA="" Then 
			TMPDATA=template.Strings(9)
		End If
		template.html(12)=Replace(template.html(12),"{$birthday}",TMPDATA)
		TMPDATA=Date()&"$$"&template.html(12)
		Dvbbs.Execute("Update Dv_setup Set Forum_BirthUser='"&TMPDATA&"'")
		Dvbbs.ReloadSetupCache TMPDATA,16
		
	End If
	Strings=Split(Dvbbs.CacheData(16,0),"$$")
	Strings(1)=Replace(Strings(1),"{$bpic}",template.pic(3))
	Response.Write Strings(1)
End Sub

%>