<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<!--#include file="inc/ubblist.asp"-->
<%
If Dvbbs.BoardID = 0 Then
	Response.Write "参数错误"
	Response.End 
End If
Dvbbs.LoadTemplates("dispbbs")
Dim AnnounceID,ReplyID,Star,Skin,followup
Dim CanReply,IsTop,IsVote,TopicCount,PollID,TotaluseTable,ViewNum,Topic,TopicMode
Dim PostBuyUser,abgcolor,bgcolor,UserName
Chk_Topic_Err
Dvbbs.Showerr()
Dvbbs.Nav()
Dvbbs.Showerr()
Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
Dvbbs.ActiveOnline()
Dim Page,LockTopic
Dim action
Dim TopicNum,n,SplitPageNum
Dim EmotPath
EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)		'em心情路径
action=Request("action")
Page=Request("Page")
If isNumeric(Page) = 0 or Page="" Then Page=1
Page=Clng(Page)
Show_Topic_Top()
Dvbbs.ShowErr()
Dim dv_ubb
Set dv_ubb=new Dvbbs_UbbCode
Show_Topic_Body
Set dv_ubb=Nothing
If Skin=1 Then showtree()
If CanReply Then Show_Topic_FastRe
If Dvbbs.UserID>0 Then Show_Topic_ManageAction
Dvbbs.NewPassword()
Dvbbs.Footer()

Function Chk_Topic_Err
	AnnounceID=Request("ID")
	If AnnounceID="" Or Not IsNumeric(AnnounceID) Then Dvbbs.AddErrCode(30)
	ReplyID=Request("ReplyID")
	If ReplyID="" Or Not IsNumeric(ReplyID) Then ReplyID=AnnounceID
	Star=Request("Star")
	If Star="" Or Not IsNumeric(Star) Then Star=1
	Star=Clng(Star)
	Skin=Request("Skin")
	If Skin="" Or Not IsNumeric(Skin) Then Skin=Dvbbs.Board_setting(24)
	If Dvbbs.ErrCodes<>"" Then Exit Function
	Dim BrowserType
	Set BrowserType=New Cls_Browser
	If BrowserType.IsSearch Then Response.redirect "printpage.asp?BoardID="&Dvbbs.BoardID&"&ID="&AnnounceID
	Set BrowserType=Nothing
	Dim SQl,Rs
	Dim MyCanReply
	Set Rs=Server.CreateObject("Adodb.RecordSet")
	SQL="Select title,istop,isbest,PostUserName,PostUserid,hits,isvote,child,pollid,LockTopic,PostTable,BoardID,TopicMode from dv_topic where topicID="&Announceid
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open SQL,Conn,1,3
	Dvbbs.SqlQueryNum=Dvbbs.SqlQueryNum+1
	'Set Rs=Dvbbs.Execute(SQL)
	If Not(Rs.BOF and Rs.EOF) then
		If Rs(11)<>Dvbbs.BoardID Then Dvbbs.AddErrCode(29)
		Rs(5)=Rs(5)+1
		Rs.Update
		Topic=Rs(0)
		istop=rs(1)
		isVote=rs(6)
		TopicCount=rs(7)+1
		pollid=rs(8)
		Locktopic=rs(9)
		TotalUseTable=rs(10)
		TopicMode=rs(12)
		ViewNum=Rs(5)
		If rs(3)=Dvbbs.Membername then
			MyCanReply=Dvbbs.GroupSetting(4)
		Else
			MyCanReply=Dvbbs.GroupSetting(5)
			If Cint(Dvbbs.GroupSetting(2))=0 Then Dvbbs.AddErrcode(31)
		End If
		If Len(Topic) > Cint(Dvbbs.Board_Setting(25)) And Not TopicMode>0 Then
			Topic=Left(Topic,Dvbbs.Board_Setting(25))&"..."
		End If
		If TopicMode>0 Then
			If TopicMode=1 Then
				Topic = Replace(Topic,"<!--","&lt;!--")		
			Else
				Topic = Replace(Topic,"<","&lt;")
				Topic = Replace(Topic,">","&gt;")	
				Topic=Dvbbs_TopicMode(Topic,TopicMode)
			End If
		Else
			Topic = Replace(Topic,"<","&lt;")
			Topic = Replace(Topic,">","&gt;")
		End If
		Topic=Dvbbs.ChkBadWords(Topic)
		Dvbbs.Stats=Topic
	Else
		Dvbbs.AddErrcode(32)
	End If
	Rs.Close
	Set Rs=Nothing
	CanReply=False
	If (Not Dvbbs.Board_Setting(0)="1" And Cint(mycanreply)=1 And Cint(locktopic)=0) Or (Dvbbs.master Or Dvbbs.superboardmaster Or Dvbbs.boardmaster) Then
		CanReply=True
	End If
End Function

Function Show_Topic_Top()
	 
	Dim TempStr,MyTempStr
	If (Dvbbs.Board_Setting(43)="0" And Dvbbs.Board_Setting(0)="0") Or (Dvbbs.Board_Setting(43)="0" And Dvbbs.Board_Setting(0)="1" And (Dvbbs.Master Or Dvbbs.SuperBoardMaster Or Dvbbs.BoardMaster)) Then
		MyTempStr=Split(template.html(1),"||")
		TempStr=Replace(MyTempStr(0),"{$pic_newpost}",Dvbbs.mainpic(7))
		TempStr=Replace(TempStr,"{$pic_newvote}",Dvbbs.mainpic(8))
		If Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(7)=1 Then
			TempStr=TempStr & MyTempStr(1)
		End If
	Else
		If Dvbbs.Board_Setting(0)="1" Then TempStr=template.Strings(0)
	End If
	TempStr=Replace(template.html(0),"{$topicpostinfo}",TempStr)
	TempStr=Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	TempStr=Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
	TempStr=Replace(TempStr,"{$page}",page)
	TempStr=Replace(TempStr,"{$replyid}",ReplyID)
	TempStr=Replace(TempStr,"{$star}",Star)
	TempStr=Replace(TempStr,"{$announceid}",AnnounceID)
	TempStr=Replace(TempStr,"{$viewnum}",ViewNum)
	Dim Skinpic,Skinname,nskin
	If Skin="1" Then
		nskin=0
		Skinpic=template.pic(1)
		Skinname=template.Strings(2)
	Else
		nskin=1
		Skinpic=template.pic(0)
		Skinname=template.Strings(1)
	End If
	TempStr=Replace(TempStr,"{$skin}",nskin)
	TempStr=Replace(TempStr,"{$skinname}",skinname)
	TempStr=Replace(TempStr,"{$skinpic}",skinpic)
	TempStr=Replace(TempStr,"{$topic}",Topic)
	If IsVote=1 Then
		TempStr=Replace(TempStr,"{$voteinfo}",Show_Topic_Vote)
	Else
		TempStr=Replace(TempStr,"{$voteinfo}","")
	End If
	Response.Write TempStr
End Function

Function Show_Topic_Body()
	If UBound(Dvbbs.Forum_ads)>13 Then 
		Dvbbs.Forum_ads(14)=Split(Dvbbs.Forum_ads(14),vbNewLine)
	End If
	Dim SQL,Rs,i
	Dim TopicPageList,Pcount
	Dim layer
	TopicPageList=Dvbbs.Board_Setting(27)
	With Response
	.Write "<Script Language=JavaScript>"
	.Write  template.html(4)
	.Write  "var TempStr='"&Replace(Replace(Replace(Replace(Replace(template.html(2),"{$boardtype}",Server.Htmlencode(Dvbbs.Board_Data(1,0))),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write  "var sTempStr='"&Replace(Replace(Replace(Replace(template.html(3),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write  "sTempStr=sTempStr.split(""||"");"
	.Write  "var alertcolor='"&Dvbbs.mainsetting(1)&"';"
	.Write  "var Board_Setting27='"&TopicPageList&"';"
	.Write  "var fontsize='"&Dvbbs.Board_Setting(28)&"';"
	.Write  "var lineheight='"&Dvbbs.Board_Setting(29)&"';"
	.Write  "var Forum_Setting42='"&Dvbbs.Forum_Setting(42)&"';"
	.Write  "var facesetting='"&Dvbbs.Forum_Setting(53)&"';"
	.Write  "var votemoney='"&Dvbbs.GroupSetting(47)&"';"
	.Write  "var Forum_ChanSetting0='"&Dvbbs.Forum_ChanSetting(0)&"';"
	.Write  "var Forum_ChanSetting5='"&Dvbbs.Forum_ChanSetting(5)&"';"
	.Write  "var Forum_ChanSetting6='"&Dvbbs.Forum_ChanSetting(6)&"';"
	.Write  "var Forum_ChanSetting7='"&Dvbbs.Forum_ChanSetting(7)&"';"
	.Write  "var topfloor='"&template.Strings(3)&"';"
	.Write  "var floor='"&template.Strings(4)&"';"
	.Write  "var lockuserinfo1='"&template.Strings(5)&"';"
	.Write  "var lockuserinfo2='"&template.Strings(7)&"';"
	.Write  "var noviewbest='"&template.Strings(6)&"';"
	.Write  "var actioninfo1='"&template.Strings(8)&"';"
	.Write  "var actioninfo2='"&template.Strings(9)&"';"
	.Write  "var GroupSetting41='"&Dvbbs.GroupSetting(41)&"';"
	.Write  "var TopicMode='"&TopicMode&"';"
	.Write  "var mainsetting='"&Dvbbs.mainhtml(0)&"';"
	.Write  "var mainsetting=mainsetting.split(""||"");"
	.Write  "var TopicNum='"&TopicCount&"';"
	.Write  "</Script>"
	SQL="B.AnnounceID,B.BoardID,B.UserName,B.Topic,B.dateandtime,B.body,B.Expression,B.ip,B.RootID,B.signflag,B.isbest,B.PostUserid,B.layer,b.isagree,U.useremail,U.UserIM,U.UserMobile,U.Usersign,U.userclass,U.Usertitle,U.Userwidth,U.Userheight,U.UserPost,U.Userface,U.JoinDate,U.userWealth,U.userEP,U.userCP,U.Userbirthday,U.Usersex,u.UserGroup,u.LockUser,u.userPower,U.titlepic,U.UserGroupID,U.LastLogin,B.PostBuyUser,U.UserHidden,U.IsChallenge,B.Ubblist,B.LockTopic"
	If cint(skin)=1 and Clng(replyid)=Clng(Announceid) Then
		SQL="Select top 1 "&SQL&" From "&TotalUseTable&" B Inner Join [dv_user] U On U.UserID=B.PostUserID Where B.BoardID="&Dvbbs.BoardID&" And B.RootID="&AnnounceID&"  Order By B.AnnounceID, B.DateAndTime"
	ElseIf cint(skin)=1 Then
		SQL="Select "&SQL&" From "&TotalUseTable&" B Inner Join [dv_user] U On U.UserID=B.PostUserID Where B.BoardID="&Dvbbs.BoardID&" And B.AnnounceID="&replyID
	Else
		Dim AnnounceIDlists
		AnnounceIDlists=AnnounceIDlist()
		SQL="Select "&SQL&" From "&TotalUseTable&" B Inner Join [dv_user] U On U.UserID=B.PostUserID Where B.RootID="&Announceid&" And B.BoardID="&Dvbbs.BoardID&" And  B.AnnounceID in ("&AnnounceIDlists&") Order BY B.AnnounceID, B.DateAndTime"
	End If
	Set Rs =Dvbbs.Execute(SQL)
	If Rs.EOF And Rs.BOF Then
		Dvbbs.AddErrCode(33)
		Exit Function
	End If

	If Not(Rs.EOF And Rs.BOF) Then
		followup = Rs("AnnounceID")
		If TopicCount mod Cint(TopicPageList)=0 then
			Pcount= TopicCount \ Cint(TopicPageList)
		Else
			Pcount= TopicCount \ Cint(TopicPageList)+1
		End If
		'Rs.MoveFirst
		If star > Pcount Then star = Pcount
		If star < 1 Then star = 1				
		'If Cint(skin) <> 1 Then Rs.Move (star-1) * TopicPageList
		.Write "<Script Language=JavaScript>"
		.Write  "var star='"&star&"';"
		.Write  "</Script>"
		'AnnounceID=0,BoardID=1,UserName=2,Topic=3,dateandtime=4,body=5,
		'Expression=6,ip=7,RootID=8,signflag=9,isbest=10,PostUserid=11,
		'layer=12,isagree=13,useremail=14,UserIM=15,UserMobile=16,sign=17,
		'userclass=18,title=19,width=20,height=21,article=22,face=23,JoinDate=24,
		'userWealth=25,userEP=26,userCP=27,birthday=28,sex=29,UserGroup=30,LockUser=31,
		'userPower=32,titlepic=33,UserGroupID=34,LastLogin=35,PostBuyUser=36,Ubblists=39,LockTopic=40
		Dim k,rndnum,TopicAddata,Topic_Ads,UserSign,TopicBody
		K=0
		Dim MyTempStr,ReplaceStr
		SQL=Rs.GetRows(TopicPageList)
		Set Rs=Nothing
		For i=0 To Ubound(SQL,2)
			.Write "<a name="&SQL(0,i)&"></a>"
			.Write "<Script Language=JavaScript>"
			UserName=Dvbbs.HtmlEncode(SQL(2,i))
			If SQL(40,i)=2 Then
				.Write "var actioninfo3='"&template.Strings(16)&"';"
			Else
				.Write "var actioninfo3='"&template.Strings(15)&"';"
			End If
			If bgcolor="tablebody1" Then
				bgcolor="tablebody2"
				abgcolor="tablebody1"
			Else
				bgcolor="tablebody1"
				abgcolor="tablebody2"
			End If
			ReplyID=SQL(0,i)
			PostBuyUser=SQL(36,i)
			Topic_Ads=""
			If Dvbbs.Forum_ChanSetting(5)="1" Then
				If Dvbbs.Forum_AdLoop3<>"" Then
					Randomize
					rndnum=Cint((i3-1)*rnd+1)
					If rndnum=0 Then rndnum=1
					TopicAddata=Ad_3(rndnum-1)
					TopicAddata=Replace(Replace(TopicAddata,"\","\\"),"'","\'")
					Topic_Ads=TopicAddata
					TopicAddata=""
				Else
					Topic_Ads=""
				End If
			Else
				If UBound(Dvbbs.Forum_ads)>13 Then
					If Topic_Ads="" And UBound(Dvbbs.Forum_ads(14)) > -1 Then
						Randomize
						Topic_Ads=Dvbbs.Forum_ads(14)(CInt(UBound(Dvbbs.Forum_ads(14))*Rnd))
						Topic_Ads= Replace(Replace(Topic_Ads,"\","\\"),"'","\'")
					End If
				Else
					Topic_Ads=""
				End If
			End If	
			UserSign=""
			If Not Isnull(SQL(17,i)) Or Not SQL(17,i)="" Then
				If SQL(9,i)=1 and SQL(31,i)=0 and Cint(Dvbbs.forum_setting(42))=1 Then
					UserSign = dv_ubb.Dv_SignUbbCode(SQL(17,i),SQL(34,i))
					UserSign=Replace(Replace(UserSign,"\","\\"),"'","\'")
					'UserSign = Replace(UserSign, vbNewLine,"\n")
				End If
			End If
			Ubblists=SQL(39,i)
			If Not (SQL(31,i)=2 Or (SQL(10,i)=1 And CInt(Dvbbs.GroupSetting(41))=0) Or SQL(31,i)=1) Then
				If InStr(Ubblists,",39,") > 0  Then
					TopicBody = dv_ubb.Dv_UbbCode(SQL(5,i),SQL(34,i),1,0)
				Else
					TopicBody = dv_ubb.Dv_UbbCode(SQL(5,i),SQL(34,i),1,1)
				End If			
			End If
			TopicBody = Replace(Replace(TopicBody ,"\","\\"),"'","\'")
			'TopicBody = Replace(TopicBody, vbNewLine,"\n")
			SQL(13,i)=Replace(Lcase(SQL(13,i))&"","[isubb]","")
			.Write "document.write (dvbbs_show_topic('"
			.Write SQL(0,i)
			.Write "','"
			.Write Dvbbs.BoardID
			.Write "',"
			MyTempStr	= "'"&SQL(2,i)&"','"
			ReplaceStr	= SQL(3,i)&""
			If Not (i=0 And Star=1 And TopicMode=1 ) Then
				ReplaceStr = Replace(ReplaceStr&"","<","&lt;")
				ReplaceStr = Replace(ReplaceStr,">","&gt;")
			End If
			ReplaceStr	=	Replace(Replace(ReplaceStr ,"\","\\"),"'","\'")
			MyTempStr	=	MyTempStr & ReplaceStr
			ReplaceStr	=	SQL(14,i)&""
			ReplaceStr	=	Replace(Replace(ReplaceStr ,"\","\\"),"'","\'")
			SQL(14,i)	=	ReplaceStr
			ReplaceStr	=	SQL(15,i)&""
			ReplaceStr	=	Replace(Replace(ReplaceStr ,"\","\\"),"'","\'")
			SQL(15,i)	=	ReplaceStr
			ReplaceStr	=	SQL(16,i)&""
			ReplaceStr	=	Replace(Replace(ReplaceStr ,"\","\\"),"'","\'")
			SQL(16,i)	=	ReplaceStr
			ReplaceStr	=	SQL(23,i)&""
			ReplaceStr	=	Replace(Replace(ReplaceStr ,"\","\\"),"'","\'")
			ReplaceStr	=	Replace(ReplaceStr&"","<","&lt;")
			ReplaceStr	=	Replace(ReplaceStr,">","&gt;")
			SQL(23,i)	=	ReplaceStr
			ReplaceStr	=	SQL(19,i)&""
			ReplaceStr	=	Replace(Replace(ReplaceStr ,"\","\\"),"'","\'")
			SQL(19,i)	=	ReplaceStr
			MyTempStr	=	MyTempStr & "','"&SQL(13,i)&"','"&SQL(14,i)&"','"&SQL(15,i)&"','"&SQL(16,i)&"','"& SQL(23,i) &"','"&Topic_Ads&"','"&SQL(19,i)&"','"&UserSign&"','"&SQL(30,i)&"','"&TopicBody&"'"
			MyTempStr	= Dvbbs.ChkBadWords(MyTempStr)
			MyTempStr	= Replace(Replace(Replace(MyTempStr,chr(13),""),chr(10),""),"$","&#36;")
			.Write MyTempStr
			.Write ",'"
			.Write SQL(4,i)
			.Write "','"
			.Write SQL(6,i)
			.Write "','"
			If Dvbbs.GroupSetting(30)="0" Then
				.Write "*.*.*.*"
			Else
				.Write SQL(7,i)
			End If
			.Write  "','"
			.Write AnnounceID
			.Write "',"
			.Write SQL(9,i)
			.Write ","
			.Write SQL(10,i)
			.Write ","
			.Write SQL(11,i)
			.Write ","
			.Write SQL(12,i)
			.Write ",'"
			.Write SQL(18,i)
			.Write "','"
			.Write SQL(20,i)
			.Write "','"
			.Write SQL(21,i)
			.Write "','"
			.Write SQL(22,i)
			.Write "','"
			REM 修正因用户注册时间为空值时出错 2004-5-22 Dv.Yz
			If Not Isdate(SQL(24,i)) Then
				.Write FormatDateTime(Now(),2)
			Else
				.Write FormatDateTime(SQL(24,i),2)
			End If
			.Write "','"
			.Write SQL(25,i)
			.Write "','"
			.Write SQL(26,i)
			.Write "','"
			.Write SQL(27,i)
			.Write "','"
			.Write SQL(28,i)
			.Write "','"
			.Write SQL(29,i)
			.Write "',"
			.Write SQL(31,i)
			.Write ",'"
			.Write SQL(32,i)
			.Write "','"
			.Write SQL(33,i)
			.Write "',"
			.Write SQL(34,i)
			.Write ",'"
			.Write SQL(35,i)
			.Write "','"
			.Write SQL(38,i)
			.Write "',"
			.Write i
			.Write ",'"
			.Write bgcolor
			.Write "','"
			If SQL(37,i)=1 Or DateDiff("s",SQL(35,i),Now())>Cint(dvbbs.Forum_Setting(8))*60 Then
				.Write  "0"
			Else
				.Write  "1"
			End If
			.Write "','"
			.Write SQL(40,i)
			.Write  "'));"
			UbbLists=""
			.Write  "</Script>"
		Next
		SQL=Null
		.Write "<Script Language=JavaScript>"
		MyTempStr = template.html(5)
		MyTempStr = Replace(MyTempStr,"{$width}",Dvbbs.mainsetting(0))
		MyTempStr = Replace(MyTempStr,"{$boardid}",Dvbbs.BoardID)
		MyTempStr = Replace(MyTempStr,"{$replyid}",ReplyID)
		MyTempStr = Replace(MyTempStr,"{$announceid}",AnnounceID)
		MyTempStr = Replace(MyTempStr,"{$skin}",Skin)
		MyTempStr = Replace(MyTempStr,"{$page}",Page)
		MyTempStr = Replace(MyTempStr,"{$topicnum}",TopicCount)
		MyTempStr = Replace(MyTempStr,"{$boardjump}",Dvbbs.BoardJumpList)
		.Write MyTempStr
		.Write  "</Script>"
	End If
	End With
End Function

Function Show_Topic_FastRe()
	Dim TempStr
	With Response
	.Write "<Script Language=JavaScript>"
	.Write "var Board_Setting5='"&Dvbbs.Board_Setting(5)&"';"
	.Write "var Board_Setting6='"&Dvbbs.Board_Setting(6)&"';"
	.Write "var Board_Setting7='"&Dvbbs.Board_Setting(7)&"';"
	.Write "var Board_Setting8='"&Dvbbs.Board_Setting(8)&"';"
	.Write "var Board_Setting9='"&Dvbbs.Board_Setting(9)&"';"
	.Write "var Board_Setting16='"&Dvbbs.Board_Setting(16)&"';"
	.Write "var Board_Setting44='"&Dvbbs.Board_Setting(44)&"';"
	.Write "var Forum_Setting3='"&Dvbbs.Forum_Setting(3)&"';"
	.Write "var Forum_PostFace='"&Dvbbs.Forum_PostFace&"';"
	.Write "var Forum_PostFace=Forum_PostFace.split(""|||"");"
	.Write "</Script>"
	TempStr = template.html(6)
	TempStr = Replace(TempStr,"{$topic}",Topic)
	TempStr = Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
	If Dvbbs.Board_Setting(4)="0" Then
		'Dim re
		'Set re=new RegExp
		're.IgnoreCase =True
		're.Global=True
		're.Pattern="<(.[^>]*)>"
		'Topic=re.Replace(Topic,"")	
		'Set re=Nothing
		'Topic=server.htmlencode(Topic)
		'Topic=Left(Topic,50)
		'TempStr = Replace(TempStr,"{$getcode}","&nbsp;<B>标题：</B><input name=""topic"" size=20 class=FormClass value=""Re："&Topic&""">")
		TempStr = Replace(TempStr,"{$getcode}","")
	Else
		TempStr = Replace(TempStr,"{$getcode}","&nbsp;<B>验证码：</B>"&Dvbbs.GetCode())
	End If
	TempStr = Replace(TempStr,"{$membername}",Dvbbs.membername)
	TempStr = Replace(TempStr,"{$followup}",followup)
	TempStr = Replace(TempStr,"{$announceid}",AnnounceID)
	TempStr = Replace(TempStr,"{$star}",Star)
	TempStr = Replace(TempStr,"{$totalusetable}",TotalUseTable)
	TempStr = Replace(TempStr,"{$Forum_Emot}",Replace(Dvbbs.Forum_emot&"","|||","<><><>"))
	TempStr = Replace(TempStr,"{$MaxLength}",Clng(Dvbbs.Board_Setting(16)))
	Dim Content
	Content=Session(Dvbbs.CacheName & "UserID")
	If IsArray(Content) And Dvbbs.userID > 0 Then
		TempStr = Replace(TempStr,"{$content}",Server.HTMLEncode(Content(37)))
	Else
		TempStr = Replace(TempStr,"{$content}","")
	End If 
	.Write TempStr
	TempStr = ""
	.Cookies("Dvbbs")=""
	End With
End Function

Function Show_Topic_ManageAction()
	Dim TempStr
	TempStr = template.html(7)
	TempStr = Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	TempStr = Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
	TempStr = Replace(TempStr,"{$announceid}",AnnounceID)
	TempStr = Replace(TempStr,"{$replyid}",ReplyID)
	If IsTop > 0 Then
		TempStr = Replace(TempStr,"{$topstr}",template.Strings(10))
	Else
		TempStr = Replace(TempStr,"{$topstr}",template.Strings(11))
	End If
	Response.Write TempStr
	TempStr = ""
End Function

Function Show_Topic_Vote()
	Dim TempStr,Rs,Trs
	Set Rs=Dvbbs.Execute("Select * From Dv_Vote Where VoteID="&PollID)
	If Not (Rs.Eof And Rs.Bof) Then
		Response.Write "<Script Language=JavaScript>"
		Response.Write "var vote='"&Rs("vote")&"';"
		Response.Write "var votenum='"&Rs("votenum")&"';"
		Response.Write "var votetype='"&Rs("votetype")&"';"
		Response.Write "var voters='"&Rs("voters")&"';"
		Response.Write "</Script>"
		TempStr = template.html(8)
		TempStr = Replace(TempStr,"{$topic}",Topic)
		TempStr = Replace(TempStr,"{$announceid}",AnnounceID)
		TempStr = Replace(TempStr,"{$votetype}",Rs("votetype"))
		If Dvbbs.UserID=0 Or datediff("d",rs("timeout"),Now())>0 Or locktopic=1 Then
			TempStr = Replace(TempStr,"{$uservoteinfo}",Split(template.html(9),"||")(0))
		Else
			Set Trs=Dvbbs.Execute("Select Count(*) From Dv_voteuser Where voteid="&PollID&" And userid="&Dvbbs.userid)
			If Trs(0)=0 Then 
				TempStr = Replace(TempStr,"{$uservoteinfo}",Split(template.html(9),"||")(1))
			Else
				TempStr = Replace(TempStr,"{$uservoteinfo}",Split(template.html(9),"||")(2))
			End If
			Set Trs=Nothing 
		End  If
		TempStr = Replace(TempStr,"{$timeout}",Rs("timeout"))
		TempStr = Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
		TempStr = Replace(TempStr,"{$voteid}",PollID)
		TempStr = Replace(TempStr,"{$uarticle}",Rs("UArticle")&"")
		TempStr = Replace(TempStr,"{$uep}",Rs("UEP")&"")
		TempStr = Replace(TempStr,"{$ucp}",Rs("UCP")&"")
		TempStr = Replace(TempStr,"{$upower}",Rs("UPower")&"")
		TempStr = Replace(TempStr,"{$umoney}",Rs("UWealth")&"")
		'Response.Write TempStr
		Show_Topic_Vote = TempStr
		TempStr = ""
	End If
	Set Rs=Nothing
End Function

Function SimJsReplace(str)
	If IsNull(str) Or str="" Then Exit Function
	str=Replace(str,"\","\\")
	str=Replace(str,"'","\'")
	str=Replace(str,"$","&#36;")
	SimJsReplace=str
End Function
Function Dvbbs_TopicMode(str,tmode)
	Select Case tmode
	Case "1"
		Dvbbs_TopicMode=str
	Case "2"
		Dvbbs_TopicMode="<font color=red>"&Dvbbs.Htmlencode(str)&"</font>"
	Case "3"
		Dvbbs_TopicMode="<font color=blue>"&Dvbbs.Htmlencode(str)&"</font>"
	Case "4"
		Dvbbs_TopicMode="<font color=green>"&Dvbbs.Htmlencode(str)&"</font>"
	Case Else
		Dvbbs_TopicMode=Dvbbs.HtmlEncode(str)
	End Select
End Function
Sub Showtree()
	template.html(10) = Replace(template.html(10),"{$boardid}",Dvbbs.BoardID)
	template.html(10) = Replace(template.html(10),"{$replyid}",ReplyID)
	template.html(10) = Replace(template.html(10),"{$announceid}",AnnounceID)
	template.html(10) = Replace(template.html(10),"{$openid}",followup)
	Response.Write template.html(10)
End Sub
Function AnnounceIDlist()
	Dim Rs,SQL,i,starcount
	starcount=(Star-1)*Dvbbs.Board_Setting(27)
	SQL="Select Announceid From "&TotalUseTable&" Where BoardID="&Dvbbs.BoardID&" And RootID="&Announceid&" Order By AnnounceID"
	Set Rs=Dvbbs.Execute(SQL)
	If Not Rs.Eof Then
		Rs.Move Starcount
		REM 修正最后页面出错信息 2004-5-22 Dv.Yz
		If Rs.Eof Then
			Dvbbs.AddErrcode(33)
			Dvbbs.Showerr()
		End If
		AnnounceIDlist = Rs(0)
		Rs.Movenext
		For i = 1 To Dvbbs.Board_Setting(27)
			If Rs.Eof Then Exit For
			AnnounceIDlist = AnnounceIDlist & "," & Rs(0)
			Rs.Movenext
		Next
	Else
		Dvbbs.AddErrcode(32)
		Dvbbs.Showerr()
	End If 
	Set Rs=Nothing 
End Function 
%>
