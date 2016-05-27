<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<%
Dim Page
Page=Request("Page")
If isNumeric(Page) = 0 or Page="" Then Page=1
Page=Clng(Page)
Dim BrowserType
Set BrowserType=New Cls_Browser
If BrowserType.IsSearch Then Response.redirect "List_show.asp?BoardID="&Dvbbs.BoardID&"&page="&page
Set BrowserType=Nothing

If Cint(Dvbbs.Board_Data(6,0)) > 0 Then
	Dvbbs.LoadTemplates("index")
Else
	Dvbbs.LoadTemplates("list")
End If
If Cint(Dvbbs.Board_Setting(43))=0 Then
	Dvbbs.Stats=Dvbbs.LanStr(7)
Else
	Dvbbs.Stats=Dvbbs.LanStr(8)
End If
Dvbbs.Nav()
Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
Dvbbs.Showerr()
Dim action
Dim TopicNum,n,SplitPageNum
Dim Forum_AllTopNum
Forum_AllTopNum = 0
If Dvbbs.boardmaster or Dvbbs.master or Dvbbs.superboardmaster Then
	action=Request("action")
ElseIf Dvbbs.GroupSetting(45)=1 Then
	action=Request("action")
Else
	action=""
End If
Page=Request("Page")
If isNumeric(Page) = 0 or Page="" Then Page=1
Page=Clng(Page)
'如果有下属版面，则显示
If Cint(Dvbbs.Board_Data(6,0)) > 0 Then
	GetChildBoardList
	Dvbbs.LoadTemplates("list")
End If
Dim BoardTopic,BoardTopicImg,BoardTopicMode,BoardTopicMode_a,iii,TopicMode,SelectBoardTopic
TopicMode=0
BoardTopic=Split(Dvbbs.Board_Setting(48),"$$")
BoardTopicImg=Split(Dvbbs.Board_Setting(49),"$$")
If Ubound(BoardTopic)>0 Then
	If Request("topicmode")<>"" and IsNumeric(Request("topicmode")) Then TopicMode=Cint(Request("topicmode"))
	For iii=0 to Ubound(BoardTopic)-1
		If BoardTopicImg(iii)<>"" and Instr(BoardTopicImg(iii),".gif") Then BoardTopicMode=BoardTopicMode+"<img src="&BoardTopicImg(iii)&" border=0  align=absmiddle>"
		BoardTopicMode=BoardTopicMode+"<a href=list.asp?boardid="&Dvbbs.boardid&"&topicmode="&iii+1&">["
		BoardTopicMode_a=BoardTopicMode_a+"<a href=list.asp?boardid="&Dvbbs.boardid&"&topicmode="&iii+1&">["
		If TopicMode=iii+1 Then
			BoardTopicMode=BoardTopicMode+"<font color="&Dvbbs.mainsetting(1)&">"&BoardTopic(iii)&"</font>"
			BoardTopicMode_a=BoardTopicMode_a+"<font color="&Dvbbs.mainsetting(1)&">"&BoardTopic(iii)&"</font>"
		Else 
			BoardTopicMode=BoardTopicMode+BoardTopic(iii)
			BoardTopicMode_a=BoardTopicMode_a+BoardTopic(iii)
		End If
		BoardTopicMode=BoardTopicMode+"]</a>"
		BoardTopicMode_a=BoardTopicMode_a+"]</a>"
		SelectBoardTopic=SelectBoardTopic+"<option value="&(iii+1)
		SelectBoardTopic=SelectBoardTopic+" >"&BoardTopic(iii)&"</option>"
		If iii<>(Ubound(BoardTopic)-1) Then
			BoardTopicMode=BoardTopicMode+ " | "
			BoardTopicMode_a=BoardTopicMode_a+ " | "
		End If
	Next
End If
'分版浮动广告
If Dvbbs.Forum_ads(2)="1" or Dvbbs.Forum_ads(13)="1" Then Response.Write "<script language=""javascript"" src=""inc/Dv_Adv.js""></script>"
If Dvbbs.Board_Setting(43)="0" Then
	Call News
	Call Board_Online
	Call Show_List_Top
	Call Show_List_TopTopic
	Call Show_List_Topic
	Call Show_List_Footer
Else
	Response.Write "<iframe width=""0"" height=""0"" src="""" name=""hiddenframe""></iframe>"
End If
Dvbbs.ActiveOnline()
Dvbbs.Footer()

Function news()
	Dim TempStr,SQL
	TempStr=Dvbbs.Board_Data(23,0)
	SQL=Split(TempStr,"|||")
	If UBound(SQL)< 2 Then
		Dvbbs.Name = "BoardInfo_" & Dvbbs.BoardID
		Dvbbs.LoadBoardNews_Paper Dvbbs.BoardID
		Dvbbs.Board_Data=Dvbbs.Value
		TempStr=Dvbbs.Board_Data(23,0)
		SQL=Split(TempStr,"|||")
	End If
	Dim tmpdata,nexhour
	TempStr=template.html(0)
	If Dvbbs.Board_Setting(21)="1" Then
		tmpdata=split(Dvbbs.Board_Setting(22),"|")
		nexhour=Hour(Now())+1
		nexhour=nexhour mod 24
		If tmpdata(nexhour)="0" And Minute(now())>40 Then
			sql(1)=sql(1)&"--本版将于"&(60-Minute(now()))&"分钟后暂停开放,敬请留意"
		End If
	End If
	TempStr=Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	TempStr=Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
	TempStr=Replace(TempStr,"{$news}",SQL(0)&"")
	TempStr=Replace(TempStr,"{$newstime}",SQL(1))
	Response.Write TempStr
	TempStr="":SQL=Null
End Function
Function Board_online()
	Dim TempStr
	TempStr=template.html(1)
	TempStr=Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
	TempStr=Replace(TempStr,"{$allonline}",MyBoardOnline.Forum_Online)
	TempStr=Replace(TempStr,"{$boardtype}",Dvbbs.Boardtype)
	TempStr=Replace(TempStr,"{$boardonline}",MyBoardOnline.Board_UserOnline)
	TempStr=Replace(TempStr,"{$boardguest}",MyBoardOnline.Board_GuestOnline)
	TempStr=Replace(TempStr,"{$todaynum}",Dvbbs.Board_Data(12,0))
	TempStr=Replace(TempStr,"{$alertcolor}",Dvbbs.mainsetting(1))
	Response.Write TempStr
	TempStr=""
	If Dvbbs.forum_setting(14)="1" Or Dvbbs.forum_setting(15)="1" Then 
		Response.Write "<iframe width=""0"" height=""0"" src=""Online.asp?action=1&Boardid="&Dvbbs.Boardid&""" name=""hiddenframe""></iframe>"
	Else
		Response.Write "<iframe width=""0"" height=""0"" src="""" name=""hiddenframe""></iframe>"
	End If
End Function

Function Show_List_Top()
	Dim TempStr,TempBoardMaster,BoardMaster,i
	If Dvbbs.BoardMaster="" Then
		BoardMaster=template.Strings(4)
	Else
		TempBoardMaster=Split(Dvbbs.BoardMasterList & "","|")
		For i=0 To Ubound(TempBoardMaster)
			BoardMaster = BoardMaster & "<a href=dispuser.asp?name="&TempBoardMaster(i)&">"&TempBoardMaster(i)&"</a>&nbsp;"
		Next
	End If
	If (Dvbbs.Board_Setting(43)="0" And Dvbbs.Board_Setting(0)="0") Or (Dvbbs.Board_Setting(43)="0" And Dvbbs.Board_Setting(0)="1" And (Dvbbs.Master Or Dvbbs.SuperBoardMaster Or Dvbbs.BoardMaster)) Then
		TempStr=template.html(3)
		TempStr=Replace(TempStr,"{$pic_postnew}",Dvbbs.mainpic(7))
		TempStr=Replace(TempStr,"{$pic_postvote}",Dvbbs.mainpic(8))
		TempStr=Replace(TempStr,"{$pic_postxzb}",Dvbbs.mainpic(9))
	Else
		If Dvbbs.Board_Setting(0)="1" Then TempStr=template.Strings(1)
	End If
	TempStr=Replace(template.html(2),"{$showpostinfo}",TempStr)
	TempStr=Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
	TempStr=Replace(TempStr,"{$page}",page)
	TempStr=Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	TempStr=Replace(TempStr,"{$alertcolor}",Dvbbs.mainsetting(1))
	TempStr=Replace(TempStr,"{$boardmasterlist}",BoardMaster)
	TempStr=Replace(TempStr,"{$smallpaper}",Split(Dvbbs.Board_Data(23,0),"|||")(2))
	If Dvbbs.Board_Setting(3)="1" Or Dvbbs.Board_Setting(57)="1" Then
		Dim allaudit,rs
		Set rs=dvbbs.execute("select count(*) from "&Dvbbs.Nowusebbs&" where boardid=777 and locktopic="&Dvbbs.BoardID)
		allaudit=rs(0)
		If IsNull(allaudit) Then allaudit=0
		Set Rs=Nothing
		TempStr=Replace(TempStr,"{$isaudit}","| <a href=AccessTopic.asp?boardid="&Dvbbs.BoardID&" title="&Replace(template.Strings(3),"{$auditnum}",allaudit)&">"&template.Strings(2)&"</a>(<font color="&Dvbbs.mainsetting(1)&">"&allaudit&"</font>)")
	Else
		TempStr=Replace(TempStr,"{$isaudit}","")
	End If
	If BoardTopicMode="" Then
		TempStr=Replace(TempStr,"{$topictype}","")
	Else
		TempStr=Replace(TempStr,"{$topictype}",template.html(14))
		TempStr=Replace(TempStr,"{$TopicMode}",BoardTopicMode)
	End If
	Response.Write TempStr
	TempStr=""
End Function

Function Show_List_Footer()
	Dim TempStr
	TempStr=template.html(5)
	TempStr=Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
	TempStr=Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	TempStr=Replace(TempStr,"{$boardjump}",Dvbbs.BoardJumpList)
	TempStr=Replace(TempStr,"{$forumname}",Dvbbs.Forum_Info(0))
	TempStr=Replace(TempStr,"{$timestr}",Dvbbs.Forum_Info(9))
	TempStr=Replace(TempStr,"{$pic_toptopic}",Dvbbs.mainpic(1))
	TempStr=Replace(TempStr,"{$pic_atoptopic}",Dvbbs.mainpic(0))
	TempStr=Replace(TempStr,"{$pic_opentopic}",Dvbbs.mainpic(2))
	TempStr=Replace(TempStr,"{$pic_hottopic}",Dvbbs.mainpic(3))
	TempStr=Replace(TempStr,"{$pic_locktopic}",Dvbbs.mainpic(4))
	TempStr=Replace(TempStr,"{$pic_besttopic}",Dvbbs.mainpic(5))
	TempStr=Replace(TempStr,"{$pic_votetopic}",Dvbbs.mainpic(6))
	TempStr=Replace(TempStr,"{$pic_toptopic1}",Dvbbs.mainpic(19))
	Response.Write TempStr
	TempStr=""
	'广告代码
	Response.Write "<script language=""javascript"">"
	If Dvbbs.Forum_ads(2)="1" Then
		 Response.Write "move_ad('"&Dvbbs.Forum_ads(3)&"','"&Dvbbs.Forum_ads(4)&"','"&Dvbbs.Forum_ads(5)&"','"&Dvbbs.Forum_ads(6)&"');"
	End If
	If Dvbbs.Forum_ads(13)="1" Then
		Response.Write "fix_up_ad('"& Dvbbs.Forum_ads(8) & "','" & Dvbbs.Forum_ads(10) & "','" & Dvbbs.Forum_ads(11) & "','" & Dvbbs.Forum_ads(9) & "');"		
	End If 
	Response.Write "</script>"
End Function

Function Show_List_TopTopic()
With Response
	.Write "<Script Language=JavaScript>"
	Dim PostTime,ListMainTemplate
	ListMainTemplate=template.html(6)
	If Dvbbs.Board_Setting(60)="0" or Dvbbs.Board_Setting(60)="" Then
		ListMainTemplate=Replace(ListMainTemplate,"{$ShowNewPic}","")
	End If
	.Write Replace(Replace(template.html(11),"{$ShowNewPic}",Dvbbs.Board_Setting(60)),"{$IcoLimMinute}",Dvbbs.Board_Setting(61))
	.Write "var TempStr='"&Replace(Replace(Replace(Replace(ListMainTemplate,"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write "var TempStr_Page='"&Replace(Replace(Replace(Replace(template.html(7),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write "var TempStr_topicinfo='"&Replace(Replace(Replace(Replace(template.html(8),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write "var TempStr_load='"&Replace(Replace(Replace(Replace(template.html(9),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write "var topicpage='"&Dvbbs.Forum_Setting(44)&"';"
	.Write "var alertcolor='"&Dvbbs.mainsetting(1)&"';"
	.Write "var ztopic='"&Dvbbs.mainpic(0)&"';"
	.Write "var istopic='"&Dvbbs.mainpic(1)&"';"
	.Write "var opentopic='"&Dvbbs.mainpic(2)&"';"
	.Write "var hottopic='"&Dvbbs.mainpic(3)&"';"
	.Write "var ilocktopic='"&Dvbbs.mainpic(4)&"';"
	.Write "var besttopic='"&Dvbbs.mainpic(5)&"';"
	.Write "var votetopic='"&Dvbbs.mainpic(6)&"';"
	.Write "var picnofollow='"&Dvbbs.mainpic(10)&"';"
	.Write "var picfollow='"&Dvbbs.mainpic(11)&"';"
	If TopicMode>0 Then
		Set Rs=Dvbbs.Execute("Select count(Topicid) From Dv_topic Where Boardid="&Dvbbs.Boardid&" and mode="&TopicMode)
		TopicNum=Rs(0)
		Rs.close:Set Rs=Nothing
	Else
		TopicNum=Dvbbs.Board_Data(10,0)
	End If
	SplitPageNum=Dvbbs.Board_Setting(26)
	.Write "var TopicNum='"&TopicNum&"';"
	.Write "var page='"&page&"';"
	.Write "var Board_Setting26='"&SplitPageNum&"';"
	.Write "var Board_Setting27='"&Dvbbs.Board_Setting(27)&"';"
	.Write "var BoardTopicMode='"&BoardTopicMode_a&"';"
	.Write "var TopicLimByte='"&Dvbbs.Board_Setting(25)&"';"
	.Write "var MyAction='"&action&"';"
	.Write "var GroupSetting45='"&Dvbbs.GroupSetting(45)&"';"
	.Write "var ListTopicMode='"&TopicMode&"';"
	.Write "var TrueBoardID="&Dvbbs.BoardID&";"
	If TopicMode>0 Then .Write "var BoardTopicMode='';"
	.Write "</Script>"
	If Page=1 Then
		Forum_AllTopNum=Dvbbs.CacheData(28,0)
		If Trim(Dvbbs.Board_Data(20,0))<>"" Then
			If Trim(Forum_AllTopNum)<>"" Then
				Forum_AllTopNum = Forum_AllTopNum & "," & Dvbbs.Board_Data(20,0)
			Else
				Forum_AllTopNum = Dvbbs.Board_Data(20,0)
			End If
		End If
		If Trim(Forum_AllTopNum)<>"" Then
			Dim Rs,SQL,i,TopicTempStr,Showtitle,postusername
			Set Rs=Dvbbs.Execute("Select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode from dv_topic Where istop>0 and TopicID in ("&Forum_AllTopNum&") Order By istop desc, Lastposttime Desc")
			If Rs.Eof And Rs.Bof Then
				Forum_AllTopNum = 0
			Else
				SQL=Rs.GetRows(-1)
				Forum_AllTopNum = 0
				For i=0 To Ubound(SQL,2)
					.Write "<Script Language=JavaScript>"
					Showtitle=SQL(2,i)
					Showtitle=Replace(Showtitle,"\","\\")
					Showtitle=Replace(Showtitle,"""","\""")
					Showtitle=Replace(Showtitle,"'","\'")
					Showtitle=Replace(Showtitle,"$","＄")
					If SQL(16,i)=1 Then
						If Dv_FilterJS(Showtitle) Then
							Showtitle=Replace(Showtitle,"<","&lt;")
							Showtitle = Replace(Showtitle,">","&gt;")
						End If
					Else
						Showtitle=Replace(Showtitle,"<","&lt;")
						Showtitle = Replace(Showtitle,">","&gt;")
					End If 
					postusername=SQL(3,i)
					postusername=Replace(postusername,"\","\\")
					'postusername=Replace(postusername,"""","\""")
					postusername=Replace(postusername,"'","\'")
					TopicTempStr = ",'"&Showtitle&"','"&postusername&"','"&Replace(Replace(Replace(SQL(9,i),"\","\\"),"'","\'"),"<","&lt;")&"','"& SQL(15,i) &"',"
					TopicTempStr = Replace(Replace(Replace(Dvbbs.ChkBadWords(TopicTempStr),VbCrLf,"\n"),chr(13),""),chr(10),"")
					'If SQL(16,i)=1 Then 
					'	TopicTempStr = Replace(TopicTempStr,"<!--","&lt;!--")
					'Else
					'	TopicTempStr = Replace(TopicTempStr,"<","&lt;")
					'	TopicTempStr = Replace(TopicTempStr,">","&gt;")
					'End If
					If Dvbbs.Board_Setting(38) = "0" Then
						PostTime = Split(SQL(9,i),"$")(2)		'最后跟帖时间
					Else
						PostTime = SQL(5,i)		'帖子发表时间
					End If
					.Write "document.write (dvbbs_topic_list(TempStr,'"
					.Write SQL(0,i)
					.Write "','"
					.Write SQL(1,i)
					.Write "'"
					.Write TopicTempStr
					.Write "'"
					.Write SQL(4,i)
					.Write "','"
					.Write SQL(5,i)
					.Write "','"
					.Write SQL(6,i)
					.Write "','"
					.Write SQL(7,i)
					.Write "','"
					.Write SQL(8,i)
					.Write "','"
					.Write SQL(10,i)
					.Write "','"
					.Write SQL(11,i)
					.Write "','"
					.Write SQL(12,i)
					.Write "','"
					.Write SQL(13,i)
					.Write "','"
					.Write SQL(14,i)
					.Write "','"
					.Write SQL(16,i)
					.Write "','"
					.Write SQL(17,i)
					.Write "','"
					If IsDate(PostTime) Then
						.Write DateDiff("n",PostTime,now)+cint(Dvbbs.Forum_setting(0))
					End If
					.Write "'));"
					.Write "hiddentr('follow"
					.Write SQL(0,i)
					.Write "');"
					.Write "</Script>"
					Forum_AllTopNum = Forum_AllTopNum + 1
				Next
				SQL=Null
			End If
			Rs.Close
			Set Rs=Nothing
		Else
			Forum_AllTopNum = 0
		End If	
	Else
		Forum_AllTopNum = 0
	End If
End With
End Function

Function Show_List_Topic()
	Dim Cmd,limitime,SQL,Rs,i,TempStr,ti,TopicTempStr
	Dim Posttime
	If IsSqlDataBase=1 And IsBuss=1 Then
		Set Cmd = Server.CreateObject("ADODB.Command")
		Set Cmd.ActiveConnection=conn
		Cmd.CommandText="dv_list"
		Cmd.CommandType=4
		Cmd.Parameters.Append cmd.CreateParameter("@boardid",3)
		Cmd.Parameters.Append cmd.CreateParameter("@pagenow",3)
		Cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
		Cmd.Parameters.Append cmd.CreateParameter("@tl",3)
		Cmd.Parameters.Append cmd.CreateParameter("@topicmode",3)
		Cmd.Parameters.Append cmd.CreateParameter("@totalrec",3,2)
		Cmd("@boardid")=Dvbbs.BoardID
		Cmd("@pagenow")=page
		Cmd("@pagesize")=Cint(Dvbbs.Board_Setting(26))
		Cmd("@topicmode")=TopicMode
		If limitime="" Then
			Cmd("@tl")=0
		Else
			Cmd("@tl")=limitime
		End If
		set Rs=Cmd.Execute
	Else
		Set Rs = server.CreateObject ("adodb.recordset")
		If Cint(TopicMode)=0 Then
		Sql="Select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode From Dv_Topic Where BoardID="&Dvbbs.BoardID&" And IsTop=0 Order By LastPostTime Desc"
		Else
		Sql="Select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode From Dv_Topic Where BoardID="&Dvbbs.BoardID&" And IsTop=0 And Mode="&TopicMode&" Order By LastPostTime Desc"
		End If
		Rs.Open Sql,Conn,1,1
	End If
	Dvbbs.SqlQueryNum = Dvbbs.SqlQueryNum + 1
	If Not (Rs.Eof And Rs.Bof) Then
		If IsSqlDatabase = 1 And IsBuss=1 Then
			SQL=Rs.GetRows(-1)
		Else
			If TopicNum Mod Cint(Dvbbs.Board_Setting(27))=0 Then
				n = TopicNum \ Cint(Dvbbs.Board_Setting(27))
			Else
	     		n = TopicNum \ Cint(Dvbbs.Board_Setting(27))+1
  			End If
			Rs.MoveFirst
			If page > n Then page = n
			If page < 1 Then page = 1
			If page >1 Then 				
				Rs.Move (page-1) * Clng(Dvbbs.Board_Setting(26))
			End if
			If Rs.Eof Then Exit Function
			SQL=Rs.GetRows(Dvbbs.Board_Setting(26))
		End If
		'TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode
		Dim Showtitle,postusername
		With Response
		For ti=0 To Ubound(SQL,2)
			.Write "<Script Language=JavaScript>"
			Showtitle=SQL(2,ti)
			Showtitle=Replace(Showtitle,"\","\\")
			Showtitle=Replace(Showtitle,"""","\""")
			Showtitle=Replace(Showtitle,"'","\'")
			Showtitle=Replace(Showtitle,"$","＄")
			If SQL(16,ti)=1 Then
				If Dv_FilterJS(Showtitle) Then
					Showtitle=Replace(Showtitle,"<","&lt;")
					Showtitle = Replace(Showtitle,">","&gt;")
				End If
			Else
				Showtitle=Replace(Showtitle,"<","&lt;")
				Showtitle = Replace(Showtitle,">","&gt;")
			End If
			postusername=SQL(3,ti)
			postusername=Replace(postusername,"\","\\")
			'postusername=Replace(postusername,"""","\""")
			postusername=Replace(postusername,"'","\'")
			TopicTempStr = ",'"&Showtitle&"','"&postusername&"','"&Replace(Replace(Replace(SQL(9,ti),"\","\\"),"'","\'"),"<","&lt;")&"','"& SQL(15,ti) &"',"
			TopicTempStr = Replace(Replace(Replace(Dvbbs.ChkBadWords(TopicTempStr),VbCrLf,""),chr(13),""),chr(10),"")
			'If SQL(16,ti)=1 Then 
			'	TopicTempStr = Replace(TopicTempStr,"<!--","&lt;!--")
			'Else
			'	TopicTempStr = Replace(TopicTempStr,"<","&lt;")
			'	TopicTempStr = Replace(TopicTempStr,">","&gt;")
			'End If				
			If Dvbbs.Board_Setting(38) = "0" Then
				PostTime = Split(SQL(9,ti),"$")(2)		'最后跟帖时间
			Else
				PostTime = SQL(5,ti)		'帖子发表时间
			End If
			.Write "document.write (dvbbs_topic_list(TempStr,'"
			.Write SQL(0,ti)
			.Write "','"
			.Write SQL(1,ti)
			.Write "'"
			.Write TopicTempStr
			.Write "'"
			.Write SQL(4,ti)
			.Write "','"
			.Write SQL(5,ti)
			.Write "','"
			.Write SQL(6,ti)
			.Write "','"
			.Write SQL(7,ti)
			.Write "','"
			.Write SQL(8,ti)
			.Write "','"
			.Write SQL(10,ti)
			.Write "','"
			.Write SQL(11,ti)
			.Write "','"
			.Write SQL(12,ti)
			.Write "','"
			.Write SQL(13,ti)
			.Write "','"
			.Write SQL(14,ti)
			.Write "','"
			.Write SQL(16,ti)
			.Write "','"
			.Write SQL(17,ti)
			.Write "','"
			If IsDate(Posttime) Then
				.Write DateDiff("n",Posttime,now)+cint(Dvbbs.Forum_setting(0))
			End If
			.Write "'));"
			.Write "hiddentr('follow"
			.Write SQL(0,ti)
			.Write "');"
			.Write "</Script>"
		Next
		SplitPageNum=Ubound(SQL,2)+1
		SQL=Null
		
		If TopicNum Mod Cint(SplitPageNum)=0 Then
			n = TopicNum \ Cint(SplitPageNum)
		Else
	     	n = TopicNum \ Cint(SplitPageNum)+1
  		End If
		If action="batch" And Dvbbs.GroupSetting(45)=1 Then
			Dim Forum_Boards,Board_Datas,BoardJumpList,ii,Depth
			Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
			For i=0 To Ubound(Forum_Boards)
				Dvbbs.Name="BoardInfo_" & Forum_Boards(i)
  				If Dvbbs.ObjIsEmpty() Then Dvbbs.ReloadBoardInfo(Forum_Boards(i))
				Board_Datas=Dvbbs.Value
				BoardJumpList = BoardJumpList & "<option value="""&Forum_Boards(i)&""" "
				BoardJumpList = BoardJumpList & ">"
				Depth=Board_Datas(4,0)
				Select Case Depth
				Case 0
					BoardJumpList = BoardJumpList & "╋"
				Case 1
					BoardJumpList = BoardJumpList & "&nbsp;&nbsp;├"
				End Select
				If Depth>1 Then
					For ii=2 To Depth
						BoardJumpList = BoardJumpList & "&nbsp;&nbsp;│"
					Next
					BoardJumpList = BoardJumpList & "&nbsp;&nbsp;├"
				End If
				BoardJumpList = BoardJumpList & Board_Datas(1,0)&"</option>"
			Next
			Board_Datas=Null
			Forum_Boards=Null
			TempStr=template.html(12)
			TempStr=Replace(TempStr,"{$boardjump}",BoardJumpList)
			TempStr=Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
			TempStr=Replace(TempStr,"{$TopicMode}",SelectBoardTopic)
			.Write TempStr
		End If
		.Write "<Script Language=JavaScript>"
		TempStr=template.html(10)
		TempStr=Replace(TempStr,"{$nowpage}",page)
		TempStr=Replace(TempStr,"{$allpage}",n)
		TempStr=Replace(TempStr,"{$pagetopicnum}",SplitPageNum + Forum_AllTopNum)
		TempStr=Replace(TempStr,"{$topicnum}",TopicNum)
		TempStr=Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
		TempStr=Replace(TempStr,"{$myaction}",action)
		TempStr=Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
		.Write TempStr
		.Write "</Script>"
		End With
	End If
	If Forum_AllTopNum = 0 And ti = 0 Then Response.Write template.html(4)
	SQL=Null
	Rs.Close
	Set Rs=Nothing
	Set Cmd=Nothing
End Function

Function Chk_List_Err
	If Dvbbs.BoardID=0 Then
		Dvbbs.AddErrCode(29)
		Exit Function
	End If
	If Cint(Dvbbs.Board_Setting(2))=1 Then
		If Dvbbs.UserID=0 Then
			Dvbbs.AddErrCode(24)
		Else
			If Dvbbs.Board_Setting(46)>0 And Chkboardlogin(Dvbbs.Boardid,dvbbs.Membername)=False Then Response.Redirect "pay_boardlimited.asp?boardid=" & Dvbbs.BoardID
			If Chkboardlogin(Dvbbs.Boardid,dvbbs.Membername)=False Then Dvbbs.AddErrCode(25)
		End If
	End If
	If Cint(Dvbbs.Board_Setting(1))=1 and Cint(Dvbbs.GroupSetting(37))=0 Then Dvbbs.AddErrCode(26)
	
	If Cint(Dvbbs.GroupSetting(0))=0  Then Dvbbs.AddErrCode(27)
	
	If action="batch" Then
		If CInt(Dvbbs.GroupSetting(45))<>1 Then Dvbbs.AddErrCode(28)
	End If
End Function

Sub GetChildBoardList()
	Dim TempListArray,havenew,loadboard,Board_Datas
	TempListArray = Split(template.html(8),"||")
	With Response
	.Write Replace(Replace(template.html(7),"{$follow}",Dvbbs.mainpic(11)),"{$nofollow}",Dvbbs.mainpic(10))
	.Write "<script language=""javascript"">"
	.Write vbNewLine
	'传送图片变量到JS
	For i=0 to UBound(template.pic)-1
		.Write "piclist["&i&"]='"&template.pic(i)&"';"
		.Write vbNewLine		
	Next
	'传递论坛主设置数据到JS
	For i=0 to UBound(Dvbbs.mainsetting)
		.Write "mainsetting["&i&"]='"&Dvbbs.mainsetting(i)&"';"
		.Write vbNewLine	
	Next 
	'传送模板数据到JS以备调用
	.Write "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(4),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write vbNewLine
	.Write "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(0),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write vbNewLine
	.Write "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(5),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write vbNewLine
	.Write "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(1),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write vbNewLine
	.Write "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(6),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write vbNewLine
	.Write "template[template.length]='"&Replace(Replace(Replace(Replace(TempListArray(2),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write vbNewLine
	.Write "template[template.length]='"&Replace(Replace(Replace(Replace(template.html(10),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")&"';"
	.Write vbNewLine
	'传送字符串变量到JS
	For i=0 to 10
		.Write "Strings[Strings.length]='"& template.Strings(i)&"';"		
	Next
	Dim i,BoardID,Rs,ClassID
	Dim setings,lastposttime,depth,lastpost,BoardType,BoardReadme
	Set Rs=Dvbbs.Execute("select boardid,BoardType,ParentID,ParentStr,Depth,RootID,Child,readme,BoardMaster,PostNum,TopicNum,indexIMG,todayNum,boarduser,LastPost,Sid,Board_Setting,Board_Ads,Board_user,IsGroupSetting,BoardTopStr from Dv_board where ParentID="& Dvbbs.BoardID &" Or BoardID = "&Dvbbs.BoardID&" order by orders")
	If Not Rs.Eof Then Board_Datas=Rs.GetRows(-1)
	ClassID=""
	For i=0 To Ubound(Board_Datas,2)
		If Board_Datas(0,i)=Dvbbs.BoardID Then 
			If ClassID<>"" Then 
				.Write "classfooter();"		
			End If
			ClassID=Board_Datas(0,i)
			BoardType=Board_Datas(1,i)
			BoardType=Replace(BoardType,"\","\\")
			BoardType=Replace(BoardType,"'","\'")
			.Write "showclass("&Board_Datas(0,i)&",'"&BoardType&"','"&Board_Datas(16,i)&"','"&Request.Cookies("List")("list"&Board_Datas(0,i))&"',"&Board_Datas(6,i)&");"		
		Else
			havenew=0
			loadboard=True
			setings=split(Board_Datas(16,i),",")(1)
			lastpost=Dvbbs.iHtmlEncode(Board_Datas(14,i))
			lastpost=Replace(Replace(lastpost,Chr(10),""),Chr(13),"")
			lastposttime=split(Board_Datas(14,i),"$")(2)
			If Not IsDate(lastposttime) Then lastposttime=Now()
			If datediff("h",Dvbbs.Lastlogin,lastposttime)=0 Then havenew=1
			If CInt(setings)=1 And CInt(Dvbbs.GroupSetting(37))<>1 Then loadboard=False

			If loadboard Then
				BoardType=Board_Datas(1,i)
				BoardType=Replace(BoardType,"\","\\")
				BoardType=Replace(BoardType,"'","\'")
				BoardReadme=Board_Datas(7,i)&""
				.Write "showboard("&Board_Datas(0,i)&",'"&BoardType&"',"&Board_Datas(6,i)&",'"&BoardReadme&"','"&Board_Datas(8,i)&"',"&Board_Datas(9,i)&","&Board_Datas(10,i)&",'"&Board_Datas(11,i)&"',"&Board_Datas(12,i)&",'"&lastpost&"','"&Board_Datas(16,i)&"',"& havenew &");"
			End If
			If Board_Datas(6,i)>0 Or Not loadboard Then
				.Write "boardcount++;"
				.Write "Child=Child-1;"
				.Write "showcode('','');"
			End If
		End If	
		.Write vbNewLine
	Next
	If ClassID<>"" Then 
		.Write "classfooter();"		
	End If
	Set Rs=Nothing
	Board_Datas = Null
	.Write vbNewLine
	.Write "</script>"
	End With
End Sub
Function Dv_FilterJS(v)
	Dim Re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	If Not Isnull(V) Then
		Dim t1,test,Replacelist
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
			Replacelist="(--|&#([0-9][0-9]*)|function|meta|language|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"
		Else
			Replacelist="("&Dvbbs.forum_setting(77)&"--|&#([0-9][0-9]*)|function|meta|language|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit|error|click|key|load|focus|Blur))"
		End If
		re.Pattern="<((.[^>]*"&Replacelist&"[^>]*)|"&Replacelist&")>"
		Test=re.Test(t1)
		Dv_FilterJS=test
	End If
	Set Re=Nothing
End Function

%>