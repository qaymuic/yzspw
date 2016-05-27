<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<%
If Dvbbs.BoardID = 0 Then
	Response.Write "参数错误"
	Response.End 
End If
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
Dim Page
Dim action
Dim TopicNum,n,SplitPageNum
Dim Forum_AllTopNum
Forum_AllTopNum = 0
action=Request("action")
If Not(Dvbbs.boardmaster or Dvbbs.master or Dvbbs.superboardmaster) Then action=""
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
		BoardTopicMode=BoardTopicMode+"<a href=?boardid="&Dvbbs.boardid&"&topicmode="&iii+1&">["
		BoardTopicMode_a=BoardTopicMode_a+"<a href=?boardid="&Dvbbs.boardid&"&topicmode="&iii+1&">["
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

If Cint(Dvbbs.Board_Setting(43))=0 Then
	News
	Board_Online
	Show_List_Top
	Show_List_TopTopic
	Show_List_Topic
	Show_List_Footer
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
	TempStr=template.html(0)
	TempStr=Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	TempStr=Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
	TempStr=Replace(TempStr,"{$news}",SQL(0))
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
	If Cint(Dvbbs.Board_Setting(3))=1 Then
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
End Function

Function Show_List_TopTopic()
	
	If TopicMode>0 Then
		Set Rs=Dvbbs.Execute("Select count(Topicid) From Dv_topic Where Boardid="&Dvbbs.Boardid&" and mode="&TopicMode)
		TopicNum=Rs(0)
		Rs.close:Set Rs=Nothing
	Else
		TopicNum=Dvbbs.Board_Data(10,0)
	End If
	SplitPageNum=Dvbbs.Board_Setting(26)
	Forum_AllTopNum=Dvbbs.CacheData(28,0)
	If Trim(Dvbbs.Board_Data(20,0))<>"" Then
		If Trim(Forum_AllTopNum)<>"" Then
			Forum_AllTopNum = Forum_AllTopNum & "," & Dvbbs.Board_Data(20,0)
		Else
			Forum_AllTopNum = Dvbbs.Board_Data(20,0)
		End If
	End If
	Dim tmpstr
	If Trim(Forum_AllTopNum)<>"" And Page=1 Then
		Dim Rs,SQL,i,TopicTempStr,Showtitle,postusername
		Set Rs=Dvbbs.Execute("Select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression,TopicMode,Mode from dv_topic Where istop>0 and TopicID in ("&Forum_AllTopNum&") Order By istop desc, Lastposttime Desc")
		If Rs.Eof And Rs.Bof Then
		Forum_AllTopNum = 0
		Else
		SQL=Rs.GetRows(-1)
		Forum_AllTopNum = 0
		For i=0 To Ubound(SQL,2)
			tmpstr=template.html(15)
			postusername=SQL(3,i)
			postusername=Dvbbs.htmlEncode(postusername)
			tmpstr=Replace(tmpstr,"{$userid}",SQL(4,i))
			tmpstr=Replace(tmpstr,"{$username}",postusername)
			tmpstr=Replace(tmpstr,"{$boardid}",SQL(1,i))
			Showtitle=SQL(2,i)
			Showtitle=killhtml(Showtitle)
			Showtitle=Dvbbs.htmlEncode(Showtitle)
			tmpstr=Replace(tmpstr,"{$topic}",Showtitle)
			tmpstr=Replace(tmpstr,"{$linkinfo}","&ID="&SQL(0,i)&"&page="&page&"")
			tmpstr=Replace(tmpstr,"{$lastposttime}",SQL(10,i))
			tmpstr=Replace(tmpstr,"{$hits}",SQL(7,i))
			tmpstr=Replace(tmpstr,"{$child}",SQL(6,i))
			If CLng(SQL(11,i))=2 Then 
				tmpstr=Replace(tmpstr,"{$stats}","区固")
			Else
				tmpstr=Replace(tmpstr,"{$stats}","总固")
			End If
			Response.Write tmpstr		
			Forum_AllTopNum = Forum_AllTopNum + 1
		Next
		SQL=Null
		End If
		Rs.Close
		Set Rs=Nothing
	Else
		Forum_AllTopNum = 0
	End If
End Function

Function Show_List_Topic()
	Dim Cmd,limitime,SQL,Rs,i,TempStr,ti,TopicTempStr,tmpstr
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
		For ti=0 To Ubound(SQL,2)
			tmpstr=template.html(15)
			postusername=SQL(3,ti)
			postusername=Dvbbs.htmlEncode(postusername)
			tmpstr=Replace(tmpstr,"{$userid}",SQL(4,ti))
			tmpstr=Replace(tmpstr,"{$username}",postusername)
			tmpstr=Replace(tmpstr,"{$boardid}",SQL(1,ti))
			Showtitle=SQL(2,ti)
			Showtitle=killhtml(Showtitle)
			Showtitle=Dvbbs.htmlEncode(Showtitle)
			tmpstr=Replace(tmpstr,"{$topic}",Showtitle)
			tmpstr=Replace(tmpstr,"{$linkinfo}","&ID="&SQL(0,ti)&"&page="&page&"")
			tmpstr=Replace(tmpstr,"{$lastposttime}",SQL(10,ti))
			tmpstr=Replace(tmpstr,"{$hits}",SQL(7,ti))
			tmpstr=Replace(tmpstr,"{$child}",SQL(6,ti))
			If CLng(SQL(11,ti))=1 Then 
				tmpstr=Replace(tmpstr,"{$stats}","固顶")
			ElseIf CLng(SQL(13,ti))=1 Then 
				tmpstr=Replace(tmpstr,"{$stats}","精华")
			Else
				tmpstr=Replace(tmpstr,"{$stats}","普通")
			End If
				
			Response.Write tmpstr
		Next
		SplitPageNum=Ubound(SQL,2)+1
		SQL=Null
		
		If TopicNum Mod Cint(SplitPageNum)=0 Then
			n = TopicNum \ Cint(Dvbbs.Board_Setting(26))
		Else
	     		n = TopicNum \ Cint(Dvbbs.Board_Setting(26))+1
  		End If	
  		Dim Endpage
		Endpage=n
			Response.Write "<table border=0 cellpadding=0 cellspacing=3 width="""&Dvbbs.mainsetting(0)&""" align=center>"
			Response.Write "<tr><td valign=middle nowrap>"
			Response.Write "页次：<b>"&page&"</b>/<b>"&n&"</b>页"
			Response.Write "每页<b>"& Dvbbs.Board_Setting(26) &"</b> 贴数<b>"& TopicNum &"</b></td>"
			Response.Write "<td valign=middle nowrap><div align=right><p>分页： <b>"
			If page > 4 Then
				Response.Write "<a href=""?BoardID="&Dvbbs.BoardID&"&page=1"">[1]</a> ..."
			End If
			
			If n >page+3 Then
				Endpage=page+3
			End If
			For i=page-3 to Endpage
				If Not i<1 Then
					If i = CLng(page) Then
						response.write " <font color="&dvbbs.mainsetting(1)&">["&i&"]</font>"
					Else
						Response.Write " <a href=""?BoardID="&Dvbbs.BoardID&"&page="&i&""">["&i&"]</a>"
					End If
				End If
			Next
			If page+3 < n Then
				response.write "... <a href=""?BoardID="&Dvbbs.BoardID&"&page="&n&""">["&n&"]</a></b>"
			End If
			Response.Write "</p></div></td></tr></table>"
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
	Dim Chachedata,ishidden,ShowMasters,j
	Dim Forum_Boards,i,BoardID,Board_Data,ClassID
	Dim setings,lastposttime,depth,lastpost,BoardType,BoardReadme,htmlstr
	template.html(8)=Split(template.html(8),"||")
	ClassID=""
	Dim TempListArray,havenew,loadboard
	Dim Rs
	Set Rs=Dvbbs.Execute("select boardid from Dv_board where ParentID="& Dvbbs.BoardID &" Or BoardID = "&Dvbbs.BoardID&" order by orders")
	If Not Rs.Eof Then Forum_Boards=Rs.GetRows(-1)
	Set Rs=Nothing
	For i=0 to UBound(Forum_Boards,2)
		Dvbbs.Name="BoardInfo_" & Forum_Boards(0,i)
  		If Dvbbs.ObjIsEmpty() Then Dvbbs.ReloadBoardInfo(Forum_Boards(0,i))
		Board_Data=Dvbbs.Value
		If Board_Data(2,0)="0" Then
			BoardType=Board_Data(1,0)&""
			If ClassID<>"" Then 
				Response.Write template.html(8)(1)		
				Response.Write "<br>"
			End If
			ClassID=Forum_Boards(0,i)
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
						Board_Data(11,0)="<table align=""left""><tr><td><a href=""?boardid="&Board_Data(0,0)&"""><img src="""&Board_Data(11,0)&""" align=""top"" border=""0""></a></td><td width=""20""></td></tr></table>"
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
Function killhtml(Str)
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="<(.[^>]*)>"
	str=re.Replace(str,"")	
	set re=Nothing
	killhtml=str
End Function 
%>
