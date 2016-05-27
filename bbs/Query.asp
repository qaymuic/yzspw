<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dvbbs.LoadTemplates("query")
If request("stype")="" Then
	Dvbbs.stats=template.Strings(0)
	Dvbbs.nav()
	If DVbbs.BoardID=0 then
		Dvbbs.Head_var 0,0,template.Strings(0),"query.asp"
	Else
		Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
	End If
	If Dvbbs.boardid>0 Then GetBoardPermission
	If Cint(Dvbbs.GroupSetting(14))=0 Then Dvbbs.AddErrCode(60)
	Dvbbs.ShowErr()
	main()
Else
	Dvbbs.Stats=template.Strings(1)

	Dim stype,pSearch,nSearch,keyword,stable,page,searchday,searchboard,page_count,Pcount
	Dim totalrec,endpage,ordername,hidboardid,FobBoardID
	Dim SqlColumn
	Dim SearchMaxPageList
	If Dvbbs.Forum_Setting(12)<>"0" Then
		If IsNumeric(Dvbbs.Forum_Setting(12)) Then
		If Clng(Dvbbs.Forum_Setting(12)) Mod Cint(Dvbbs.Forum_Setting(11))=0 Then
			SearchMaxPageList = Clng(Dvbbs.Forum_Setting(12)) \ Cint(Dvbbs.Forum_Setting(11))
		Else
	     	SearchMaxPageList = Clng(Dvbbs.Forum_Setting(12)) \ Cint(Dvbbs.Forum_Setting(11))+1
  		End If
		Else
			SearchMaxPageList = 50
		End If
	Else
		SearchMaxPageList = 50
	End If

	CheckRequestInfo()
	Dvbbs.ShowErr()
	SearchResult()
	Dvbbs.ShowErr()
End If

Dvbbs.ActiveOnline
Dvbbs.footer()

Sub main()
	Dim TempStr,i
	TempStr = template.html(0)
	Dim Forum_Boards,Board_Data,BoardJumpList,ii,Depth
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.Name="BoardInfo_" & Forum_Boards(i)
  		If Dvbbs.ObjIsEmpty() Then Dvbbs.ReloadBoardInfo(Forum_Boards(i))
		Board_Data=Dvbbs.Value
		BoardJumpList = BoardJumpList & "<option value="""&Forum_Boards(i)&""" "
		If Clng(Forum_Boards(i))=Dvbbs.BoardID Then BoardJumpList = BoardJumpList & "selected"
		BoardJumpList = BoardJumpList & ">"
		Depth=Board_Data(4,0)
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
		BoardJumpList = BoardJumpList & Board_Data(1,0)&"</option>"
	Next
	Board_Data=Null
	Forum_Boards=Null
	TempStr=Replace(TempStr,"{$BoardJumpList}",BoardJumpList)
	
	Dvbbs.name="Tablelist"
	If Dvbbs.ObjIsEmpty() Then
		Dim Rs,Tablelist
		Set Rs=Dvbbs.Execute("select * from Dv_TableList")
		Do while Not Rs.Eof
			Tablelist = Tablelist & "<Option value="&Rs("tablename")&">"&Rs("tabletype")&"</Option>"
		Rs.MoveNext
		Loop
		Set Rs=Nothing
		Dvbbs.value=Tablelist
	End If 
	TempStr=Replace(TempStr,"{$tablelist}",Dvbbs.value)
	If Dvbbs.Forum_Setting(16)<>"0" Then
		TempStr=Replace(TempStr,"{$searchbody}",template.html(5))
	Else
		TempStr=Replace(TempStr,"{$searchbody}","")
	End If
	If Dvbbs.Forum_Setting(4)<>"0" Then
		Dim keywordlimited
		keywordlimited = Split(Dvbbs.Forum_Setting(4),"|")
		If Ubound(keywordlimited)=1 Then
			TempStr=Replace(TempStr,"{$minlength}",keywordlimited(0))
			TempStr=Replace(TempStr,"{$maxlength}",keywordlimited(1))
		Else
			TempStr=Replace(TempStr,"{$minlength}",4)
			TempStr=Replace(TempStr,"{$maxlength}",20)
		End If
	Else
		TempStr=Replace(TempStr,"{$minlength}",4)
		TempStr=Replace(TempStr,"{$maxlength}",20)
	End If
	If Dvbbs.Forum_Setting(3)<>"0" Then
		TempStr=Replace(TempStr,"{$timelimited}",Dvbbs.Forum_Setting(3))
	Else
		TempStr=Replace(TempStr,"{$timelimited}",120)
	End If
	Response.Write TempStr
End Sub

Function CheckRequestInfo()
	Dim i
	stype=Trim(request("stype"))
	pSearch=Trim(request("pSearch"))
	nSearch=Trim(request("nSearch"))
	keyword=Trim(Dvbbs.checkStr(request("keyword")))
	stable=Replace(Request("stable"),"'","")
	If not IsNumeric(pSearch) Then pSearch=1
	If not IsNumeric(nSearch) Then nSearch=1
	If stable="" or len(stable)>20 Then stable=Dvbbs.NowUseBbs
	If request("page")<>"" and IsNumeric(request("page")) Then
		page=Clng(request("page"))
	Else
		page=0
	End If
	If Cint(Dvbbs.GroupSetting(14))=0 Then Dvbbs.AddErrCode(60)
	If Len(stable)>8 Then Dvbbs.AddErrCode(35)
	if stype<3 then
		If keyword="" Then Dvbbs.AddErrCode(61)
		If keyword<>"" Then
			Dim Foundmykeyword
			Foundmykeyword = False
			If Dvbbs.Forum_Setting(9)<>"0" Then
				Dim mykeyword
				mykeyword = Split(Dvbbs.Forum_Setting(9),"|")
				For i = 0 To Ubound(mykeyword)
					If Instr(Lcase(keyword),Lcase(mykeyword(i)))>0 Then
						Foundmykeyword = True
						Exit For
					End If
				Next
			End If
			If Dvbbs.Forum_Setting(4)<>"0" And Not Foundmykeyword And Not (Dvbbs.Master Or Dvbbs.BoardMaster Or Dvbbs.SuperBoardMaster) Then
				Dim keywordlimited
				keywordlimited = Split(Dvbbs.Forum_Setting(4),"|")
				If Ubound(keywordlimited)=1 Then
					If IsNumeric(keywordlimited(0)) Then
						If Len(keyword)<Clng(keywordlimited(0)) Then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(17),"{$minlength}",keywordlimited(0))&"&action=OtherErr"
					End If
					If IsNumeric(keywordlimited(1)) Then
						If Len(keyword)>Clng(keywordlimited(1)) Then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(18),"{$maxlength}",keywordlimited(1))&"&action=OtherErr"
					End If
				End If
			End If
		End If
		
		'搜索多少天内帖子
		If Lcase(request("SearchDate"))="all" Then
			searchday=" "
		Else
			If request("SearchDate")<>"" And IsNumeric(Request("SearchDate"))  Then
				If IsSqlDataBase=1 Then
					searchday=" datediff(d,DateAndTime,"&SqlNowString&") < "&Dvbbs.checkStr(request("SearchDate"))&" and "
				Else
					searchday=" datediff('d',DateAndTime,"&SqlNowString&") < "&Dvbbs.checkStr(request("SearchDate"))&" and "
				End If
			Else
				Dvbbs.AddErrCode(62)
			End If
		End If
	End If
	searchboard = " "
	If Dvbbs.BoardID>0 Then searchboard=" BoardID="&Dvbbs.BoardID&" and "

	'判断隐含板块
	Dim Forum_Boards,Board_Data
	Forum_Boards=Split(Dvbbs.cachedata(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.Name="BoardInfo_" & Forum_Boards(i)
  		If Dvbbs.ObjIsEmpty() Then Dvbbs.ReloadBoardInfo(Forum_Boards(i))
		Board_Data=Dvbbs.Value
		If Split(Board_Data(16,0),",")(1)="1" Then
			If hidboardid="" Then
				hidboardid=Forum_Boards(i)
			Else
				hidboardid=hidboardid & "," & Forum_Boards(i)
			End If
		End If
		If Split(Board_Data(16,0),",")(2)="1" Then
			If FobBoardID="" Then 
				FobBoardID=Forum_Boards(i)
			Else
				FobBoardID=FobBoardID & "," & Forum_Boards(i)
			End If
		End If
	Next
	Board_Data=Null
	Forum_Boards=Null
	'If Not (Dvbbs.GroupSetting(37)="1" and hidboardid="") Then searchboard=searchboard & " Not BoardID in ("&hidboardid&") and "

	Dim FobWords
	'搜索过滤字
	FobWords = Array(91,92,304,305,430,431,437,438,12460,12461,12462,12463,12464,12465,12466,12467,12468,12469,12470,12471,12472,12473,12474,12475,12476,12477,12478,12479,12480,12481,12482,12483,12485,12486,12487,12488,12489,12490,12496,12497,12498,12499,12500,12501,12502,12503,12504,12505,12506,12507,12508,12509,12510,12532,12533,65339,65340)
	For i = 1 to Ubound(FobWords,1)
		If InStr(keyword,ChrW(FobWords(i))) > 0 Then
			Dvbbs.AddErrCode(61)
			Exit For
		End If
	Next
	FobWords = Array("~","!","@","#","$","%","^","&","*","(",")","_","+","=","`","[","]","{","}",";",":","""","'",",","<",">",".","/","\","|","?","_","about","1","2","3","4","5","6","7","8","9","0","a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","after","all","also","an","and","another","any","are","as","at","be","because","been","before","being","between","both","but","by","came","can","come","could","did","do","each","for","from","get","got","had","has","have","he","her","here","him","himself","his","how","if","in","into","is","it","like","make","many","me","might","more","most","much","must","my","never","now","of","on","only","or","other","our","out","over","said","same","see","should","since","some","still","such","take","than","that","the","their","them","then","there","these","they","this","those","through","to","too","under","up","very","was","way","we","well","were","what","where","which","while","who","with","would","you","your","的","一","不","在","人","有","是","为","以","于","上","他","而","后","之","来","及","了","因","下","可","到","由","这","与","也","此","但","并","个","其","已","无","小","我","们","起","最","再","今","去","好","只","又","或","很","亦","某","把","那","你","乃","它")
	keyword = Left(keyword,100)
	keyword = Replace(keyword,"!"," ")
	keyword = Replace(keyword,"]"," ")
	keyword = Replace(keyword,"["," ")
	keyword = Replace(keyword,")"," ")
	keyword = Replace(keyword,"("," ")
	keyword = Replace(keyword,"　"," ")
	keyword = Replace(keyword,"-"," ")
	keyword = Replace(keyword,"/"," ")
	keyword = Replace(keyword,"+"," ")
	keyword = Replace(keyword,"="," ")
	keyword = Replace(keyword,","," ")
	keyword = Replace(keyword,"'"," ")
	For i = 0 To Ubound(FobWords,1)
		If keyword=FobWords(i) Then
			Dvbbs.AddErrCode(61)
			Exit for
		End If
	Next
End Function

Function SQLQueryStr()
	Dim SearchUserID,Rs
	SqlColumn = "Select Top " & Cint(Dvbbs.Forum_Setting(11))*SearchMaxPageList
	If stype=1 And (nSearch=2 or nSearch=3) Then
		SqlColumn = SqlColumn & " BoardID,RootID,Topic,Expression,UserName,PostUserID,DateAndTime,IsBest,LockTopic,Body,AnnounceID From "
	ElseIf stype=2 And pSearch=2 Then
		If IsSqlDataBase Then
		SqlColumn = SqlColumn & " T1.BoardID,T1.RootID,T1.Topic,T1.Expression,T1.UserName,T1.PostUserID,T1.DateAndTime,T1.IsBest,T1.LockTopic,T1.Body,T1.AnnounceID From "
		Else
		SqlColumn = SqlColumn & " BoardID,RootID,Topic,Expression,UserName,PostUserID,DateAndTime,IsBest,LockTopic,Body,AnnounceID From "
		End If
	ElseIf stype=3 Then
'		SqlColumn = "Select Top 50 BoardID,TopicID,Title,Expression,PostUserName,PostUserID,DateAndtime,IsBest,LockTopic From "
		SqlColumn = "Select Top 50 BoardID,rootid,topic,Expression,username,postuserid,dateandtime,IsBest,LockTopic,Body,Announceid from "
	Else
		SqlColumn = SqlColumn & " BoardID,TopicID,Title,Expression,PostUserName,PostUserID,DateAndtime,IsBest,LockTopic From "
	End If
	'Dvbbs.Stats = template.Strings(2)
	Dvbbs.Stats = template.Strings(4)
	
	If Trim(searchday)<>"" Then
		Dvbbs.Stats = Dvbbs.Stats & Replace(template.Strings(5),"{$searchday}",request("SearchDate"))
	Else
		Dvbbs.Stats = Dvbbs.Stats & template.Strings(6)
	End If
	Select Case stype
	Case 1
		Set Rs=Dvbbs.Execute("Select UserID From Dv_User Where UserName='"&keyword&"'")
		If Rs.Eof And Rs.Bof Then
			Set Rs=Nothing
			Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(21)&"&action=OtherErr"
		Else
			SearchUserID = Rs(0)
		End If
		Select Case nSearch
		'主题作者
		Case 1
			SqlColumn = SqlColumn & " dv_Topic Where "&searchboard&" "&searchday&" PostUserID="&SearchUserID&" Order By TopicID Desc"
			Dvbbs.Stats = Dvbbs.Stats & template.Strings(7)
		'回复作者
		Case 2
			SqlColumn = SqlColumn & stable & " Where "&searchboard&" "&searchday&" ParentID>0 And PostUserID="&SearchUserID&" Order By AnnounceID Desc"
			Dvbbs.Stats = Dvbbs.Stats & template.Strings(8)
		'主题和回复作者
		Case 3
			SqlColumn = SqlColumn & stable & " Where "&searchboard&" "&searchday&" PostUserID="&SearchUserID&" Order By AnnounceID Desc"
			Dvbbs.Stats = Dvbbs.Stats & template.Strings(9)
		End Select
	Case 2
		Select Case pSearch
		'标题
		Case 1
			SqlColumn = SqlColumn & " dv_Topic Where "&searchboard&" "&searchday&" Title like '%"&keyword&"%' Order By TopicID Desc"
			Dvbbs.Stats = Dvbbs.Stats & template.Strings(10)
		'内容，SQL全文索引
		Case 2
			If Dvbbs.Forum_Setting(16)<>"0" Then
				If IsSqlDataBase Then
				If Trim(searchboard)="" And Trim(searchday)="" Then
					SqlColumn = SqlColumn & stable & " T1 Inner Join ContainsTable("&stable&",body,'" & keyword & "'," & Dvbbs.Forum_Setting(11)*SearchMaxPageList & ") As T2 ON T1.AnnounceID = T2.[KEY] Order By T1.AnnounceID Desc"
				ElseIf Trim(searchboard)="" Then
					SqlColumn = SqlColumn & stable & " T1 Inner Join ContainsTable("&stable&",body,'" & keyword & "'," & Dvbbs.Forum_Setting(11)*SearchMaxPageList & ") As T2 ON T1.AnnounceID = T2.[KEY] Where  "&Replace(Replace(searchday,"and",""),"DateAndTime","T1.DateAndTime")&" Order By T1.AnnounceID Desc"
				ElseIf Trim(searchday)="" Then
					SqlColumn = SqlColumn & stable & " T1 Inner Join ContainsTable("&stable&",body,'" & keyword & "'," & Dvbbs.Forum_Setting(11)*SearchMaxPageList & ") As T2 ON T1.AnnounceID = T2.[KEY] Where "&Replace(Replace(searchboard,"and",""),"BoardID","T1.BoardID")&" Order By T1.AnnounceID Desc"
				Else
					SqlColumn = SqlColumn & stable & " T1 Inner Join ContainsTable("&stable&",body,'" & keyword & "'," & Dvbbs.Forum_Setting(11)*SearchMaxPageList & ") As T2 ON T1.AnnounceID = T2.[KEY] Where "&Replace(searchboard,"BoardID","T1.BoardID")&" "&Replace(Replace(searchday,"and",""),"DateAndTime","T1.DateAndTime")&" Order By T1.AnnounceID Desc"
				End If
				Else
				SqlColumn = SqlColumn & stable & " Where "&searchboard&" "&searchday&" body like '%"&keyword&"%' Order By AnnounceID Desc"
				End If
				Dvbbs.Stats = Dvbbs.Stats & template.Strings(11)
			Else
				 Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(19)&"&action=OtherErr"
			End If
		End Select
	'最新50贴
	Case 3
'		SqlColumn = SqlColumn & " dv_Topic Order By TopicID Desc"
		If Request("BoardID")>0 then
			SqlColumn = SqlColumn &" "&stable&" where BoardID="&trim(request("BoardID"))&" ORDER BY announceID desc"
		Else
			SqlColumn = SqlColumn &" "&stable&" ORDER BY announceID desc"
		End if
		Dvbbs.Stats = template.Strings(12)
	Case 4
		If keyword<>"" Then
			Set Rs=Dvbbs.Execute("Select UserID From Dv_User Where UserName='"&keyword&"'")
			If Rs.Eof And Rs.Bof Then
				Set Rs=Nothing
				Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(21)&"&action=OtherErr"
			Else
				SearchUserID = Rs(0)
			End If
		End If
		Dim HotTopicDay,HotTopicView,MyHotTopic
		If Dvbbs.Forum_Setting(13)<>"0" Then
			MyHotTopic = Split(Dvbbs.Forum_Setting(13),"|")
			If Ubound(MyHotTopic)=1 Then
				HotTopicDay = MyHotTopic(0)
				HotTopicView = MyHotTopic(1)
			Else
				HotTopicDay = 10
				HotTopicView = 200
			End If
		Else
			HotTopicDay = 10
			HotTopicView = 200
		End If
		Dvbbs.Stats = Replace(Replace(template.Strings(13),"{$daylimited}",HotTopicDay),"{$viewlimited}",HotTopicView)
		If IsSqlDataBase=1 Then
			searchday=" datediff(d,DateAndTime,"&SqlNowString&") < "&HotTopicDay&" and "
		Else
			searchday=" datediff('d',DateAndTime,"&SqlNowString&") < "&HotTopicDay&" and "
		End If
		If keyword<>"" Then keyword = " And PostUserID="&SearchUserID
		SqlColumn = SqlColumn & " dv_Topic Where "&searchday&" hits>"&HotTopicView&" "&keyword&" Order By TopicID Desc"
	Case 5
		If Dvbbs.UserID=0 Then
			Dvbbs.AddErrCode(61)
			Exit Function
		End If
		Dim s
		s=request("s")
		If s="" Or Not IsNumerIc(s) Then s=1
		s=clng(s)
		If s=1 Then
			SqlColumn="select top 200 BoardID,TopicID,Title,Expression,PostUserName,PostUserID,DateAndtime,IsBest,LockTopic from Dv_Topic where topicid in (select top 200 rootid from "&stable&" where ParentID>0 And PostUserID="&Dvbbs.UserID&" order by AnnounceID desc) and Boardid<>444 order by topicid desc"
			Dvbbs.Stats = template.Strings(14)
		Else
			SqlColumn="select top 200 BoardID,TopicID,Title,Expression,PostUserName,PostUserID,DateAndtime,IsBest,LockTopic from dv_topic where postUserID="&Dvbbs.UserID&" and Boardid<>444 ORDER BY topicid desc"
			Dvbbs.Stats = template.Strings(15)
		End If
	Case 6
		If keyword<>"" Then
			Set Rs=Dvbbs.Execute("Select UserID From Dv_User Where UserName='"&keyword&"'")
			If Rs.Eof And Rs.Bof Then
				Set Rs=Nothing
				Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(21)&"&action=OtherErr"
			Else
				SearchUserID = Rs(0)
			End If
		End If
		If Trim(searchboard)="" Then
		If keyword<>"" Then keyword = " Where PostUserID="&SearchUserID
		SqlColumn = "select BoardID,RootID,Title,Expression,PostUserName,PostUserID,DateAndtime,PostUserID As IsBest,PostUserID As LockTopic From dv_BestTopic "&keyword&" Order By ID Desc"
		Else
		If keyword<>"" Then keyword = " And PostUserID="&SearchUserID
		SqlColumn = "select BoardID,RootID,Title,Expression,PostUserName,PostUserID,DateAndtime,PostUserID As IsBest,PostUserID As LockTopic From dv_BestTopic Where "&Replace(searchboard,"and","")&" "&keyword&" Order By ID Desc"
		End If
		Dvbbs.Stats = template.Strings(16)
	Case Else
		Dvbbs.AddErrCode(61)
		Exit Function
	End Select

	Dvbbs.Nav()
	If Dvbbs.BoardID=0 then
		Dvbbs.Head_var 0,0,template.Strings(0),"query.asp"
	Else
		Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
	End If
	If Dvbbs.boardid>0 Then GetBoardPermission
	If IsEmpty(Session("QueryLimited")) Then
		Session("QueryLimited") = keyword & "|" & stype & "|" & Now()
	Else
		Dim QueryLimited
		QueryLimited = Split(Session("QueryLimited"),"|")
		If Ubound(QueryLimited) = 2 Then
			If Cstr(Trim(QueryLimited(0))) = Cstr(keyword) And Cstr(Trim(QueryLimited(1))) = Cstr(stype) Then
				Session("QueryLimited") = keyword & "|" & stype & "|" & Now()
			Else
				If DateDiff("s",QueryLimited(2),Now()) < Clng(Dvbbs.Forum_Setting(3)) And Not(Dvbbs.Master Or Dvbbs.BoardMaster Or Dvbbs.SuperBoardMaster) Then
					Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(20),"{$timelimited}",Dvbbs.Forum_Setting(3))&"&action=OtherErr"
				Else
					Session("QueryLimited") = keyword & "|" & stype & "|" & Now()
				End If
			End If
		Else
			Session("QueryLimited") = keyword & "|" & stype & "|" & Now()
		End If
	End If
	'Response.Write Session("QueryLimited")
End Function

Function SearchResult()
	Dim FirstNum,TotalRec,ThisPageNum
	Dim Rs,i,TempData,Board_Data
	Dim TempStr,TempStr1,TempStr2,TempStr3
	TempStr = template.html(1)
	TempStr1 = template.html(2)
	TotalRec = -1
	If page < 1 Then
		FirstNum = 0
	Else
		FirstNum = page * Dvbbs.Forum_Setting(11) + 1
	End If
	SQLQueryStr()
	If Dvbbs.ErrCodes<>"" Then Exit Function
	If Not IsObject(Conn) Then ConnectionDatabase
	'Response.Write SqlColumn
	'On Error Resume Next
	'Set Rs=Dvbbs.Execute(SqlColumn)
	Dvbbs.SqlQueryNum = Dvbbs.SqlQueryNum + 1
	Set Rs=server.createobject("adodb.recordset")
	Rs.Open SqlColumn,Conn,1,1
	If Err Then
		Dvbbs.AddErrCode(61)
		Exit Function
	End If
	If Not (Rs.Eof And Rs.Bof) Then
		If TotalRec = -1 Then
			TotalRec = Rs.RecordCount
			If TotalRec = -1 Then
				For i = 1 to FirstNum
					If Not Rs.Eof Then
						Rs.MoveNext
					Else
						Exit For
					End If
				Next
			Else
				If FirstNum > TotalRec Then FirstNum = TotalRec - Cint(Dvbbs.Forum_Setting(11))
				If FirstNum > 0 and FirstNum <= TotalRec Then Rs.absoluteposition = FirstNum
			End If
		Else
			If FirstNum > TotalRec Then FirstNum = TotalRec - Cint(Dvbbs.Forum_Setting(11))
			For i = 1 to FirstNum
				Rs.MoveNext
			Next
		End If
		If Not Rs.Eof Then
			TempData = Rs.GetRows(Cint(Dvbbs.Forum_Setting(11))+1)
			ThisPageNum = Ubound(TempData,2)+1
		Else
			TotalRec = 0
			ThisPageNum = 0
		End If
	Else
		TotalRec = 0
		ThisPageNum = 0
	End If
	If ThisPageNum > TotalRec Then TotalRec = ThisPageNum
	Set Rs=Nothing
	If TotalRec = 0 Then TempStr = Replace(TempStr,"{$searchresultloop}",template.html(3))
	'BoardID=0,RootID=1,Topic=2,Expression=3,UserName=4,PostUserID=5,DateAndTime=6,IsBest=7,LockTopic=8,Body=9,AnnounceID=10
	Dim TopicStats
	If Dvbbs.BoardID>0 Then TempStr1 = Replace(TempStr1,"{$boardtype}",Dvbbs.BoardType)
	TopicStats = Dvbbs.mainpic(2)
	For i = 0 To ThisPageNum-1 Step 1
		TempStr2 = TempStr1
		If TempData(8,i)=1 Then TopicStats=Dvbbs.mainpic(4)
		If TempData(7,i)=1 Or stype=6 Then TopicStats=Dvbbs.mainpic(5)
		TempStr2 = Replace(TempStr2,"{$statpic}",TopicStats)
		TopicStats = Dvbbs.mainpic(2)
		TempStr2 = Replace(TempStr2,"{$boardid}",TempData(0,i))
		TempStr2 = Replace(TempStr2,"{$userid}",TempData(5,i))
		TempStr2 = Replace(TempStr2,"{$dateandtime}",TempData(6,i))
		TempStr2 = Replace(TempStr2,"{$username}",Dvbbs.HtmlEncode(TempData(4,i)))
		TempStr2 = Replace(TempStr2,"{$expression}",TempData(3,i)&"")
		If InStr("," & hidboardid & ",","," & TempData(0,i) & ",") > 0 Then
			TempStr2 = Replace(TempStr2,"{$topic}","隐含版面帖子，请点击链接浏览。")
		End If
		If InStr("," & FobBoardID & ",","," & TempData(0,i) & ",") > 0 Then
			TempStr2 = Replace(TempStr2,"{$topic}","认证版面帖子，请点击链接浏览。")
		End If
		If (TempData(0,i)=444 Or TempData(0,i)=777) Then
			TempStr2 = Replace(TempStr2,"{$topic}","帖子已被删除或者在认证中")
		End If
		'If Dvbbs.BoardID=0 Then
			If TempData(0,i)=444 Or TempData(0,i)=777 Then
				TempStr2 = Replace(TempStr2,"{$boardtype}","回收站")
			Else
				If Not InStr((","&Dvbbs.cachedata(27,0)&","),(","&TempData(0,i)&","))>0 Then
					 TempStr2 = Replace(TempStr2,"{$boardtype}","在错乱的版面")
				Else
					
					Dvbbs.Name="BoardInfo_" & TempData(0,i)
  					If Dvbbs.ObjIsEmpty() Then Dvbbs.ReloadBoardInfo(TempData(0,i))
					Board_Data=Dvbbs.Value
					TempStr2 = Replace(TempStr2,"{$boardtype}",Board_Data(1,0))
				End If 
			End If		
		'End If
		If InStr(SqlColumn,"Body")>0 Then
			TempStr2 = Replace(TempStr2,"{$linkinfo}","&ID=" & TempData(1,i) & "&replyID=" & TempData(10,i) & "&skin=1")
			If Trim(TempData(2,i))="" Then
				TempStr2 = Replace(TempStr2,"{$topic}","Re:"&cutStr(Replace(Replace(reUBBCode(TempData(9,i)),chr(10),""),chr(13),""),35))
			Else
				TempStr2 = Replace(TempStr2,"{$topic}",cutStr(TempData(2,i),35))
			End If
		Else
			TempStr2 = Replace(TempStr2,"{$topic}",cutStr(TempData(2,i),35))
			TempStr2 = Replace(TempStr2,"{$linkinfo}","&ID=" & TempData(1,i))
		End If
		TempStr3 = TempStr3 & TempStr2
	Next
	TempStr = Replace(TempStr,"{$searchresultloop}",TempStr3)
	If TotalRec > 0 Then
	Dim SearchStr,Temp,TempPage
	SearchStr="stype="&stype&"&pSearch="&pSearch&"&nSearch="&nSearch&"&keyword="&keyword&"&SearchDate="&request("SearchDate")&"&BoardID="&Dvbbs.BoardID&"&stable="&stable

	TempStr3 = Split(template.html(4),"||")
	If TotalRec > Cint(Dvbbs.Forum_Setting(11)) Then
		Temp = Page + 2
	Else
		Temp = Page + 1
	End If
	If Temp >= SearchMaxPageList Then Temp = SearchMaxPageList
	For i = 1 to Temp
		If i = page+1 Then
			TempPage = TempPage & Replace(TempStr3(1),"{$nowpage}",i)
		Else
			TempStr1 = Replace(TempStr3(2),"{$rpage}",i-1)
			TempStr1 = Replace(TempStr1,"{$rnpage}",i)
			TempStr1 = Replace(TempStr1,"{$s}",Request("s"))
			TempPage = TempPage & TempStr1
		End If
	Next
	TempStr2 = Replace(TempStr3(0),"{$pagenum}",TempPage)
	TempStr2 = Replace(TempStr2,"{$ThisPageNum}",ThisPageNum)
	TempStr2 = Replace(TempStr2,"{$pagelistnum}",Dvbbs.Forum_Setting(11))
	TempStr2 = Replace(TempStr2,"{$SearchStr}",SearchStr)
	TempStr2 = Replace(TempStr2,"{$alertcolor}",Dvbbs.mainsetting(1))
	TempStr = Replace(TempStr,"{$pagelist}",TempStr2)
	End If
	TempStr = Replace(TempStr,"{$pagelist}","")
	Response.Write TempStr
End Function
Function reUBBCode(strContent)
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	strContent=replace(strContent,"&nbsp;"," ")
	re.Pattern="(\[QUOTE\])(.|\n)*(\[\/QUOTE\])"
	strContent=re.Replace(strContent,"$2")
	re.Pattern="(\[point=*([0-9]*)\])(.|\n)*(\[\/point\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[post=*([0-9]*)\])(.|\n)*(\[\/post\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[power=*([0-9]*)\])(.|\n)*(\[\/power\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[usercp=*([0-9]*)\])(.|\n)*(\[\/usercp\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[money=*([0-9]*)\])(.|\n)*(\[\/money\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[replyview\])(.|\n)*(\[\/replyview\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[usemoney=*([0-9]*)\])(.|\n)*(\[\/usemoney\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="\[username=(.[^\[]*)\](.[^\[]*)\[\/username\]"
	strContent=re.Replace(strContent,"&nbsp;")
	strContent=replace(strContent,"<I></I>","")
	set re=Nothing
	reUBBCode=strContent
End Function
'截取指定字符
Function cutStr(str,strlen)
	'去掉所有HTML标记
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="<(.[^>]*)>"
	str=re.Replace(str,"")
	set re=Nothing
	str = Dvbbs.HTMLEncode(str)
	Dim l,t,c,i
	l=Len(str)
	t=0
	For i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		If c>255 Then
			t=t+2
		Else
			t=t+1
		End If
		If t>=strlen Then
			cutStr=left(str,i)&"..."
			Exit For
		Else
			cutStr=str
		End If
	Next
End Function
%>
