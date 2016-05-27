<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/ubblist.asp"-->
<!--#include file="inc/Email.asp"-->
<%
Dim MyPost
Dim postbuyuser,bgcolor,abgcolor
Dvbbs.Loadtemplates("post")
Set MyPost = New Dvbbs_Post
Dvbbs.Stats = MyPost.ActionName
Dvbbs.Nav()
Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
MyPost.Save_CheckData
Set MyPost = Nothing
Dvbbs.ActiveOnline
Dvbbs.Footer

Class Dvbbs_Post
	Public Action,ActionName,Star,Page,IsAudit,TotalUseTable,ToAction,TopicMode,Reuser
	Private AnnounceID,ReplyID,ParentID,RootID,Topic,Content,char_changed,signflag,mailflag,iLayer,iOrders
	Private TopTopic,IsTop,LastPost,LastPost_1,UpLoadPic_n,ihaveupfile,smsuserlist,upfileinfo
	Private UserName,UserPassWord,UserPost,GroupID,UserClass,DateAndTime,DateTimeStr,Expression,MyLastPostTime,LastPostTimes
	Private LockTopic,MyLockTopic,MyIsTop,MyIsTopAll,MyTopicMode,Child
	Private CanLockTopic,CanTopTopic,CanTopTopic_a,CanEditPost,Rs,SQL,i,IsAuditcheck
	Private vote,votetype,votenum,votetimeout,voteid,isvote
	Private Sub Class_Initialize()
		If Dvbbs.Board_Setting(0)="1" And Not Dvbbs.Master Then
			Response.redirect "showerr.asp?action=lock&boardid="&dvbbs.boardID&"" 
		End If
		If Dvbbs.IsReadonly()  And Not Dvbbs.Master Then Response.redirect "showerr.asp?action=readonly&boardid="&dvbbs.boardID&"" 
		Action = Request("Action")
		TotalUseTable = Dvbbs.NowUseBBS
		Select Case Action
		Case "snew"
			Action = 5
			ActionName = template.Strings(1)
			If Dvbbs.GroupSetting(3)="0" Then Dvbbs.AddErrCode(70)
		Case "sre"
			Action = 6
			ActionName = template.Strings(3)
			If Dvbbs.GroupSetting(5)="0" then Dvbbs.AddErrCode(71)
		Case "svote"
			Action = 7
			ActionName = template.Strings(5)
			If Dvbbs.GroupSetting(8)="0" then Dvbbs.AddErrCode(56)
		Case "sedit"
			Action = 8
			ActionName = template.Strings(7)
		Case Else
			Action = 1
			ActionName = template.Strings(0)
		End Select
		Star = Request("star")
		If Star = "" Or Not IsNumeric(Star) Then Star = 1
		Star = Clng(Star)
		Page = Request("page")
		If Page = "" Or Not IsNumeric(Page) Then Page = 1
		Page = Clng(Page)
		IsAudit = Cint(Dvbbs.Board_Setting(3))
		Reuser=False'此变量标识是否更名发贴
	End Sub

	'通用判断
	Public Function Chk_Post()
		If Dvbbs.Board_Setting(43)="1" Then Dvbbs.AddErrCode(72)
		If Dvbbs.Board_Setting(1)="1" and Dvbbs.GroupSetting(37)="0" Then Dvbbs.AddErrCode(26)
		If Dvbbs.UserID>0 Then
			If Clng(Dvbbs.GroupSetting(52))>0 And  DateDiff("s",Dvbbs.MyUserInfo(14),Now)<Clng(Dvbbs.GroupSetting(52))*60 Then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(21),"{$timelimited}",Dvbbs.GroupSetting(52))&"&action=OtherErr"
			If Dvbbs.GroupSetting(62)<>"0" And Not Action = 8 Then
				If Clng(Dvbbs.GroupSetting(62))<=Clng(Dvbbs.UserToday(0)) Then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(27),"{$topiclimited}",Dvbbs.GroupSetting(62))&"&action=OtherErr"
			End If
		End If
		If Dvbbs.GroupSetting(3)="0" And (Action = 5 Or Action = 7) Then Response.redirect "showerr.asp?ErrCodes=<li>您没有发表新主题的权限&action=OtherErr"
		If Dvbbs.GroupSetting(5)="0" And (Action = 6) Then Response.redirect "showerr.asp?ErrCodes=<li>您没有回复别人贴子的权限&action=OtherErr"
	End Function

	'判断用户是否有编辑权限且提取相关信息
	Public Function Get_Edit_PermissionInfo()
		Dim old_user
		If Action = 4 Then
		Set Rs=Dvbbs.Execute("select b.username,b.topic,b.body,b.dateandtime,u.UserGroupID,b.signflag,b.emailflag from "&TotalUseTable&" b,[dv_user] u where b.postuserid=u.userid and b.RootID="&AnnounceID&" and b.AnnounceID="&replyID)
		Else
		Set Rs=Dvbbs.Execute("select b.username,b.topic,b.body,b.dateandtime,u.UserGroupID,b.signflag,b.emailflag from "&TotalUseTable&" b,[dv_user] u where b.postuserid=u.userid and b.RootID="&RootID&" and b.AnnounceID="&AnnounceID)
		End If
		If Rs.Eof And Rs.Bof Then
			Dvbbs.AddErrCode(48)
		Else
			If Action = 4 Then
				signflag=Rs("signflag")
				mailflag=Rs("emailflag")
				Topic=rs("topic")
				If Topic<>"" Then Topic = Server.HtmlEncode(Topic)
				Content=rs("body")
				old_user=rs("username")
			Else
				If Clng(Dvbbs.forum_setting(50))>0 then
					If Datediff("s",rs("dateandtime"),Now())>Clng(Dvbbs.forum_setting(50))*60 then
						Content = Content+chr(13)+chr(10)+char_changed+chr(13)
					End If
				Else
					Content = Content+chr(13)+chr(10)+char_changed+chr(13)
				End If
			End If
			If Clng(Dvbbs.forum_setting(51))>0 and not (Dvbbs.master or Dvbbs.boardmaster or Dvbbs.superboardmaster) Then 
				If DateDiff("s",rs("dateandtime"),Now())>Clng(Dvbbs.forum_setting(51))*60 then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(Replace(template.Strings(22),"{$posttime}",Datediff("s",rs("dateandtime"),Now())/60),"{$etlimited}",Dvbbs.forum_setting(51))&"&action=OtherErr"
			End If 
			If Rs("username")=Dvbbs.membername Then 
				If Dvbbs.GroupSetting(10)="0" then
					Dvbbs.AddErrCode(74)
					CanEditPost=False
				Else 
					CanEditPost=True
				End If 
			Else 
				If (Dvbbs.master or Dvbbs.superboardmaster or Dvbbs.boardmaster) and Dvbbs.GroupSetting(23)="1" then
					CanEditPost=True
				Else 
					CanEditPost=False
				End  If 
				If Cint(Dvbbs.UserGroupID) > 3 And Dvbbs.GroupSetting(23)="1" Then CanEditPost=true
				If Dvbbs.GroupSetting(23)="1" and Dvbbs.founduserPer Then 
					CanEditPost=True
				ElseIf Dvbbs.GroupSetting(23)="0" And Dvbbs.founduserPer Then 
					CanEditPost=False
				End If
				If Cint(Dvbbs.UserGroupID) < 4 And Cint(Dvbbs.UserGroupID) = rs("UserGroupID") Then 
					Dvbbs.AddErrCode(75)
				ElseIf Cint(Dvbbs.UserGroupID) < 4 and Cint(Dvbbs.UserGroupID) > rs("UserGroupID") Then
					Dvbbs.AddErrCode(76)
				End If 
				If Not CanEditPost Then Dvbbs.AddErrCode(77)
			End If
		End If
		Set Rs=Nothing
		Dvbbs.ShowErr()
		If Action = 4 Then Dvbbs.MemberName=old_user
	End Function

	'返回判断和参数
	Public Function Get_M_Request()
		AnnounceID = Request("ID")
		If AnnounceID = "" Or Not IsNumeric(AnnounceID) Then Dvbbs.AddErrCode(30)
		Dvbbs.ShowErr()
		AnnounceID = Clng(AnnounceID)
	End Function

	Rem ------------------------
	Rem 保存部分函数开始
	Rem ------------------------
	'检查数据,提取数据，获得贴子数据表名等。
	Public Sub Save_CheckData()
		Chk_Post()
		CheckfromScript()
		Dim mysessiondata
		Content=Dvbbs.Checkstr(Request.Form("body"))
		'把提交的数据保存到session
		mysessiondata=Session(Dvbbs.CacheName & "UserID")
		mysessiondata(37)=Content
		Session(Dvbbs.CacheName & "UserID")=mysessiondata
		If Dvbbs.Board_Setting(4)="1" Then
			If Not Dvbbs.CodeIsTrue() Then
				 Response.redirect "showerr.asp?ErrCodes=<li>验证码校验失败，2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
			End If
		End If	
		Expression=Dvbbs.Checkstr(Request.Form("Expression"))
		If Expression="" Then Expression="face1.gif"
		Topic=Dvbbs.Checkstr(Trim(Request.Form("topic")))
		'Content=Dvbbs.Checkstr(Request.Form("Content"))
		signflag=Dvbbs.Checkstr(Trim(Request.Form("signflag")))
		mailflag=Dvbbs.Checkstr(Trim(Request.Form("emailflag")))
		MyTopicMode=Dvbbs.Checkstr(Trim(Request.Form("topicximoo")))
		MyLockTopic=Dvbbs.Checkstr(Trim(Request.Form("locktopic")))
		Myistop=Dvbbs.Checkstr(Trim(Request.Form("istop")))
		Myistopall=Dvbbs.Checkstr(Trim(Request.Form("istopall")))
		TopicMode=Request.Form("topicmode")
		If Dvbbs.strLength(topic)> CLng(Dvbbs.Board_Setting(45)) Then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(23),"{$topiclimited}",Dvbbs.Board_Setting(45))&"<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
		If Dvbbs.strLength(Content) > CLng(Dvbbs.Board_Setting(16)) Then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(24),"{$bodylimited}",Dvbbs.Board_Setting(16))&"<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
		REM 2004-4-23添加限制帖子内容最小字节数,下次在模板中添加。Dvbbs.YangZheng
		If Dvbbs.strLength(Content) < CLng(Dvbbs.Board_Setting(52)) And Not CLng(Dvbbs.Board_Setting(52)) = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(24),"大于{$bodylimited}","小于"&Dvbbs.Board_Setting(52))&"<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
		Dim testContent
		testContent=Content
		testContent=Html2Ubb(testContent)
		testContent=Replace(testContent,vbNewLine,"")
		testContent=Replace(testContent," ","")
		testContent=Replace(testContent,"&nbsp;","")		
		If testContent="" Then Response.redirect "showerr.asp?ErrCodes=<li>您没有填写内容,或因当前不支持HTML模式,内容被自动过滤<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
		If Not IsNumeric(mailflag) Or mailflag="" Then mailflag=0
		If TopicMode<>"" and IsNumeric(TopicMode) Then TopicMode=Cint(TopicMode) Else TopicMode=0
		mailflag=CInt(mailflag)
		If signflag="yes" Then
			signflag=1
		Else
			signflag=0
		End If
		If Request.form("upfilerename")<>"" Then
			ihaveupfile=1
			upfileinfo=Replace(Request.form("upfilerename"),"'","")
			upfileinfo=Replace(upfileinfo,";","")
			upfileinfo=Replace(upfileinfo,"--","")
			upfileinfo=Replace(upfileinfo,")","")
			Dim fixid,upfilelen
			fixid=Replace(upfileinfo," ","")
			fixid=Replace(fixid,",","")
			If Not IsNumeric(fixid) Then ihaveupfile=0
			upfilelen=len(upfileinfo)
			upfileinfo=left(upfileinfo,upfilelen-1)
		Else
			ihaveupfile=0
		End If
		voteid=0
		isvote=0
		If Action = 7 Then
			votetype=Dvbbs.Checkstr(request.Form("votetype"))
			If IsNumeric(votetype)=0 or votetype="" Then votetype=0
			vote=Dvbbs.Checkstr(trim(Replace(request.Form("vote"),"|","")))
			Dim j,k,vote_1,votelen,votenumlen
			If vote="" Then
				Dvbbs.AddErrCode(81)
			Else
				vote=split(vote,chr(13)&chr(10))
				j=0
				For i = 0 To ubound(vote)
					If Not (vote(i)="" Or vote(i)=" ") Then
						vote_1=""&vote_1&""&vote(i)&"|"
						j=j+1
					End If
					If i>cint(Dvbbs.Board_Setting(32))-2 Then Exit For
				Next
				For k = 1 to j
					votenum=""&votenum&"0|"
				Next
				votelen=len(vote_1)
				votenumlen=len(votenum)
				votenum=left(votenum,votenumlen-1)
				vote=left(vote_1,votelen-1)
			End If
			If Not IsNumeric(request("votetimeout")) Then
				Dvbbs.AddErrCode(82)
			Else
				If request("votetimeout")="0" Then
					votetimeout=dateadd("d",9999,Now())
				Else
					votetimeout=dateadd("d",request("votetimeout"),Now())
				End If
				votetimeout=Replace(Replace(CSTR(votetimeout+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
			End If
		End If
		If Action = 5 Or Action = 7 Then
			CanLockTopic=False
			CanTopTopic=False
			CanTopTopic_a=False
			If Topic="" Then Response.redirect "showerr.asp?ErrCodes=<li>您忘记填写标题<BR>2秒后自动返回上一页面。&action=OtherErr&autoreload=1"
			'减少判断，如果不为固顶，锁定等等的操作的话，不检查权限。
			If MyLockTopic="yes" Or Myistopall="yes" Or Myistop="yes" Then
				'判断用户是否有固顶/解除固顶帖子权限
				If (dvbbs.master or dvbbs.superboardmaster or dvbbs.boardmaster) Then 
					If Dvbbs.GroupSetting(21)="1" Then CanTopTopic=True
					If  Dvbbs.GroupSetting(20)="1" Then CanLockTopic=True
				End If
				If Dvbbs.GroupSetting(21)="1" and Cint(Dvbbs.UserGroupID)>3 Then CanTopTopic=True
				If Dvbbs.FoundUserPer and Dvbbs.GroupSetting(21)="1" Then
					CanTopTopic=True
				ElseIf Dvbbs.FoundUserPer and Dvbbs.GroupSetting(21)="0" Then
					CanTopTopic=False
				End If
				'判断用户是否有总固顶帖子权限
				If (dvbbs.master or dvbbs.superboardmaster or dvbbs.boardmaster) and Dvbbs.GroupSetting(38)="1" Then CanTopTopic_a=True
				If Dvbbs.GroupSetting(38)="1" and Cint(Dvbbs.UserGroupID)>3 Then CanTopTopic_a=True
				If Dvbbs.FoundUserPer and Dvbbs.GroupSetting(38)="1" Then
					CanTopTopic_a=True
				ElseIf Dvbbs.FoundUserPer and Dvbbs.GroupSetting(38)="0" Then
					CanTopTopic_a=False
				End If	
			End If
			If MyLockTopic="yes" Then
				MyLockTopic=1
			Else
				MyLockTopic=0
			End If
			If Myistopall="yes" Then
				Myistopall=1
			Else
				Myistopall=0
			End If
			If Not CanTopTopic_a Then Myistopall=0
			If Myistop="yes" and Myistopall=0 Then
				Myistop=1
				If Not CanTopTopic Then Myistop=0
			ElseIf Myistopall=1 Then
				Myistop=3
			Else
				Myistop=0
			End If
			If Not IsNumeric(MyTopicMode) or Dvbbs.GroupSetting(51)="0" Then MyTopicMode=0
			If Not CanLockTopic Then MyLockTopic=0
			TotalUseTable=Dvbbs.NowUseBbs
		ElseIf Action = 6 Then
			AnnounceID = request("followup")
			If AnnounceID = "" Or Not IsNumeric(AnnounceID) Then Dvbbs.AddErrCode(30)
			Dvbbs.ShowErr()
			AnnounceID = Clng(AnnounceID)
			ParentID = AnnounceID
			RootID = request("RootID")
			If RootID = "" Or Not IsNumeric(RootID) Then Dvbbs.AddErrCode(30)
			Dvbbs.ShowErr()
			RootID = Clng(RootID)
			TotalUseTable=Request.Form("TotalUseTable")
			TotalUseTable=checktable(TotalUseTable)
			MyLockTopic=0
		ElseIf Action = 8 Then
			If Not IsNumeric(MyTopicMode) or Dvbbs.GroupSetting(51)="0" Then MyTopicMode=0
			AnnounceID = request("replyID")
			If AnnounceID = "" Or Not IsNumeric(AnnounceID) Then Dvbbs.AddErrCode(30)
			Dvbbs.ShowErr()
			AnnounceID = Clng(AnnounceID)
			RootID = request("ID")
			If RootID = "" Or Not IsNumeric(RootID) Then Dvbbs.AddErrCode(30)
			Dvbbs.ShowErr()
			RootID = Clng(RootID)
			TotalUseTable=Request.Form("TotalUseTable")
			TotalUseTable=checktable(TotalUseTable)
			MyLockTopic=0
			Dvbbs.ShowErr()
			Set Rs=Dvbbs.Execute("select AnnounceID from "&TotalUseTable&" where ParentID=0 and rootid="&RootID)
			If Not Rs.eof Then 
				If AnnounceID=Rs(0) Then
					If Topic="" Then Dvbbs.AddErrCode(79)
				End If
			Else
				Dvbbs.AddErrCode(30)			
			End If
		End If
		Dvbbs.Showerr()
		SaveData()
		'清空表单内容
		mysessiondata=Session(Dvbbs.CacheName & "UserID")
		mysessiondata(37)=""
		Session(Dvbbs.CacheName & "UserID")=mysessiondata
	End Sub
	'保存数据
	Private Sub SaveData()
		If Not (Dvbbs.Master or Action = 8) Then CheckpostTime()
		Dim Forumupload
 		If Dvbbs.GroupSetting(64)="1" Then
 			IsAudit=0
 		Else
 			If Dvbbs.Board_Setting(57)="1" Then
 				IsAudit=NeedIsAudit()
 				IsAuditcheck=IsAudit
 			End If
 		End If
 		locktopic=0
		If MyLockTopic=1 Then locktopic=1
		If IsAudit=1 And Action <> 8 Then
			LockTopic=Dvbbs.BoardID
			Dvbbs.BoardID=777
			Response.Cookies("Dvbbs")=LockTopic
		Else
			Response.Cookies("Dvbbs")=Dvbbs.Boardid
		End If
		Forumupload=split(Dvbbs.Board_Setting(19),"|")
		For i=0 to ubound(Forumupload)
			If Instr(Content,"[upload="&Forumupload(i)&"]") or Instr(Content,"."&Forumupload(i)&"") or Instr(Content,"["&Forumupload(i)&"]") then
				uploadpic_n=Forumupload(i)
				Exit For
			End If
		Next
		If InStr(Content,"viewfile.asp?ID=") Then uploadpic_n="down"
		If Not Action = 8 Then 
			savepost()
			updatepostuser()
		Else
			Update_Edit_Announce()
		End If
		succeed()
	End Sub
	'保存发贴，投票和回贴
	Public Sub savepost()
		If Action = 5 Or Action = 7 Then
			ilayer=1:iOrders=0:ParentID=0
			If Myistop=3 Then
				MyLastPostTime=DateADD("d",300,Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午",""))
			ElseIf Myistop=1 Then
				MyLastPostTime=DateADD("d",100,Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午",""))
			Else
				MyLastPostTime=Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
			End If
			DateTimeStr=Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
			If Action = 7 Then Insert_To_Vote()
			Insert_To_Topic()
			'更新总固顶和固顶的数据以及缓存数据
			If MyIsTop=3 Then
				'将总固顶ID插入总设置表
				Dim iForum_AllTopNum
				Set Rs=Dvbbs.Execute("Select Forum_AllTopNum From Dv_Setup")
				If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
					iForum_AllTopNum = RootID
				Else
					iForum_AllTopNum = Rs(0) & "," & RootID
				End If
				Dvbbs.Execute("Update Dv_Setup Set Forum_AllTopNum='"&iForum_AllTopNum&"'")
				Dvbbs.ReloadSetupCache iForum_AllTopNum,28
				Set Rs=Nothing
			ElseIf MyIsTop=1 Then
				Dim BoardTopStr
				Set Rs=Dvbbs.Execute("Select BoardID,BoardTopStr From Dv_Board Where BoardID="&Clng(Dvbbs.BoardID))
				If Not (Rs.Eof And Rs.Bof) Then
					If Rs(1)="" Or IsNull(Rs(1)) Then
						BoardTopStr = RootID
					Else
						If InStr(","&Rs(1)&",","," & RootID & ",")>0 Then
							BoardTopStr = Rs(1)
						Else
							BoardTopStr = Rs(1) & "," & RootID
						End If
					End If
					Dvbbs.Execute("Update Dv_Board Set BoardTopStr='"&BoardTopStr&"' Where BoardID="&Rs(0))
					Dvbbs.ReloadBoardInfo(Rs(0))
				End If
				Set Rs=Nothing
			End If
		ElseIf Action = 6 Then
			Get_SaveRe_TopicInfo()

			Get_ForumTreeCode()

			DateTimeStr=Replace(Replace(CSTR(NOW()+Dvbbs.Forum_Setting(0)/24),"上午",""),"下午","")
		End If
		Insert_To_Announce()
		If Action = 6 Then
			topic=Replace(Replace(cutStr(topic,14),chr(10),""),"'","")
			If topic="" Then
				topic=Content
				topic=Replace(cutStr(topic,14),chr(10),"")
			Else
				topic=Replace(cutStr(topic,14),chr(10),"")
			End If
			If ihaveupfile=1 Then Dvbbs.Execute("update dv_upfile set F_AnnounceID='"&RootID&"|"&AnnounceID&"',F_Readme='"&Replace(toptopic,"'","")&"',F_flag=0 where F_ID in ("&upfileinfo&")")
		Else
			If ihaveupfile=1 Then Dvbbs.Execute("update dv_upfile set F_AnnounceID='"&RootID&"|"&AnnounceID&"',F_Readme='"&Topic&"',F_flag=0  where F_ID in ("&upfileinfo&")")
		End If		
		LastPost=Replace(username,"$","") & "$" & AnnounceID & "$" & DateTimeStr & "$" & Replace(cutStr(topic,20),"$","&#36;") & "$" & uploadpic_n & "$" & Dvbbs.UserID & "$" & RootID & "$" & Dvbbs.BoardID
		LastPost=reubbcode(Replace(LastPost,"'",""))
		LastPost=Dvbbs.ChkBadWords(LastPost)
		If IsAudit<>1 Then
			If Action = 6 Then 
				If istop=3 Then
					If IsSqlDataBase=1 Then
						SQL="update Dv_topic set child=child+1,LastPostTime=dateadd(day,300,"&SqlNowString&"),LastPost='"&LastPost&"' where TopicID="&RootID
					Else
						SQL="update Dv_topic set child=child+1,LastPostTime=dateadd('d',300,"&SqlNowString&"),LastPost='"&LastPost&"' where TopicID="&RootID
					End If
				ElseIf istop=1 Then
					If IsSqlDataBase=1 Then
						SQL="update Dv_topic set child=child+1,LastPostTime=dateadd(day,100,"&SqlNowString&"),LastPost='"&LastPost&"' where TopicID="&RootID
					Else
						SQL="update Dv_topic set child=child+1,LastPostTime=dateadd('d',100,"&SqlNowString&"),LastPost='"&LastPost&"' where TopicID="&RootID
					End If		
				Else
					SQL="update Dv_topic set child=child+1,LastPostTime="&SqlNowString&",LastPost='"&LastPost&"' where TopicID="&RootID
				End If
				Dvbbs.Execute(SQL)
				Child=Child+2
				If Child mod Dvbbs.Board_Setting(27)=0 Then 
					star=Child\Dvbbs.Board_Setting(27)
				Else
					star=(Child\Dvbbs.Board_Setting(27))+1
				End If
				Get_Chan_TopicOrder()
			Else
				Dvbbs.Execute("update Dv_topic set LastPost='"&LastPost&"' where topicid="&RootID)
			End If	
		End If
		If Action = 5 Or Action = 7 Then
			toptopic=Replace(topic,"$","&#36;")
		Else
			toptopic=Replace(cutStr(toptopic,20),"$","&#36;")
		End If
		'toptopic =主标题,Content=内容
		'判断是否是加密的论坛，如果是，则不显示最后发贴内容。
		If Dvbbs.Board_Setting(2)="1" Then
			LastPost_1="保密$" & AnnounceID & "$" & DateTimeStr & "$请认证用户进入查看$" & uploadpic_n & "$" & Dvbbs.UserID & "$" & RootID & "$" & Dvbbs.BoardID
		Else
			LastPost_1=Replace(username,"$","") & "$" & AnnounceID & "$" & DateTimeStr & "$" & toptopic & "$" & uploadpic_n & "$" & Dvbbs.UserID & "$" & RootID & "$" & Dvbbs.BoardID
		End If
		LastPost_1=reubbcode(Replace(LastPost_1,"'",""))
		LastPost_1=Dvbbs.ChkBadWords(LastPost_1)
		If IsAudit=0 Then
			UpDate_BoardInfoAndCache()
			UpDate_ForumInfoAndCache()
		End If
		Response.Cookies("Dvbbs")=Dvbbs.BoardID
	End Sub
	Public Sub Get_SaveRe_TopicInfo()
		SQL="select locktopic,LastPost,title,smsuserlist,IsSmsTopic,istop,Child from Dv_topic where BoardID="&Dvbbs.BoardID&" And topicid="&cstr(RootID)
		Set Rs=dvbbs.Execute(sql)
		If Not Rs.EOF And Not Rs.BOF Then
			toptopic=rs(2)
			istop=rs(5)
			Child=Rs("Child")
			If rs("IsSmsTopic")=1 Then smsuserlist=rs("smsuserlist")
			If Rs("LockTopic")=1 And Not (Dvbbs.Master Or Dvbbs.BoardMaster Or Dvbbs.SuperBoardMaster) Then Dvbbs.AddErrCode(78)
			If Not IsNull(rs(1)) Then
				LastPost=split(rs(1),"$")
				If ubound(LastPost)=7 Then
					UpLoadPic_n=LastPost(4)
				Else
					UpLoadPic_n=""
				End If
			End If
		End If
		Dvbbs.showErr()
		Set Rs = Nothing
	End Sub
	Public Sub Insert_To_Vote()
		'插入投票记录 GroupSetting(68)投票项目中是否可以使用HTML
		If Dvbbs.GroupSetting(68)<>"1" Then vote=server.htmlencode(vote)
		dvbbs.execute("insert into dv_vote (vote,votenum,votetype,timeout) values ('"&vote&"','"&votenum&"',"&votetype&",'"&votetimeout&"')")
		set rs=dvbbs.execute("select Max(voteid) from dv_vote")
		voteid=Rs(0)
		isvote=1
		Set Rs=Nothing
	End Sub
	Public Sub Insert_To_Topic()
		'插入主题表
		SQL="insert into Dv_topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,PostTable,locktopic,istop,TopicMode,isvote,PollID,Mode) values ('"&topic&"',"&Dvbbs.boardid&",'"&username&"',"&Dvbbs.userid&",'"&DateTimeStr&"','"&Expression&"','$$"&DateTimeStr&"$$$$','"&MyLastPostTime&"','"&TotalUseTable&"',"&locktopic&","&Myistop&","&MyTopicMode&","&isvote&","&voteid&","&TopicMode&")" 
		Dvbbs.Execute(sql)
		Set Rs=Dvbbs.Execute("select Max(topicid) from Dv_topic Where PostUserid="&Dvbbs.UserID)
		RootID=rs(0)
	End Sub
	Public Sub Insert_To_Announce()
		'插入回复表
		DIM UbblistBody
		UbblistBody = Content
		Content = Html2Ubb(Content)
		UbblistBody = Ubblist(Content)
		SQL="insert into "&TotalUseTable&"(Boardid,ParentID,username,topic,body,DateAndTime,length,RootID,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isupload,IsAudit,Ubblist) values ("&Dvbbs.boardid&","&ParentID&",'"&username&"','"&topic&"','"&Content&"','"&DateTimeStr&"','"&Dvbbs.strlength(Content)&"',"&RootID&","&ilayer&","&iorders&",'"&Dvbbs.UserTrueIP&"','"&Expression&"',"&locktopic&","&signflag&","&mailflag&",0,"&Dvbbs.userid&","&ihaveupfile&","&IsAudit&",'"&UbblistBody&"')"
		Dvbbs.Execute(sql)
		Set rs=Dvbbs.Execute("select Max(AnnounceID) from "&TotalUseTable&" Where PostUserID="&Dvbbs.UserID)
		AnnounceID=rs(0)
	End Sub
	'检查贴中是否含过滤字
	Public Function NeedIsAudit()
		NeedIsAudit=0
		Dim i,ChecKData
		If Dvbbs.Board_Setting(58)<>"0" Then
			ChecKData=split(Dvbbs.Board_Setting(58),"|")
			For i=0 to UBound(ChecKData)
				If Trim(ChecKData(i))<>"" Then
					If InStr(Content,ChecKData(i))>0 Or InStr(Topic,ChecKData(i))>0 Then
						NeedIsAudit=1
						Exit Function
					End If
				End If
			Next
		End If		
	End Function
	Public Sub Get_Chan_TopicOrder()
		If Dvbbs.Forum_ChanSetting(0)="1" And Not isnull(smsuserlist) and smsuserlist<>"" Then
		'随机数
		Dim MaxUserID,MaxLength
		MaxLength=12
		set rs=Dvbbs.execute("select Max(userid) from [dv_user]")
		MaxUserID=rs(0)
		Dim num1,rndnum
		Randomize
		Do While Len(rndnum)<4
			num1=CStr(Chr((57-48)*rnd+48))
			rndnum=rndnum&num1
		loop
		MaxUserID=rndnum & MaxUserID
		MaxLength=MaxLength-len(MaxUserID)
		Select Case MaxLength
		Case 7
			MaxUserID="0000000" & MaxUserID
		Case 6
			MaxUserID="000000" & MaxUserID
		Case 5
			MaxUserID="00000" & MaxUserID
		Case 4
			MaxUserID="0000" & MaxUserID
		Case 3
			MaxUserID="000" & MaxUserID
		Case 2
			MaxUserID="00" & MaxUserID
		Case 1
			MaxUserID="0" & MaxUserID
		Case 0
			MaxUserID=MaxUserID
		End Select
		Session("challengeWord")=MaxUserID
		Response.Write "<iframe name=getchallengeword frameborder=0 width=100% height=0 scrolling=no src=""pay_topic_postforumid.asp?chanWord="&Session("challengeWord")&"""></iframe>"
		End If
	End Sub
	Public Sub Get_ForumTreeCode()
		Dim mailbody
		Set Rs=Dvbbs.Execute("select b.layer,b.orders,b.EmailFlag,b.username,u.userEmail from "&TotalUseTable&" b inner join [Dv_user] u on b.postuserid=u.userid where b.AnnounceID="&ParentID)
		If Not(rs.EOF And rs.BOF) Then
			If IsNull(Rs(0)) Then
				iLayer=1
			Else
				iLayer=Rs(0)
			End If
			If IsNull(Rs(1)) Then
				iOrders=0
			Else
				iOrders=Rs(1)
			End If
			If Rs(3)=Dvbbs.membername Then
				If Cint(Dvbbs.GroupSetting(4))=0 Then Dvbbs.AddErrCode(73)
			End If
			If rs(3) <> Dvbbs.membername Then 
				Dim sUrl,Email,TempArray,etopic
				TempArray = Split(template.html(10),"||")
				sUrl=Dvbbs.Get_ScriptNameUrl
				If Rs(2)=1 Or Rs(2)=3 Then
					etopic=Replace(template.Strings(25),"{$forumname}",Dvbbs.Forum_info(0))
					email=Rs(4)
					mailbody = TempArray(0)
					mailbody = Replace(mailbody,"{$boardid}",Dvbbs.BoardID)
					mailbody = Replace(mailbody,"{$forumname}",Dvbbs.Forum_info(0))
					mailbody = Replace(mailbody,"{$topicid}",RootID)
					mailbody = Replace(mailbody,"{$star}",Star)
					mailbody = Replace(mailbody,"{$surl}",sUrl)
					mailbody = Replace(mailbody,"{$parentid}",ParentID)
					Select Case Dvbbs.Forum_Setting(2)
					Case "1"
						jmail email,etopic,mailbody
					Case "2"
						Cdonts email,etopic,mailbody
					Case "3"
						aspemail email,etopic,mailbody
					End Select
				End If
				If Rs(2)=2 Or Rs(2)=3 Then
					mailbody = TempArray(1)
					mailbody = Replace(mailbody,"{$boardid}",Dvbbs.BoardID)
					mailbody = Replace(mailbody,"{$topicid}",RootID)
					mailbody = Replace(mailbody,"{$star}",Star)
					mailbody = Replace(mailbody,"{$surl}",sUrl)
					mailbody = Replace(mailbody,"{$parentid}",ParentID)
					Dvbbs.Execute("insert into dv_message(incept,sender,title,content,sendtime,flag,issend) values('"&Rs(3)&"','"&Dvbbs.Forum_info(0)&"','"&template.Strings(26)&"','"&mailbody&"',"&SqlNowString&",0,1)")
					update_user_msg(Rs(3))
				End If
			End If
		Else
			iLayer=1
			iOrders=0
		End If
		Set Rs=Nothing
		If RootID<>0 Then 
			iLayer=ilayer+1
			Dvbbs.Execute "update "&TotalUseTable&" set orders=orders+1 where RootID="&cstr(RootID)&" and orders>"&cstr(iOrders)
			iOrders=iOrders+1
		End If
	End Sub
	Public Sub Update_Edit_Announce()
		Dim re,LastBoard,LastTopic
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="<br>"
		Content = re.Replace(Content,"[br]")
		re.Pattern="\[align=right\]\[color=#000066\](.*)\[\/color\]\[\/align\]"
		Content = re.Replace(Content,"")
		re.Pattern="<div align=right><font color=#000066>(.*)<\/font><\/div>"
		Content = re.Replace(Content,"")
		re.Pattern="\[br\]>"
		Content = re.Replace(Content,"<br>")		
		Set re=Nothing
		If Dvbbs.membername<>UserName Then 
			If Dvbbs.forum_setting(49)="1" Then char_changed = "[align=right][color=#000066][此贴子已经被"&Dvbbs.membername&"于"&Now()&"编辑过][/color][/align]"
		Else
			If Dvbbs.forum_setting(48)="1" Then char_changed = "[align=right][color=#000066][此贴子已经被作者于"&Now()&"编辑过][/color][/align]"
		End If
		Get_Edit_PermissionInfo
		Dim Contentdata
		Contentdata=Content
		Dvbbs.ShowErr
		'取出当前版面最后回复id,如果本帖为最后回复则更新相应数据
		Set Rs = Dvbbs.Execute("select LastPost from dv_board where boardid="&Dvbbs.BoardID)
		If not (Rs.EOF And Rs.BOF) Then
			If Not IsNull(rs(0)) And rs(0)<>"" then
				LastBoard=split(rs(0),"$")
				If ubound(LastBoard)=7 Then
					If Clng(LastBoard(6))=Clng(AnnounceID) Then
						LastPost=LastBoard(0) & "$" & LastBoard(1) & "$" & Now() & "$" & Replace(cutStr(reubbcode(topic),20),"$","&#36;") & "$" & LastBoard(4) & "$" & LastBoard(5) & "$" & LastBoard(6) & "$" & Dvbbs.BoardID
						dvbbs.execute("update dv_board set LastPost='"&SimEncodeJS(Replace(LastPost,"'",""))&"' where boardid="&Dvbbs.BoardID)
					End If
				End If
			End If
		End If

		'取得当前主题最后回复id,如果本帖为最后回复则更新相应数据
		Set Rs=Dvbbs.Execute("select LastPost,istop from dv_topic where topicid="&rootid)
		If Not (Rs.Eof And Rs.Bof) Then
			istop=rs(1)
			If Not Isnull(Rs(0)) And Rs(0)<>"" Then
				LastTopic=split(rs(0),"$")
				If Ubound(LastTopic)=7 Then
					If Clng(LastTopic(1))=Clng(Announceid) Then
						LastPost=LastTopic(0) & "$" & LastTopic(1) & "$" & Now() & "$" & Replace(cutStr(reubbcode(Contentdata),20),"$","&#36;") & "$" & LastTopic(4) & "$" & LastTopic(5) & "$" & LastTopic(6) & "$" & Dvbbs.BoardID
						dvbbs.execute("update dv_topic set LastPost='"&Replace(LastPost,"'","")&"' where topicid="&rootid)
					End If
				End If
			End If
		End If

		Set Rs = Server.CreateObject("ADODB.Recordset")
		SQL="SELECT * FROM "&TotalUseTable&" where AnnounceID="&Announceid&" And username='"&trim(username)&"'"
		rs.Open SQL,conn,1,3
		If Rs.EOF And Rs.BOF Then
			If Not CanEditPost Then Dvbbs.AddErrCode(77)
		ElseIf Not Dvbbs.master And rs("locktopic")=1 then
			Dvbbs.AddErrCode(78)
		Else
			If Rs("parentid")=0 then
				If istop=1 Then
					If IsSqlDataBase=1 Then
						dvbbs.execute("update dv_topic set title='"&topic&"',LastPostTime=dateadd(day,100,"&SqlNowString&"),TopicMode="&MyTopicMode&" where topicid="&rootid)
					Else
						dvbbs.execute("update dv_topic set title='"&topic&"',LastPostTime=dateadd('d',100,"&SqlNowString&"),TopicMode="&MyTopicMode&" where topicid="&rootid)
					End If
				ElseIf istop=3 Then
					If IsSqlDataBase=1 Then
						dvbbs.execute("update dv_topic set title='"&topic&"',LastPostTime=dateadd(day,300,"&SqlNowString&"),TopicMode="&MyTopicMode&" where topicid="&rootid)
					Else
						dvbbs.execute("update dv_topic set title='"&topic&"',LastPostTime=dateadd('d',300,"&SqlNowString&"),TopicMode="&MyTopicMode&" where topicid="&rootid)
					End If
				Else
					dvbbs.execute("update dv_topic set title='"&topic&"',TopicMode="&MyTopicMode&" where topicid="&rootid)
				End If
			End If
			Content = Html2Ubb(Content)
			Rs("Topic") =Replace(Topic,"''","'")
			rs("Body") =Replace(Content,"''","'")
			rs("length")=Dvbbs.strlength(Content)
			rs("ip")=Dvbbs.UserTrueIP
			rs("Expression")=Expression
			rs("signflag")=signflag
			rs("emailflag")=mailflag
			If Rs("isupload")=0 And ihaveupfile=1 Then Rs("isupload")=1
			Dim UbblistBody
			UbblistBody = Ubblist(Content)
			Rs("Ubblist")=UbblistBody
			rs.Update
			If ihaveupfile=1 Then dvbbs.execute("update dv_upfile set F_AnnounceID='"&rootid&"|"&AnnounceID&"',F_Readme='"&Replace(Rs("Topic"),"'","''")&"',F_flag=0 where F_ID in ("&upfileinfo&")")
		End If	
		Rs.Close
		Set Rs=Nothing
		Dvbbs.ShowErr()
	End Sub
	'更新版面数据和缓存
	Public Sub UpDate_BoardInfoAndCache()
		Dim UpdateBoardID
		If Dvbbs.Board_Data(3,0)<> "" Then 
			UpdateBoardID=Dvbbs.Board_Data(3,0) & "," & Dvbbs.BoardID	
		Else
			UpdateBoardID=Dvbbs.BoardID
		End If
		Dim updateboard,i
		updateboard=Split(UpdateBoardID,",")
		If Action = 6 Then
			SQL="update Dv_board set PostNum=PostNum+1,todaynum=todaynum+1,LastPost='"&SimEncodeJS(LastPost_1)&"' where boardid in ("&UpdateBoardID&")"
		ElseIf Action = 5 Or Action = 7 Then
			SQL="update Dv_board set PostNum=PostNum+1,TopicNum=TopicNum+1,todaynum=todaynum+1,LastPost='"&SimEncodeJS(LastPost_1)&"' where boardid in ("&UpdateBoardID&")"
		End If
		Dvbbs.Execute(sql)
		For i= 0 to UBound(updateboard)
			Dvbbs.ReloadBoardCache updateboard(i),1,9,1'版面ID，发贴数，最后一个参数是1 表示相加
			If Not Action = 6 Then Dvbbs.ReloadBoardCache updateboard(i),1,10,1'主题数
			Dvbbs.ReloadBoardCache updateboard(i),1,12,1'今日贴
			Dvbbs.ReloadBoardCache updateboard(i),LastPost_1,14,0
		Next
	End Sub
	Public Sub UpDate_ForumInfoAndCache()
		Dim updateinfo,LastPostTime
		Dim Forum_LastPost,Forum_TodayNum,Forum_MaxPostNum
		Forum_LastPost=Dvbbs.CacheData(15,0)
		Forum_TodayNum=Dvbbs.CacheData(9,0)
		Forum_MaxPostNum=Dvbbs.CacheData(12,0)
		LastPostTimes=split(Forum_LastPost,"$")
		LastPostTime=LastPostTimes(2)
		If Not IsDate(LastPostTime) Then LastPostTime=Now()
		If datediff("d",LastPostTime,Now())=0 Then
			If CLng(Forum_TodayNum)+1 > CLng(Forum_MaxPostNum) Then 
				updateinfo=",Forum_MaxPostNum=Forum_TodayNum+1,Forum_MaxPostDate="&SqlNowString&""
				Dvbbs.ReloadSetupCache Now(),13
				Dvbbs.ReloadSetupCache CLng(Forum_TodayNum)+1,12
			End If
			Dvbbs.ReloadSetupCache CLng(Forum_TodayNum)+1,9
			If Action = 6 Then
				SQL="update Dv_setup set Forum_PostNum=Forum_PostNum+1,Forum_TodayNum=Forum_TodayNum+1,Forum_LastPost='"&LastPost&"' "&updateinfo&" "
			Else
				SQL="update Dv_setup set Forum_TopicNum=Forum_TopicNum+1,Forum_PostNum=Forum_PostNum+1,Forum_TodayNum=Forum_TodayNum+1,Forum_LastPost='"&LastPost&"' "&updateinfo&" "
			End If
		Else
			If Action = 6 Then
				SQL="update Dv_setup set Forum_PostNum=Forum_PostNum+1,forum_YesTerdayNum="&CLng(Forum_TodayNum)&",Forum_TodayNum=1,Forum_LastPost='"&LastPost&"' "
			Else
				SQL="update Dv_setup set Forum_TopicNum=Forum_TopicNum+1,Forum_PostNum=Forum_PostNum+1,forum_YesTerdayNum="&CLng(Forum_TodayNum)&",Forum_TodayNum=1,Forum_LastPost='"&LastPost&"' "
			End If
			Dvbbs.ReloadSetupCache 1,9
		End If
		'更新总固顶部分数据和缓存
		If Not Action = 6 Then
			If Myistop=2 Then
				Dim tmpstr
				If Dvbbs.CacheData(28,0)="" Then
					tmpstr=", Forum_alltopnum='"&RootID&"'"
					Dvbbs.ReloadSetupCache RootID,28
				Else
					tmpstr=", Forum_alltopnum='"&Dvbbs.CacheData(28,0)&","&RootID&"'"
					Dvbbs.ReloadSetupCache Dvbbs.CacheData(28,0)&","&RootID,28
				End If 
				SQL=SQl&tmpstr
			End If
			Dvbbs.ReloadSetupCache CLng(Dvbbs.CacheData(7,0))+1,7'主题数
		End If
		Dvbbs.ReloadSetupCache CLng(Dvbbs.CacheData(8,0))+1,8'文章数
		Dvbbs.ReloadSetupCache LastPost,15
		Dvbbs.Execute(SQL)
	End Sub
	Public Sub succeed()
		Dim TempStr,PostRetrunName,tourl,returnurl
		If IsAudit=1 And Action <> 8 Then Dvbbs.BoardID=LockTopic
		Dvbbs.Stats = Dvbbs.Stats & template.Strings(20)
		TempStr = template.html(8)
		Select case Dvbbs.Board_Setting(17)
		case "1"
			tourl = "index.asp"
			PostRetrunName=template.Strings(13)
		case "2"
			tourl="list.asp?boardid="&Dvbbs.boardid
			PostRetrunName=template.Strings(14)
		case "3"
			If IsAudit=1 And Action <> 8 Then
				tourl="list.asp?boardid="&Dvbbs.boardid
				If IsAuditcheck=1 Then
					PostRetrunName="由于您发表的贴子含敏感内容，您的贴子需要管理员审核过才可以见到。"
				Else
					PostRetrunName=template.Strings(19)
				End If 
			Else
				Select Case Action
				Case 5
				tourl="dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID
				PostRetrunName=template.Strings(15)
				Case 6
				tourl="dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID&"&star="&Star&"#"&Announceid
				PostRetrunName=template.Strings(16)
				Case 7
				tourl="dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID
				PostRetrunName=template.Strings(17)
				Case 8
				tourl="dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID&"&star="&Star&"#"&RootID
				PostRetrunName=template.Strings(18)
				End Select
			End If
		End Select
		returnurl="dispbbs.asp?boardid="&Dvbbs.boardid&"&id="&RootID
		TempStr = Replace(TempStr,"{$tourl}",tourl)
		TempStr = Replace(TempStr,"{$returnurl}",returnurl)
		TempStr = Replace(TempStr,"{$stats}",Dvbbs.Stats)
		TempStr = Replace(TempStr,"{$boardname}",Dvbbs.BoardType)
		TempStr = Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
		TempStr = Replace(TempStr,"{$page}",page)
		TempStr = Replace(TempStr,"{$PostRetrunName}",PostRetrunName)
		Response.Write TempStr
	End Sub
	Private Function checktable(Table)
		Table=Right(Trim(Table),2)
		If Not IsNumeric(table) Then Table=Right(Trim(Table),1)
		If Not IsNumeric(table) Then Dvbbs.AddErrCode(30)
		checktable="Dv_bbs"&table
	End Function
	'检查提交来源
	Public Sub CheckfromScript()
		If Not Dvbbs.ChkPost() Or  Not(IsArray(Session(Dvbbs.CacheName & "UserID"))) Then Dvbbs.AddErrCode(42):Dvbbs.Showerr()
 		If CStr(Request.Cookies("Dvbbs"))=CStr(Dvbbs.Boardid) Then Dvbbs.AddErrCode(30):Dvbbs.Showerr()
 		If (Not ChkUserLogin) And (Action = 5 Or Action = 6 Or Action = 7) Then Dvbbs.AddErrCode(12):Dvbbs.Showerr()	
	End Sub
	'判断发贴时间间隔
	Private Sub  CheckpostTime()
		If Dvbbs.Board_Setting(30)="1"  Then
			Dim mypostinfo
			mypostinfo=Session(Dvbbs.CacheName & "UserID")
			If DateDiff("s",mypostinfo(2),Now())<CLng(Dvbbs.Board_Setting(31)) Then
				 Response.redirect "showerr.asp?ErrCodes=<Br>"+"<li>本论坛限制发贴距离时间为"&Dvbbs.Board_Setting(31)&"秒，请稍后再发。&action=OtherErr"
			End If
		End If
	End Sub 
	'检查用户身份
	Public Function ChkUserLogin()
 		ChkUserLogin=False
 		'取得发贴用户名和密码
		UserName=Dvbbs.Checkstr(Trim(Request.Form("username")))
		'校验用户名和密码是否合法
		'If UserName="" Or Dvbbs.strLength(userName)>Cint(Dvbbs.Forum_setting(41)) Or Dvbbs.strLength(userName) < Cint(Dvbbs.Forum_setting(40)) Then Dvbbs.AddErrCode(17)
		If UserName="" Then Dvbbs.AddErrCode(17)
		If Not IstrueName(UserName) Then Dvbbs.AddErrCode(18)
		Dvbbs.ShowErr()
		If Action = 8 Then
			'编辑贴子，检查用户身份
			UserPassWord=Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("password")))
			SQL = "Select JoinDate,UserID,UserPost,UserGroupID,userclass,lockuser,TruePassWord From [Dv_User] Where UserID="&Dvbbs.UserID
		Else
			'检查用户是否当前用户
			If UserName<>Dvbbs.MemberName Then
				Reuser=True
				UserPassWord=Dvbbs.Checkstr(Trim(Request.Form("passwd")))
				UserPassWord=md5(UserPassWord,16)
				SQL = "Select JoinDate,UserID,UserPost,UserGroupID,userclass,lockuser,userpassword From [Dv_User] Where UserName='"&UserName&"' "
			Else
				UserPassWord=Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("password")))
				SQL = "Select JoinDate,UserID,UserPost,UserGroupID,userclass,lockuser,TruePassWord From [Dv_User] Where UserID="&Dvbbs.UserID		
			End If
		End If
		If Len(UserPassWord)<>16 AND Len(UserPassWord)<>32 Then Dvbbs.AddErrCode(18)
 		Set Rs=Dvbbs.Execute(SQL)
 		If Not Rs.EOF Then
			If Not (UserPassWord<>rs(6) Or rs(5)=1 or rs(3)=5) Then
 				ChkUserLogin=True
 				Dvbbs.UserID=Rs(1)
 				UserPost=Rs(2)
 				GroupID=Rs(3)
 				userclass=Rs(4)
				Response.cookies("upNum")=0
 			Else
 				Dvbbs.EmptyCookies
 				Dvbbs.LetGuestSession()			
			End If
 		End If
 		Set Rs = Nothing
 	End Function
 	'更新用户积分，所需外部变量,UserPost,userid,（外加发贴回贴的积分设置数据）
	Public Sub updatepostuser()
		'投票，发贴，更新积分
		Dim cUserInfo
		cUserInfo = Session(Dvbbs.CacheName & "UserID")
		'更新最后发贴时间
		cUserInfo(2) = Now
		If Action = 5 Or Action = 7 Then 
			Dvbbs.Execute("update [Dv_user] set UserLastIP='"&Dvbbs.usertrueip&"',UserPost=UserPost+1,UserTopic=UserTopic+1,userWealth=userWealth+"&Clng(Dvbbs.Forum_user(1))&",userEP=userEP+"&Clng(Dvbbs.Forum_user(6))&",userCP=userCP+"&Clng(Dvbbs.Forum_user(11))&",UserToday='"&Clng(Dvbbs.UserToday(0))+1&"|"&Clng(Dvbbs.UserToday(1))&"|"&Clng(Dvbbs.UserToday(2))&"' Where UserID="&Dvbbs.userID&"")
			If Not Reuser Then
				UserPost=UserPost+1
				cUserInfo(21)=cUserInfo(21)+Clng(Dvbbs.Forum_user(1))
				cUserInfo(22)=cUserInfo(22)+Clng(Dvbbs.Forum_user(6))
				cUserInfo(23)=cUserInfo(23)+Clng(Dvbbs.Forum_user(11))
			End If
		ElseIf Action = 6 Then '回贴更新积分。
			If Not Reuser Then 
				Dvbbs.Execute("update [Dv_user] set UserLastIP='"&Dvbbs.usertrueip&"',UserPost=UserPost+1,userWealth=userWealth+"&Clng(Dvbbs.Forum_user(2))&",userEP=userEP+"&Clng(Dvbbs.Forum_user(7))&",userCP=userCP+"&Clng(Dvbbs.Forum_user(12))&",UserToday='"&Clng(Dvbbs.UserToday(0))+1&"|"&Clng(Dvbbs.UserToday(1))&"|"&Clng(Dvbbs.UserToday(2))&"' Where UserID="&Dvbbs.userID&"")
				UserPost=UserPost+1
				cUserInfo(21)=cUserInfo(21)+Clng(Dvbbs.Forum_user(2))
				cUserInfo(22)=cUserInfo(22)+Clng(Dvbbs.Forum_user(7))
				cUserInfo(23)=cUserInfo(23)+Clng(Dvbbs.Forum_user(12))
			Else
				Dvbbs.Execute("update [Dv_user] set UserLastIP='"&Dvbbs.usertrueip&"',UserPost=UserPost+1,userWealth=userWealth+"&Clng(Dvbbs.Forum_user(2))&",userEP=userEP+"&Clng(Dvbbs.Forum_user(7))&",userCP=userCP+"&Clng(Dvbbs.Forum_user(12))&" Where UserID="&Dvbbs.userID&"")
			End If
		End If
		If Not Reuser Then 
			cUserInfo(8)=UserPost+1
			cUserInfo(36)=Clng(Dvbbs.UserToday(0))+1 & "|" & Clng(Dvbbs.UserToday(1)) & "|" & Clng(Dvbbs.UserToday(2))
		End If
		Session(Dvbbs.CacheName & "UserID") = cUserInfo
		'发贴数字能整除十则更新用户等级。(Updategrade())
		If UserPost mod 10 < 1  Then Updategrade()
	End Sub
	'更新用户等级，所需外部变量,UserPost,GroupID,userid
	Public Sub Updategrade()
		Dim titlepic
		Dim cUserInfo,GroupID_Q
		If Not Reuser Then  cUserInfo = Session(Dvbbs.CacheName & "UserID")
		'检查用户等级数据表中是否有匹配行
		Set Rs=Dvbbs.Execute("select MinArticle,IsSetting,ParentGID from Dv_UserGroups where usertitle='"&userclass&"'")
		If Rs.Eof Or Rs.BOF Then
			Set Rs=Nothing:Set Rs=Dvbbs.Execute("select top 1 usertitle,GroupPic,UserGroupID,IsSetting,ParentGID from Dv_UserGroups where (ParentGID="&GroupID&" Or UserGroupID="&GroupID&") and Minarticle<="&UserPost&" and not Minarticle=-1 order by MinArticle desc")
			If Not(Rs.EOF And Rs.BOF) Then 
				userclass=Rs(0)
				titlepic=Rs(1)
				If Rs(3)=1 Then
					GroupID=Rs(2)
				Else
					GroupID=Rs(4)
				End If
				Set RS=Nothing 
			Else
				Set Rs=Dvbbs.Execute("select top 1 usertitle,GroupPic,UserGroupID,IsSetting,ParentGID from Dv_UserGroups where UserGroupID="&GroupID&" and Minarticle=-1 order by UserGroupID")
				If Not(Rs.EOF And Rs.BOF) Then 
					userclass=Rs(0)
					titlepic=Rs(1)
					If Rs(3)=1 Then
						GroupID=Rs(2)
					Else
						GroupID=Rs(4)
					End If
					Set RS=Nothing 
				Else
					Set RS=Nothing:Set Rs=Dvbbs.Execute("select top 1 GroupPic from Dv_UserGroups where ParentGID>0 And not Minarticle=-1 order by MinArticle")
					titlepic=Rs(0)
					Set RS=Dvbbs.Execute("select usertitle from Dv_UserGroups where UserGroupID="&GroupID)
					userclass=Rs(0)
				End If
			End If
		Else	
			If Rs(0)>-1 Then
				'如果为自定义等级，则取其父类GroupID做升级依据
				GroupID_Q=GroupID
				If Rs(1)=1 And Rs(2)>0 Then GroupID_Q=Rs(2)
				Set Rs=Nothing:Set Rs=Dvbbs.Execute("select top 1 usertitle,GroupPic,UserGroupID,IsSetting,ParentGID from Dv_UserGroups where ParentGID="&GroupID_Q&" and Minarticle<="&UserPost&" and not MinArticle=-1 order by MinArticle desc,UserGroupID")
				If Not (Rs.EOF And Rs.BOF) Then 
					userclass=Rs(0)
					titlepic=Rs(1)
					If Rs(3)=1 Then
						GroupID=Rs(2)
					Else
						GroupID=Rs(4)
					End If
					Set Rs=Nothing 
				Else
					Set Rs=Nothing
					Set Rs=Dvbbs.Execute("select top 1 GroupPic from Dv_UserGroups where ParentGID>0 And not Minarticle=-1 order by MinArticle")
					titlepic=Rs(0)
					Set Rs=Nothing
					Set Rs=Dvbbs.Execute("select usertitle from Dv_UserGroups where UserGroupID="&GroupID)
					userclass=Rs(0)
					Set Rs=Nothing 
				End If
			Else
				Set Rs=Dvbbs.Execute("select usertitle,GroupPic,UserGroupID,IsSetting,ParentGID from Dv_UserGroups where usertitle='"&userclass&"'")
				If Not (Rs.EOF And Rs.BOF) Then 
					userclass=Rs(0)
					titlepic=Rs(1)
					If Rs(3)=1 Then
						GroupID=Rs(2)
					Else
						GroupID=Rs(4)
					End If
				End If
				Set Rs=Nothing 
			End If
		End If
		Dvbbs.Execute("update [Dv_User] set userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&" where userid="&dvbbs.UserID)
		If Not Reuser Then 
			cUserInfo(18)=userclass
			cUserInfo(19)=GroupID
			Session(Dvbbs.CacheName & "UserID") = cUserInfo
		End If
	End Sub
End Class

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
	cutStr=Replace(cutStr,chr(10),"")
	cutStr=Replace(cutStr,chr(13),"")
End Function
'过滤不必要UBB
Function reUBBCode(strContent)
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	strContent=Replace(strContent,"&nbsp;"," ")
	re.Pattern="(\[QUOTE\])(.|\n)*(\[\/QUOTE\])"
	strContent=re.Replace(strContent,"")
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
	strContent=Replace(strContent,"<I></I>","")
	set re=Nothing
	reUBBCode=strContent
End Function

'通用函数
Function IstrueName(uName)
	IstrueName=False
	If InStr(uName,"=")>0 Then Exit Function
	If InStr(uName,"%")>0 Then Exit Function 
	If InStr(uName,Chr(32))>0 Then Exit Function 
	If InStr(uName,"?")>0 Then Exit Function 
	If InStr(uName,"&")>0 Then Exit Function 
	If InStr(uName,";")>0 Then Exit Function 
	If InStr(uName,",")>0 Then Exit Function 
	If InStr(uName,"'")>0 Then Exit Function 
	If InStr(uName,Chr(34))>0 Then Exit Function 
	If InStr(uName,chr(9))>0 Then Exit Function 
	If InStr(uName,"")>0 Then Exit Function 
	If InStr(uName,"$")>0 Then Exit Function
	IstrueName=True 	
End Function
Function SimEncodeJS(str)
	If Not IsNull(str) Then
		str = Replace(str, "\", "\\")
		str = Replace(str, chr(34), "\""")
		str = Replace(str, chr(39), "\'")
		str = Replace(str, chr(10), "\n")
		str = Replace(str, chr(13), "\r")
		SimEncodeJS=str
	End If
End Function
'发贴时用，为了减少入库量
Function Html2Ubb(str)
	If Str<>"" And Not IsNull(Str) Then
		Dim re,tmpstr
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern = "(<br>)"
		Str = re.Replace(Str,"[br]")
		If Dvbbs.Board_Setting(5)="0" Then
			'先去掉标记中的换行
			re.Pattern="(<(i|b|p)>)"
			Str=re.Replace(Str,"[$2]")
			re.Pattern="(<(\/i|\/b|\/p)>)"
			Str=re.Replace(Str,"[$2]")
			re.Pattern="(>)("&vbNewLine&")(<)"
			Str=re.Replace(Str,"$1$3") 
			re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
			Str=re.Replace(Str,"$1$3")
			re.Pattern="(<DIV class=quote>)((.|\n)*)(<\/div>)"
			Str=re.Replace(Str,"[quote]$2[/quote]")
			re.Pattern="<(.[^>]*)>"
			Str=re.Replace(Str,"")
			re.Pattern="(\[(i|b|p)\])"
			Str=re.Replace(Str,"<$2>")
			re.Pattern="(\[(\/i|\/b|\/p)\])"
			Str=re.Replace(Str,"<$2>")
		End If
		Str = Replace(Str, "[br]", CHR(13) & CHR(10))
		re.Pattern = "(&nbsp;)"
		Str = re.Replace(Str,Chr(9))
		re.Pattern = "(<STRONG>)"
		Str = re.Replace(Str,"<b>")
		re.Pattern = "(<\/STRONG>)"
		Str = re.Replace(Str,"</b>")
		re.Pattern ="(<TBODY>)"
		Str = re.Replace(Str,"")
		re.Pattern ="(<\/TBODY>)"
		Str = re.Replace(Str,"")
		Set Re=Nothing
		Html2Ubb = Str
	Else
		Html2Ubb = ""
	End If
End Function
'更新用户短信通知信息（新短信条数||新短讯ID||发信人名）
Sub UPDATE_User_Msg(username)
	Dim msginfo,i,UP_UserInfo,newmsg
	newmsg=newincept(username)
	If newmsg>0 Then
		msginfo=newincept(username) & "||" & inceptid(1,username) & "||" & inceptid(2,username)
	Else
		msginfo="0||0||null"
	End If
	Dvbbs.execute("UPDATE [Dv_User] Set UserMsg='"&Dvbbs.CheckStr(msginfo)&"' WHERE username='"&Dvbbs.CheckStr(username)&"'")
	If username=Dvbbs.MemberName Then
		UP_UserInfo=Session(Dvbbs.CacheName & "UserID")
		UP_UserInfo(30)=msginfo
		Session(Dvbbs.CacheName & "UserID")=UP_UserInfo
	Else
		Call Dvbbs.NeedUpdateList(username,1)
	End If
End Sub

'统计留言
Function newincept(iusername)
Dim Rs
Rs=Dvbbs.execute("SELECT Count(id) FROM Dv_Message WHERE flag=0 And issend=1 And DelR=0 And incept='"& iusername &"'")
    newincept=Rs(0)
	Set Rs=nothing
	If isnull(newincept) Then newincept=0
End Function

Function inceptid(stype,iusername)
	Dim Rs
	Set Rs=Dvbbs.execute("SELECT top 1 id,sender FROM Dv_Message WHERE flag=0 And issend=1 And DelR=0 And incept='"& iusername &"'")
	If not rs.eof Then
		If stype=1 Then
			inceptid=Rs(0)
		Else
			inceptid=Rs(1)
		End If
	Else
		If stype=1 Then
			inceptid=0
		Else
			inceptid="null"
		End If
	End If
	Set Rs=nothing
End Function
%>