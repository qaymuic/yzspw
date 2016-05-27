<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<%
Dim totalusetable,AnnounceID,ReplyID
Dim MyPost,UserName
Dim postbuyuser,bgcolor
Dvbbs.Loadtemplates("post")
Set MyPost = New Dvbbs_Post
Dvbbs.Stats = MyPost.ActionName
Dvbbs.Nav()
Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
Dim dv_ubb,abgcolor
Dim EmotPath
EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)		'em心情路径
Set dv_ubb=new Dvbbs_UbbCode
MyPost.Show_Post_Body
Set MyPost = Nothing
Dvbbs.ActiveOnline
Dvbbs.Footer

Class Dvbbs_Post
	Public Action,ActionName,Star,Page,IsAudit,ToAction
	Private ParentID,RootID,Topic,Content,char_changed,signflag,mailflag,iLayer,iOrders
	Private TopTopic,IsTop,LastPost,LastPost_1,UpLoadPic_n,ihaveupfile,smsuserlist,upfileinfo
	Private UserPassWord,UserPost,GroupID,UserClass,DateAndTime,DateTimeStr,Expression,MyLastPostTime,LastPostTimes
	Private LockTopic,MyLockTopic,MyIsTop,MyIsTopAll,MyTopicMode
	Private CanLockTopic,CanTopTopic,CanTopTopic_a,CanEditPost,Rs,SQL,i
	Private vote,votetype,votenum,votetimeout,voteid,isvote
	Private Sub Class_Initialize()
		If Dvbbs.Board_Setting(0)="1" And Not Dvbbs.Master Then
			Response.redirect "showerr.asp?action=lock&boardid="&dvbbs.boardID&"" 
		End If
		If Dvbbs.IsReadonly() And Not Dvbbs.Master  Then Response.redirect "showerr.asp?action=readonly&boardid="&dvbbs.boardID&"" 
		Action = Request("Action")
		TotalUseTable = Dvbbs.NowUseBBS
		Select Case Action
		Case "new"
			Action = 1
			ActionName = template.Strings(0)
			ToAction = "SavePost.asp?Action=snew&boardid="&Dvbbs.BoardID
			If Dvbbs.GroupSetting(3)="0" Then Dvbbs.AddErrCode(70)
		Case "re"
			Action = 2
			ActionName = template.Strings(2)
			ToAction = "SavePost.asp?Action=sre&method=Topic&boardid="&Dvbbs.BoardID
			If Dvbbs.GroupSetting(5)="0" then Dvbbs.AddErrCode(71)
		Case "vote"
			Action = 3
			ActionName = template.Strings(4)
			ToAction = "SavePost.asp?Action=svote&boardid="&Dvbbs.BoardID
			If Dvbbs.GroupSetting(8)="0" then Dvbbs.AddErrCode(56)
		Case "edit"
			Action = 4
			ActionName = template.Strings(6)
			ToAction = "SavePost.asp?Action=sedit&boardid="&Dvbbs.BoardID
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
	End Sub
	'Action 1=发贴、2=回帖、3=投票、4=编辑 主体部分
	Public Function Show_Post_Body()
		Chk_Post()
		Dim TempStr,TempArray,TempStr1,TempStr2,PostType
		signflag=1
		mailflag=0
		MyTopicMode=0
		If Dvbbs.UserID = 0 Then Dvbbs.AddErrCode(6)
		Dvbbs.ShowErr()
		TempStr = template.html(0)
		TempArray = Split(template.html(6),"||")
		If IsAudit = 1 Then TempStr = Replace(TempStr,"{$auditinfo}",template.Strings(9))
		TempStr = Replace(TempStr,"{$auditinfo}","")
		If Action=1 Then TempStr = Replace(TempStr,"{$topicmode}",TopicMode(TempArray(4)))
		If signflag=1 Then TempStr = Replace(TempStr,"{$checkbox1}","checked")
		TempStr = Replace(TempStr,"{$checkbox1}","")
		Select Case mailflag
		Case 0
			TempStr = Replace(TempStr,"{$checkbox2}","checked")
		Case 1
			TempStr = Replace(TempStr,"{$checkbox3}","checked")
		Case 2
			TempStr = Replace(TempStr,"{$checkbox4}","checked")
		Case 3
			TempStr = Replace(TempStr,"{$checkbox5}","checked")
		Case Else
			TempStr = Replace(TempStr,"{$checkbox2}","checked")
		End Select
		TempStr = Replace(TempStr,"{$checkbox2}","")
		TempStr = Replace(TempStr,"{$checkbox3}","")
		TempStr = Replace(TempStr,"{$checkbox4}","")
		TempStr = Replace(TempStr,"{$checkbox5}","")
		If Dvbbs.GroupSetting(51)="1" And (Action=1 Or Action=3) Then TempStr = Replace(TempStr,"{$useraction}",TempArray(3))
		
		'话题
		TempStr1 = Split(template.Strings(8),",")
		For i = 0 To Ubound(TempStr1)
			TempStr2 = TempStr2 & "<option value="&TempStr1(i)&">"&TempStr1(i)&"</option>"
		Next
		TempStr = Replace(TempStr,"{$topictype}",TempStr2)
		'特殊标题
		If Dvbbs.GroupSetting(51)="1" Then TempStr = Replace(TempStr,"{$topicstatsinfo}",TempArray(1))
		TempStr = Replace(TempStr,"{$topicstatsinfo}","")
		'验证码Board_Setting(4)
		If Dvbbs.Board_Setting(4)="0" Then
			TempStr = Replace(TempStr,"{$getcode}","")
		Else
			TempArray(5)= Replace(TempArray(5),"{$codestr}",Dvbbs.GetCode&"<span id=GetCode></span>")
			TempStr = Replace(TempStr,"{$getcode}",TempArray(5))
		End If
		'头像
		TempStr = Replace(TempStr,"{$Forum_PostFace}",Dvbbs.Forum_PostFace)
		TempStr = Replace(TempStr,"{$SelectFace}",Expression)
		'上传
		If (Dvbbs.GroupSetting(7)="1" Or Dvbbs.GroupSetting(7)="2") and Dvbbs.Forum_setting(43)<>999 Then TempStr = Replace(TempStr,"{$uploadinfo}",Temp_FileUpload)
		TempStr = Replace(TempStr,"{$uploadinfo}","")
		'标签判断部分
		TempStr = Replace(TempStr,"{$ihtml}",Dvbbs.Board_Setting(5))
		TempStr = Replace(TempStr,"{$iubb}",Dvbbs.Board_Setting(6))
		TempStr = Replace(TempStr,"{$iimg}",Dvbbs.Board_Setting(7))
		TempStr = Replace(TempStr,"{$iflash}",Dvbbs.Board_Setting(44))
		TempStr = Replace(TempStr,"{$imidea}",Dvbbs.Board_Setting(9))
		TempStr = Replace(TempStr,"{$iemot}",Dvbbs.Board_Setting(8))
		TempStr = Replace(TempStr,"{$iupload}",Dvbbs.GroupSetting(7))
		TempStr = Replace(TempStr,"{$bodylimited}",Dvbbs.Board_Setting(16))
		TempStr = Replace(TempStr,"{$imoney}",Dvbbs.Board_Setting(10))
		TempStr = Replace(TempStr,"{$ipoint}",Dvbbs.Board_Setting(11))
		TempStr = Replace(TempStr,"{$iusercp}",Dvbbs.Board_Setting(12))
		TempStr = Replace(TempStr,"{$ipower}",Dvbbs.Board_Setting(13))
		TempStr = Replace(TempStr,"{$iarticle}",Dvbbs.Board_Setting(14))
		TempStr = Replace(TempStr,"{$ireplyview}",Dvbbs.Board_Setting(15))
		TempStr = Replace(TempStr,"{$iusemoney}",Dvbbs.Board_Setting(23))
		TempStr = Replace(TempStr,"{$iuseusername}",Dvbbs.Board_Setting(56))	
		'ubb部分
		PostType = 1
		TempStr = Replace(TempStr,"{$PostType}",PostType)
		TempStr = Replace(TempStr,"{$getubb}",Temp_UBBHTML())
		'发贴心情
		TempStr = Replace(TempStr,"{$Forum_emot}",Dvbbs.Forum_emot)
		TempStr = Replace(TempStr,"{$Forum_sn}",Dvbbs.Forum_sn)
		TempStr = Replace(TempStr,"{$star}",Star)
		TempStr = Replace(TempStr,"{$page}",Page)
		TempStr = Replace(TempStr,"{$actionname}",ActionName)
		TempStr = Replace(TempStr,"{$toaction}",ToAction)
		TempStr = Replace(TempStr,"{$topiclimited}",Dvbbs.Board_Setting(45))
		TempStr = Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
		TempStr = Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
		TempStr = Replace(TempStr,"{$alertcolor}",Dvbbs.mainsetting(1))
		TempStr = Replace(TempStr,"{$Forum_Emot}",Replace(Dvbbs.Forum_Emot,"|||","<><><>"))
		TempStr = Replace(TempStr,"{$htmltool}",template.html(12))
		If Dvbbs.GroupSetting(62)="0"  Then
			TempStr = Replace(TempStr,"{$AffordPost}","无限制")
		Else
			TempStr = Replace(TempStr,"{$AffordPost}",Dvbbs.GroupSetting(62))
		End If
		TempStr = Replace(TempStr,"{$UserToday}",Dvbbs.UserToday(0))
		Select Case Action
		Case 1
			TempStr = Replace(TempStr,"{$rehiddeninput}","")
			TempStr = Replace(TempStr,"{$reinfo1}","")
			TempStr = Replace(TempStr,"{$edithiddeninput}","")
			TempStr = Replace(TempStr,"{$voteinfo}","")
			TempStr = Replace(TempStr,"{$useraction}","")
			TempStr = Replace(TempStr,"{$retopicloop}","")
			TempStr = Replace(TempStr,"{$membername}",Dvbbs.MemberName)
			TempStr = Replace(TempStr,"{$memberword}","**********")
			TempStr = Replace(TempStr,"{$content}",server.htmlencode(Session(Dvbbs.CacheName & "UserID")(37)))
			TempStr = Replace(TempStr,"{$retitle}","")
			TempStr = Replace(TempStr,"{$topic}","")
			TempStr = Replace(TempStr,"{$TopModeSelect}",MyTopicMode)
			TempStr = Replace(TempStr,"{$totalusetable}",TotalUseTable)
		Case 2
			TempStr = Replace(TempStr,"{$voteinfo}","")
			TempStr = Replace(TempStr,"{$edithiddeninput}","")
			TempStr = Replace(TempStr,"{$topic}","")
			Dim retopicloop
			retopicloop = Get_Re_TopicInfo
			retopicloop=Replace(retopicloop,"$","&#36;")
			TempStr = Replace(TempStr,"{$membername}",Dvbbs.MemberName)
			TempStr = Replace(TempStr,"{$memberword}","**********")
			TempStr = Replace(TempStr,"{$useraction}","")
			TempStr = Replace(TempStr,"{$retitle}",Topic)
			TempStr = Replace(TempStr,"{$rehiddeninput}",Re_HiddenInput())
			TempStr = Replace(TempStr,"{$reinfo1}",TempArray(0))
			TempStr = Replace(TempStr,"{$topicmode}","")
			TempStr = Replace(TempStr,"{$retopicloop}",retopicloop)
			TempStr = Replace(TempStr,"{$TopModeSelect}",MyTopicMode)
			TempStr = Replace(TempStr,"{$totalusetable}",TotalUseTable)
			If Session(Dvbbs.CacheName & "UserID")(37)<>""  And Trim(Content)="" Then
				TempStr = Replace(TempStr,"{$content}",server.htmlencode(Session(Dvbbs.CacheName & "UserID")(37)))
			Else
				TempStr = Replace(TempStr,"{$content}",Content)
			End If
			TempStr = Replace(TempStr,"&#36;","$")
		Case 3
			TempStr1 = template.html(1)
			TempStr1 = Replace(TempStr1,"{$votelimited}",Dvbbs.Board_Setting(32))
			TempStr1 = Replace(TempStr1,"{$posttimesel}",TempArray(2))
			TempStr = Replace(TempStr,"{$membername}",Dvbbs.MemberName)
			TempStr = Replace(TempStr,"{$memberword}","**********")
			TempStr = Replace(TempStr,"{$voteinfo}",TempStr1)
			TempStr = Replace(TempStr,"{$topicmode}",TopicMode(TempArray(4)))
			TempStr = Replace(TempStr,"{$rehiddeninput}","")
			TempStr = Replace(TempStr,"{$reinfo1}","")
			TempStr = Replace(TempStr,"{$edithiddeninput}","")
			TempStr = Replace(TempStr,"{$TopModeSelect}",MyTopicMode)
			TempStr = Replace(TempStr,"{$useraction}","")
			TempStr = Replace(TempStr,"{$retopicloop}","")
			TempStr = Replace(TempStr,"{$totalusetable}",TotalUseTable)
			TempStr = Replace(TempStr,"{$content}",server.htmlencode(Session(Dvbbs.CacheName & "UserID")(37)))
			TempStr = Replace(TempStr,"{$retitle}","")
			TempStr = Replace(TempStr,"{$topic}","")
			
		Case 4
			Get_Edit_TopicInfo()
			TempStr = Replace(TempStr,"{$membername}",Dvbbs.MemberName)
			TempStr = Replace(TempStr,"{$memberword}",Dvbbs.memberword)
			TempStr = Replace(TempStr,"{$rehiddeninput}","")
			TempStr = Replace(TempStr,"{$reinfo1}","")
			TempStr = Replace(TempStr,"{$voteinfo}","")
			TempStr = Replace(TempStr,"{$retopicloop}","")
			TempStr = Replace(TempStr,"{$useraction}","")
			TempStr = Replace(TempStr,"{$edithiddeninput}",Edit_HiddenInput())
			TempStr = Replace(TempStr,"{$topicmode}","")
			TempStr = Replace(TempStr,"{$totalusetable}",TotalUseTable)
			TempStr = Replace(TempStr,"{$TopModeSelect}",MyTopicMode)
			TempStr = Replace(TempStr,"{$topic}",Topic)
			TempStr = Replace(TempStr,"{$content}",Content)
		End Select
		Response.Write TempStr
		Response.Write "<script language=""javascript"">"
		Response.Write vbNewLine
		Response.Write "Maxtitlelength="&Dvbbs.Board_Setting(45)&";"
		Response.Write vbNewLine
		If Action = 1 Or Action = 3 Then
			Response.Write "ispostnew=1;"
			Response.Write vbNewLine
		End If	
		Response.Write "MaxConlength="&Dvbbs.Board_Setting(16)&";"
		Response.Write vbNewLine	
		Response.Write "</script>"
		Response.Cookies("Dvbbs")=""		
	End Function

	'专题下拉模式读取
	Public Function TopicMode(SelectMode)
		If Cint(Dvbbs.GroupSetting(65))=0 Then Exit Function
		If Replace(Dvbbs.Board_Setting(48),"$$","")="" Then Exit Function
		Dim BoardTopic,iii
		BoardTopic=Split(Dvbbs.Board_Setting(48),"$$")
		For iii=0 to Ubound(BoardTopic)-1
			TopicMode=TopicMode+"<option value="&(iii+1)
			TopicMode=TopicMode+" >"&BoardTopic(iii)&"</option>"
		Next
		TopicMode=Replace(SelectMode,"{$TopicMode}",TopicMode)
	End Function
	'通用判断
	Public Function Chk_Post()
		If Dvbbs.Board_Setting(43)="1" Then Dvbbs.AddErrCode(72)
		If Dvbbs.Board_Setting(1)="1" and Dvbbs.GroupSetting(37)="0" Then Dvbbs.AddErrCode(26)
		If Dvbbs.UserID>0 Then
			If Clng(Dvbbs.GroupSetting(52))>0 And DateDiff("s",Dvbbs.MyUserInfo(14),Now)<Clng(Dvbbs.GroupSetting(52))*60 Then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(21),"{$timelimited}",Dvbbs.GroupSetting(52))&"&action=OtherErr"
			If Dvbbs.GroupSetting(62)<>"0" And Not Action = 4 Then
				If Clng(Dvbbs.GroupSetting(62))<=Clng(Dvbbs.UserToday(0)) Then Response.redirect "showerr.asp?ErrCodes=<li>"&Replace(template.Strings(27),"{$topiclimited}",Dvbbs.GroupSetting(62))&"&action=OtherErr"
			End If
		End If
		If Dvbbs.GroupSetting(3)="0" And (Action = 1 Or Action = 3) Then Response.redirect "showerr.asp?ErrCodes=<li>您没有发表新主题的权限&action=OtherErr"
			If Dvbbs.GroupSetting(5)="0" And (Action = 2) Then Response.redirect "showerr.asp?ErrCodes=<li>您没有回复别人贴子的权限&action=OtherErr"
	End Function
	'得到回复或引用帖子的判断和相关信息
	Public Function Get_Re_TopicInfo()
		Get_M_Request()
		ReplyID = Request("replyid")
		If ReplyID = "" Or Not IsNumeric(ReplyID) Then ReplyID = AnnounceID
		Set Rs=Dvbbs.Execute("select PostTable from dv_topic where BoardID="&Dvbbs.BoardID&" And TopicID="&AnnounceID)
		If Not (Rs.EOF And Rs.BOF) Then
			TotalUseTable=rs(0)
		Else
			Dvbbs.AddErrCode(48)
		End If
		Set Rs=Nothing
		Dvbbs.ShowErr()
		If ReplyID = AnnounceID Then
			Set Rs=Dvbbs.Execute("select top 1 AnnounceID from "&TotalUseTable&" where RootID="&AnnounceID&" order by AnnounceID")
			If Not(Rs.BOF And Rs.EOF) Then
				ReplyID=rs(0)
			Else
				Dvbbs.AddErrCode(48)
			End If
			Set Rs=Nothing
			Dvbbs.ShowErr()
		End If
		Set Rs=Dvbbs.Execute("select b.body,b.topic,b.locktopic,b.username,b.dateandtime,b.isbest,u.lockuser,u.UserGroupID,b.UbbList from "&TotalUseTable&" b inner join [dv_user] u on b.postuserid=u.userid Where b.AnnounceID="&ReplyID)
		If Rs.EOF And Rs.BOF Then
			Dvbbs.AddErrCode(48)
		Else
			If rs("lockuser")=1 Or rs("lockuser")=2 Then
				Content=""
			ElseIf Rs("locktopic")=2 Then 
				Content=""
			ElseIf (rs("isbest")=1 and Dvbbs.GroupSetting(41)="0")Then
				Content=""
			Else
				Content=rs("body")
			End If
			Topic=Rs("topic")
			UserName=rs("username")
			DateAndTime=rs("dateandtime")
			UbbLists=Rs("UbbList")
			If UserName=Dvbbs.membername Then
				If Cint(Dvbbs.GroupSetting(4))=0 Then Dvbbs.AddErrCode(73)
			Else
				If Cint(Dvbbs.GroupSetting(2))=0 Then Dvbbs.AddErrCode(31)
			End If
		End If
		Set Rs=Nothing
		Dvbbs.ShowErr()
		If Topic<> "" Then
			Topic="回复：（" & UserName & "）" & Topic
		Else
			Topic="回复：（" & UserName & "）" & Content
		End If
		Topic=cutStr(Topic,50)
		Topic=Replace(Replace(Replace(Replace(Topic,"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")
		If Request("reply")="true" Then
			Content = reubbcode(Content)
			Content = Ubb2Html(Content)
			Content = "<DIV class=quote><B>以下是引用<i>"&UserName&"</i>在"&DateAndTime&"的发言：</B><br>"& Content & "</DIV><p>"
			Content = Server.HtmlEncode(Content)
		Else
			Content = ""
		End If
		'主题跟贴部分信息
		Dim PostUserGroup,TempStr1,TempStr2,TempStr3
		TempStr1 = Replace(template.html(7),"{$width}",Dvbbs.mainsetting(0))	'<!--post.asp##回帖帖子循环部分-->
		Set Rs=Dvbbs.Execute("Select top 10 b.UserName,b.Topic,b.dateandtime,b.body,b.AnnounceID,b.isbest,u.lockuser,u.UserGroupID,b.postbuyuser,b.ubblist,b.IsAudit,b.locktopic from "&TotalUseTable&" b inner join [dv_user] u on b.postuserid=u.userid where b.boardid="&Dvbbs.boardid&" and b.RootID="&AnnounceID&" order by b.AnnounceID desc")
		Do While Not Rs.EOF
			TempStr2 = TempStr1
			UserName = Rs("UserName")
			PostUserGroup=rs("UserGroupID")
			postbuyuser=rs("postbuyuser")
			UbbLists=Rs("UbbList")
			If bgcolor="tablebody1" Then 
				bgcolor="tablebody2"
				abgcolor="tablebody1"
			Else 
				bgcolor="tablebody1"
				abgcolor="tablebody2"
			End If
			TempStr2 = Replace(TempStr2,"{$tablebody}",bgcolor)
			TempStr2 = Replace(TempStr2,"{$username}",Dvbbs.HtmlEncode(Rs("UserName")))
			TempStr2 = Replace(TempStr2,"{$dateandtime}",Rs("DateAndTime"))
			
			If Rs("lockuser")=2 or Rs("locktopic")=2 Then
				TempStr2 = Replace(TempStr2,"{$body}",template.Strings(10))
			ElseIf Rs("lockuser")=1 Then
				TempStr2 = Replace(TempStr2,"{$body}",template.Strings(11))
			ElseIf Rs("isbest")=1 and Dvbbs.GroupSetting(41)="0" Then
				TempStr2 = Replace(TempStr2,"{$body}",template.Strings(12))
			Else
				If InStr(Ubblists,",39,") > 0  Then
				TempStr2 = Replace(TempStr2,"{$body}",dv_ubb.Dv_UbbCode(Rs("body"),PostUserGroup,1,0))
				Else
				TempStr2 = Replace(TempStr2,"{$body}",dv_ubb.Dv_UbbCode(Rs("body"),PostUserGroup,1,1))
				End If
			End If
			TempStr2 = Replace(TempStr2,"{$topic}",Dvbbs.HtmlEncode(Rs("Topic")))
			TempStr3 = TempStr3 & TempStr2
		Rs.MoveNext 
		Loop
		Rs.close
		Set Rs=Nothing
		Get_Re_TopicInfo = TempStr3
	End Function
	'取得编辑贴页面信息
	Public Function Get_Edit_TopicInfo()
		Get_M_Request()
		ReplyID = Request("replyid")
		If ReplyID = "" Or Not IsNumeric(ReplyID) Then Dvbbs.AddErrCode(30)
		Dvbbs.ShowErr()
		ReplyID = Clng(ReplyID)
		Set Rs=Dvbbs.Execute("select PostTable,TopicMode from dv_topic where TopicID="&AnnounceID)
		If Rs.Eof And Rs.Bof Then
			Dvbbs.AddErrCode(48)
		Else
			TotalUseTable = Rs(0)
			MyTopicMode = Rs(1)
			Rem 旧帖的主题模式值可能为空，则需要加入判断。2004-5-6 Dvbbs.YangZheng
			If Isnull(MyTopicMode) Or MyTopicMode = "" Then MyTopicMode = 0
		End If
		Rs.close
		Get_Edit_PermissionInfo()

		If Content<>"" then
			Dim re
			Set re=new RegExp
			re.IgnoreCase =true
			re.Global=True
			re.Pattern=vbNewLine&"<div align=right><font color=#000066>(.|\n)*<\/font><\/div>"
			Content=re.Replace(Content,"")
			re.Pattern=vbNewLine&"\[align=right\]\[color=#000066\](.|\n)*\[\/color\]\[\/align\]"
			Content=re.Replace(Content,"")
			re.Pattern="<div align=right><font color=#000066>(.|\n)*<\/font><\/div>"
			Content=re.Replace(Content,"")
			re.Pattern="\[align=right\]\[color=#000066\](.|\n)*\[\/color\]\[\/align\]"
			Content=re.Replace(Content,"")
			're.Pattern="\[i\](.*)\[\/i\]"
			'Content=re.Replace(Content,"$1")
			set re=Nothing
			Content=Ubb2Html(Content)
			Content=Server.htmlencode(Content)
		End If
		
	End Function

	'判断用户是否有编辑权限且提取相关信息
	Public Function Get_Edit_PermissionInfo()
		Dim old_user
		If Action = 4 Then
		Set Rs=Dvbbs.Execute("select b.username,b.topic,b.body,b.dateandtime,u.UserGroupID,b.signflag,b.emailflag,b.UbbList,b.Expression from "&TotalUseTable&" b,[dv_user] u where b.postuserid=u.userid and b.RootID="&AnnounceID&" and b.AnnounceID="&replyID)
		Else
		Set Rs=Dvbbs.Execute("select b.username,b.topic,b.body,b.dateandtime,u.UserGroupID,b.signflag,b.emailflag,b.UbbList,b.Expression from "&TotalUseTable&" b,[dv_user] u where b.postuserid=u.userid and b.RootID="&RootID&" and b.AnnounceID="&AnnounceID)
		End If
		If Rs.Eof And Rs.Bof Then
			Dvbbs.AddErrCode(48)
		Else
			Expression=Rs("Expression")
			If Action = 4 Then
				signflag=Rs("signflag")
				mailflag=Rs("emailflag")
				Topic=rs("topic")
				If Topic<>"" Then Topic = Server.HtmlEncode(Topic)
				topic=Replace(topic,"amp;","")
				Content=rs("body")
				old_user=rs("username")
				UbbLists=rs("UbbList")
			Else
				If Clng(Dvbbs.forum_setting(50))>0 then
					If Datediff("s",rs("dateandtime"),Now())>Clng(Dvbbs.forum_setting(50))*60 then
						Content = Content+chr(13)+chr(10)+char_changed+chr(13)
					Else
						Content = Content
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
	'只读，获得回复隐含Input模板
	Public Property Get Re_HiddenInput()
		Re_HiddenInput = template.html(4)
		Re_HiddenInput = Replace(Re_HiddenInput,"{$announceid}",AnnounceID)
		Re_HiddenInput = Replace(Re_HiddenInput,"{$replyid}",ReplyID)
	End Property
	'只读，获得编辑隐含Input模板
	Public Property Get Edit_HiddenInput()
		Edit_HiddenInput = template.html(5)
		Edit_HiddenInput = Replace(Edit_HiddenInput,"{$announceid}",AnnounceID)
		Edit_HiddenInput = Replace(Edit_HiddenInput,"{$replyid}",ReplyID)
	End Property
	'只读，获得上传表单模板
	Public Property Get Temp_FileUpload()
		Dim TempArray,TempStr1
		Temp_FileUpload = template.html(2)
		TempArray = Split(Dvbbs.Board_Setting(19),"|")
		For i = 0 To Ubound(TempArray)
			TempStr1 = TempStr1 & "<div class=menuitems><a href=#>"&TempArray(i)&"</a></div>"
		Next
		Temp_FileUpload = Replace(Temp_FileUpload,"{$uploadlist}",TempStr1)
	End Property
	'只读，获得UBB模板
	Public Property Get Temp_UBB()
		Dim TempArray
		Temp_UBB = template.html(3)
		TempArray = Split(template.html(9),"|")
		For i = 1 To Ubound(TempArray)
			Temp_UBB = Replace(Temp_UBB,"{$ubb"&i&"}",TempArray(0) & TempArray(i))
		Next
	End Property
	'只读，获得UBB——HTML编辑器模板
	Public Property Get Temp_UBBHTML()
		Dim TempArray
		Temp_UBBHTML = template.html(11)
		Temp_UBBHTML=Replace(Temp_UBBHTML,"{$old_oldToolbars}",Temp_UBB)
	End Property
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
	re.Pattern="(<DIV class=quote><b>以下是引用(.|\n)*<\/div>)"
	strContent=re.Replace(strContent,"")
	re.Pattern="<div align=right><font color=#000066>(.|\n)*<\/font><\/div>"
	strContent=re.Replace(strContent,"")
	re.Pattern="\[align=right\]\[color=#000066\](.|\n)*\[\/color\]\[\/align\]"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\[QUOTE\])(.|\n)*(\[\/QUOTE\])"
	strContent=re.Replace(strContent,"$2")
	re.Pattern="(\[point=*([0-9]*)\])(.|\n)*(\[\/point\])"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\[post=*([0-9]*)\])(.|\n)*(\[\/post\])"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\[power=*([0-9]*)\])(.|\n)*(\[\/power\])"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\[usercp=*([0-9]*)\])(.|\n)*(\[\/usercp\])"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\[money=*([0-9]*)\])(.|\n)*(\[\/money\])"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\[replyview\])(.|\n)*(\[\/replyview\])"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\[usemoney=*([0-9]*)\])(.|\n)*(\[\/usemoney\])"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\[UserName=(.[^\[]*)\])(.|\n)*(\[\/UserName\])"
	strContent=re.Replace(strContent,"")
	re.Pattern="  "
	strContent=re.Replace(strContent,"&nbsp;&nbsp;")
	re.Pattern="<I><\/I>"
	strContent=re.Replace(strContent,"")
	set re=Nothing
	reUBBCode=strContent
End Function

'编辑时用（对旧数据兼容）
Function Ubb2Html(str)
	If Str<>"" And Not IsNull(Str) Then
		Dim re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(>)("&vbNewLine&")(<)"
		Str=re.Replace(Str,"$1$3")
		re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
		Str=re.Replace(Str,"$1$3") 
		re.Pattern=vbNewLine
		Str=re.Replace(Str,"<br>")
		re.Pattern="(  )"
		Str=re.Replace(Str,"&nbsp;&nbsp;")
		re.Pattern="<I><\/I>"
		Str=re.Replace(Str,"")
		re.Pattern="(	)"
		Str=re.Replace(Str,"&nbsp;")
		re.Pattern="<(\w+)(&nbsp;)+([^>]*)>"
		Str = re.Replace(Str,"<$1 $3>")
		If Request("reply")="true" Then
			re.Pattern="(<DIV class=quote><b>以下是引用(.|\n)*<\/div>)"
			Str=re.Replace(Str,"")
			re.Pattern="(\[quote\]<b>以下是引用(.|\n)*\[\/quote\])"
			Str=re.Replace(Str,"")
			re.Pattern="(\[quote\]\[b\]以下是引用(.|\n)*\[\/quote\])"
			Str=re.Replace(Str,"")
		End If
		Set Re=Nothing 
		Ubb2Html = Str
	Else
		Ubb2Html = ""
	End If
End Function
%>