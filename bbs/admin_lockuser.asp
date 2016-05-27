<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<!-- #include file="inc/GroupPermission.asp" -->
<%
Dvbbs.Loadtemplates("")
Dim username
Dim locktype
Dim ip,BoardID
Dim TotalUseTable
Dim AdminUserPer
Dim UpdateBoardID,i,Rs,Sql
AdminUserPer=false
If (Dvbbs.master or Dvbbs.boardmaster or Dvbbs.superboardmaster) and Cint(Dvbbs.GroupSetting(42))=1 Then
	AdminUserPer=True 
Else 
	AdminUserPer=False
End If
If Dvbbs.UserGroupID > 3 And CInt(Dvbbs.GroupSetting(42))=1 Then
	AdminUserPer=True
End If 
If Dvbbs.FoundUserPer And CInt(Dvbbs.GroupSetting(42))=1 Then
	AdminUserPer=True
ElseIf Dvbbs.FoundUserPer and Cint(Dvbbs.GroupSetting(42))=0 Then
	AdminUserPer=False
End If
Dim userid
ip=Dvbbs.UserTrueIP
Dvbbs.stats="管理用户"
Response.Write "<link rel=""stylesheet"" href=""forum_admin.css"" type=""text/css"">"
Dvbbs.nav()
If request("name")="" Then
	Response.redirect "showerr.asp?ErrCodes=<li>请指定所操作的用户！&action=OtherErr"
Else
	username=Dvbbs.CheckStr(request("name"))
End If
Dvbbs.Head_Var 2,0,"",""
Dvbbs.ShowErr()
If request("userid")<> "" Then
	userid=Dvbbs.CheckStr(request("userid"))
	If Not isnumeric(userid) Then Response.redirect "showerr.asp?ErrCodes=<li>非法的参数。&action=OtherErr"
Else
	Set Rs=Dvbbs.Execute("SELECT UserID FROM [Dv_User] WHERE Username = '"&Username&"' ")
	If Not (Rs.eof And Rs.bof) Then
		UserID=Rs(0)
	Else
		UserID=0
	End If
	Set Rs=Nothing
End If 
If Not Dvbbs.ChkPost() And Request("action") <> "" Then
	Response.redirect "showerr.asp?ErrCodes=<li>您不要从外部提交数据&action=OtherErr"
End If
If request("action")="power" Then
	Call Poweruser()
ElseIf request("action")="DelTopic" then
	Call DelTopic()
ElseIf request("action")="getpermission" then
	Call boardlist()
ElseIf request("action")="userBoardPermission" then
	Call GetUserPermission()
ElseIf request("action")="saveuserpermission" then
	Call saveuserpermission()
ElseIf request("action")="DelUserReply" then
	Call DelUserReply()
Else
	Call lockuser()
End If

Dvbbs.activeonline()
Dvbbs.footer()

Sub lockuser()
	Dim canlockuser
	canlockuser=false
	if (Dvbbs.master or Dvbbs.boardmaster or Dvbbs.superboardmaster) and Cint(Dvbbs.GroupSetting(28))=1 Then
		canlockuser=True 
	Else
		canlockuser=False 
	End If
	If Dvbbs.UserGroupID > 3 And CInt(Dvbbs.GroupSetting(28))=1 Then canlockuser=True
	If Dvbbs.FoundUserPer And  Cint(Dvbbs.GroupSetting(28))=1 Then
		canlockuser=True
	ElseIf Dvbbs.FoundUserPer and Cint(Dvbbs.GroupSetting(28))=0 Then 
		canlockuser=False 
	End If 

	If Not canlockuser then Response.redirect "showerr.asp?ErrCodes=<li>您没有权限执行此操作。&action=OtherErr"

	If request("action")="lock_1" Then
		Dvbbs.Execute("update [dv_user] set LockUser=1 where userid="&userid&" and UserGroupID > 1")
		locktype="锁定"
		sql="insert into Dv_log (l_touser,l_username,l_content,l_ip,l_type) values ('"&username&"','"&Dvbbs.membername&"','用户操作："&locktype& "','"&ip&"',6)"
		Dvbbs.Execute(sql)
	ElseIf request("action")="lock_2" then
		Dvbbs.Execute("update [dv_user] set LockUser=2 where userid="&userid&" and UserGroupID>1")
		locktype="屏蔽"
		sql="insert into Dv_log (l_touser,l_username,l_content,l_ip,l_type) values ('"&username&"','"&Dvbbs.membername&"','用户操作："&locktype& "','"&ip&"',6)"
		Dvbbs.Execute(SQL)
	ElseIf request("action")="lock_3" then
		Dvbbs.Execute("update [dv_user] set LockUser=0 where userid="&userid&" and UserGroupID>1")
		locktype="解锁"
		sql="insert into Dv_log (l_touser,l_username,l_content,l_ip,l_type) values ('"&username&"','"&Dvbbs.membername&"','用户操作："&locktype& "','"&ip&"',6)"
		Dvbbs.Execute(SQL)
	Else
		Response.redirect "showerr.asp?ErrCodes=<li>请指定正确的参数！&action=OtherErr"
	End If

	Dvbbs.Dvbbs_suc("<li>您选择的用户已经"&locktype&"。您的操作已经记录在案。")
End Sub

Sub Poweruser()
	Dim title,content
	Dim canlockuser
	canlockuser=false
	If (Dvbbs.master or Dvbbs.boardmaster or Dvbbs.superboardmaster) And CInt(Dvbbs.GroupSetting(43))=1 Then
		canlockuser=True 
	Else 
		canlockuser=False 
	End If 
	If Dvbbs.UserGroupID > 3 And  Cint(Dvbbs.GroupSetting(43))=1 Then canlockuser=True 
	If Dvbbs.FoundUserPer And  CInt(Dvbbs.GroupSetting(43))=1 Then
		canlockuser=True 
	ElseIf Dvbbs.FoundUserPer And CInt(Dvbbs.GroupSetting(43))=0 Then
		canlockuser=False
	End If 

	If Not canlockuser Then Response.redirect "showerr.asp?ErrCodes=<li>您没有权限执行此操作。&action=OtherErr"
	If request("checked")="yes" Then
		Dim doWealth,douserEP,douserCP,douserPower
		Dim doWealthMsg,douserEPMsg,douserCPMsg,douserPowerMsg,allMsg
		If Not isnumeric(request("doWealth")) or request("doWealth")="0" or request("doWealth")="" Then 
			doWealth=0
			doWealthMsg=""
		Else
			doWealth=request("doWealth")
			doWealthMsg="金钱" & request("doWealth") & "，"
		End If

		If Not IsNumeric(request("douserEP")) or request("douserEP")="0" or request("douserEP")="" Then
			douserEP=0
			douserEPMsg=""
		Else 
			douserEP=request("douserEP")
			douserEPMsg="经验" & request("douserEP") & "，"
		End If 

		If Not  IsNumeric(request("douserCP")) or request("douserCP")="0" or request("douserCP")="" Then
			douserCP=0
			douserCPMsg=""
		Else
			douserCP=request("douserCP")
			douserCPMsg="魅力" & request("douserCP") & "，"
		End If

		If Not IsNumeric(request("douserPower")) or request("douserPower")="0" or request("douserPower")="" Then
			douserPower=0
			douserPowerMsg=""
		Else
			douserPower=request("douserPower")
			douserPowerMsg="威望" & request("douserPower")
		End If

		If doWealthMsg="" and douserEPMsg="" and douserCPMsg="" and douserPowerMsg="" Then
			allmsg="没有对用户进行分值操作"
		Else
			allmsg="用户操作：" & doWealthMsg & douserEPMsg & douserCPMsg & douserPowerMsg
		End If
		'Response.Write allmsg
		'response.end

		title=request.form("title")
		content=request.form("content")
		content="原因：" & title & content
		if request.form("title")="" and request.form("content")="" then Response.redirect "showerr.asp?ErrCodes=<li>请写明操作原因。&action=OtherErr"

		sql="insert into Dv_log (l_touser,l_username,l_content,l_ip,l_type) values ('"&username&"','"&Dvbbs.membername&"','用户操作："&content& "，"&allmsg&"','"&ip&"',5)"
		Dvbbs.Execute(sql)
		If allmsg<>"" Then
			Dvbbs.Execute("update [dv_user] set userWealth=userWealth+"&doWealth&",userCP=userCP+"&douserCP&",userEP=userEP+"&douserEP&",userPower=userPower+"&douserPower&" where userid="&userid&"")
		End If
		locktype="成功操作"
		Dvbbs.Dvbbs_suc("<li>您选择的用户已经"&locktype&"。<li>您的操作已经记录。")
	Else
%>
<FORM METHOD=POST ACTION="admin_lockuser.asp?action=power">
<table style="width:70%" cellspacing="1" cellpadding="3" align="center" class=tableborder1>
  <tr> 
    <th height=24>论坛管理中心－－您要进行的操作是奖励用户</th>
  </tr>   
  <tr> 
    <td class=tablebody1 height=24><b>
      操作理由</b>：  
	  <select name="title" size=1>
<option value="">自定义</option>
<option value="多次发表好文章">多次发表好文章</option>
<option value="对社区建设有贡献">对社区建设有贡献</option>
<option value="多次发表灌水帖子">多次发表灌水帖子</option>
<option value="多次发表广告帖子">多次发表广告帖子</option>
	  </select>
	  <input type="text" name="content" size=50>  *</td>
  </tr>   
  <tr> 
    <td class=tablebody1 height=24><b>
      用户操作</b>：  金钱
	<select name="doWealth" size=1>

<%for i=-50 to 50%>
<option value="<%=i%>" <%if cint(i)=cint(0) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;魅力
	<select name="douserCP" size=1>

<%for i=-50 to 50%>
<option value="<%=i%>" <%if cint(i)=cint(0) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;经验
	<select name="douserEP" size=1>

<%for i=-50 to 50%>
<option value="<%=i%>" <%if cint(i)=cint(0) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>&nbsp;威望
	<select name="douserPower" size=1>

<%for i=-5 to 5%>
<option value="<%=i%>" <%if cint(i)=cint(0) then%>selected<%end if%>><%=i%></option>
<%next%>
	</select>
  *</td>
  </tr> 
<input type=hidden value="yes" name="checked">
<input type=hidden value="<%=userid%>" name="userid">
<input type=hidden value="<%=username%>" name="name">
  <tr> 
    <td class=tablebody2 height=24>
      请慎重使用管理员的管理职能，管理员所有操作将被记录&nbsp;<input type="Submit" name=Submit value="确认操作"></td>
  </tr>   
</table>
</FORM>
<%
end if
End Sub

'删除用户1－10天内主题
Sub Deltopic()
	Dim DelDate, Suserid
	Dim Todaynum
	Dim Trs, ii, SplitBoardID
	Dim j
	DelDate = Dvbbs.checkstr(request.form("delTopicDate"))
	suserid = Dvbbs.checkstr(request.form("SetUserID"))
	Dim Canlockuser
	Canlockuser = False
	If (Dvbbs.Master Or Dvbbs.Boardmaster Or Dvbbs.Superboardmaster) And CInt(Dvbbs.GroupSetting(29)) = 1 Then
		Canlockuser = True 
	Else
		Canlockuser = False 
	End If 
	If Dvbbs.UserGroupID >3 And CInt(Dvbbs.GroupSetting(29))=1 Then canlockuser=True 
	If Dvbbs.FoundUserPer And CInt(Dvbbs.GroupSetting(29))=1 Then
		Canlockuser = True
	ElseIf Dvbbs.FoundUserPer And CInt(Dvbbs.GroupSetting(29))=0 Then 
		Canlockuser = False
	End If

	If Not canlockuser Then Response.redirect "showerr.asp?ErrCodes=<li>您没有权限执行此操作。&action=OtherErr"
	if DelDate="" or Not IsNumeric(DelDate) Then Response.redirect "showerr.asp?ErrCodes=<li>非法的参数。&action=OtherErr"
	If suserid="" or not isnumeric(suserid) then Response.redirect "showerr.asp?ErrCodes=<li>非法的参数。&action=OtherErr"
	'删除帖子表中数据
	'更新用户帖子数，用户删除帖子数为该帖所有回复数，回复人的删除帖子数不更新
	ii = 0
	REM 查询加入TOPICID排序避免有百万主题时超时 2004-5-22 Dv.Yz
	If IsSqlDataBase=1 Then
		Set Rs = Dvbbs.Execute("SELECT TopicID, Child, BoardID, PostTable, Dateandtime FROM [Dv_Topic] WHERE (NOT Boardid IN (444, 777)) AND PostUserID = " & Suserid & " AND DateDiff(d, DateAndTime, " & SqlNowString & ") < " & DelDate & " ORDER BY TopicID")
	Else
		Set Rs = Dvbbs.Execute("SELECT TopicID, Child, BoardID, PostTable, dateandtime FROM [Dv_Topic] WHERE (NOT boardid IN (444, 777)) AND PostUserID = " & Suserid & " AND DateDiff('d', DateAndTime, " & SqlNowString & ") < " & DelDate & " ORDER BY TopicID")
	End If
	If Not Rs.EOF Then
		Sql=Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		For j = 0 To Ubound(Sql,2)
			ii = ii + 1
			TotalUseTable = Sql(3,j)
			'更新发贴用户数据
			Dvbbs.Execute("UPDATE [Dv_User] SET Userpost = Userpost - 1, UserDel = UserDel - "&Sql(1,j)&"-1 WHERE userid="&suserid)
			'更新主题状态
			Dvbbs.Execute("UPDATE [Dv_topic] SET locktopic = boardid, boardid = 444 WHERE topicid = "&Sql(0,j))
			'更新帖子状态
			Dvbbs.Execute("UPDATE "&TotalUseTable&" SET locktopic = boardid, boardid=444 WHERE rootid = "&Sql(0,j))
			If DateDiff("d",Sql(4,j),Now())=0 Then
				todaynum=Sql(1,j)
			Else
				todaynum=0
			End If
			'得到该帖子所属论坛的上级论坛ID
			Set Trs=Dvbbs.Execute("SELECT ParentStr, ParentID FROM Dv_board WHERE boardid ="&Sql(2,j))
			If Trs(1)=0 Then
				UpdateBoardID=Sql(2,j)
			Else
				UpdateBoardID=trs(0) & "," & Sql(2,j)
			End If
			'更新版面数据
			Dvbbs.Execute("update Dv_board set postnum=postnum-"&Sql(1,j)&"-1,topicNum=topicNum-1,todayNum=todayNum-"&todayNum&" where boardid in ("&UpdateBoardID&")")
			'更新版面缓存
			SplitBoardID=Split(UpdateBoardID,",")
			For i=0 To Ubound(SplitBoardID)
				Dvbbs.ReloadBoardCache SplitBoardID(i),-(Sql(1,j)+1),9,1
				Dvbbs.ReloadBoardCache SplitBoardID(i),-1,10,1
				Dvbbs.ReloadBoardCache SplitBoardID(i),-todayNum,12,1
			Next
			'更新总数据
			Dvbbs.Execute("update dv_setup set Forum_TodayNum=Forum_todayNum-"&todaynum&",Forum_postNum=Forum_postNum-"&Sql(1,j)&"-1,Forum_TopicNum=Forum_topicNum-1")
			'更新总设置缓存
			Dvbbs.ReloadSetupCache CLng(Dvbbs.CacheData(7,0))-1,7
			Dvbbs.ReloadSetupCache CLng(Dvbbs.CacheData(8,0))-Sql(1,j)-1,8
			Dvbbs.ReloadSetupCache CLng(Dvbbs.CacheData(9,0))-todaynum,9
			'更新最后回复数据
			LastCount Sql(2,j),Sql(0,j),0,1
		Next
	End If
	Set Rs = Nothing 
	set Trs = Nothing
	Sql = "INSERT INTO Dv_Log (l_touser,l_username,l_content,l_ip,l_type) values ('"&username&"','"&Dvbbs.membername&"','用户操作：删除该用户主题 "&ii&" 条。','"&ip&"',6)"
	Dvbbs.Execute(SQL)
	Dvbbs.Dvbbs_suc("<li>删除该用户主题 "&ii&" 条。<li>该用户删除帖子数为所删帖所有回复数，回复人的删除帖子数不更新")
End Sub

'删除某用户指定日期内所有回复贴
Sub DelUserReply()
	Dim DelDate,suserid,DelBBS
	Dim todaynum
	Dim trs,ii,SplitBoardID
	Dim j
	DelDate=Dvbbs.checkstr(request.form("delTopicDate"))
	suserid=Dvbbs.checkstr(request.form("SetUserID"))
	DelBBS=Left(Replace(request.form("DelBBS"),"'",""),8)
	Dim canlockuser
	canlockuser=false
	if (Dvbbs.master or Dvbbs.boardmaster or dvbbs.superboardmaster) and Cint(Dvbbs.GroupSetting(29))=1 Then
		canlockuser=True 
	Else
		canlockuser=False 
	End If 
	If Dvbbs.UserGroupID >3 And Cint(Dvbbs.GroupSetting(29))=1 Then canlockuser=True
	If Dvbbs.FoundUserPer and Cint(Dvbbs.GroupSetting(29))=1 Then
		canlockuser=True 
	ElseIf Dvbbs.FoundUserPer and Cint(Dvbbs.GroupSetting(29))=0 then
		canlockuser=False 
	End If 

	If  Not canlockuser Then Response.redirect "showerr.asp?ErrCodes=<li>您没有权限执行此操作。&action=OtherErr"
	If DelDate="" or not isnumeric(DelDate) Then Response.redirect "showerr.asp?ErrCodes=<li>非法的参数。&action=OtherErr"
	If suserid="" or not isnumeric(suserid) Then Response.redirect "showerr.asp?ErrCodes=<li>非法的参数。&action=OtherErr"
	'删除帖子表中数据
	ii=0
	If IsSqlDataBase=1 then
		Set Rs=Dvbbs.Execute("SELECT Announceid, rootid, boardid, dateandtime FROM "&DelBBS&" WHERE (NOT boardid IN (444, 777)) AND PostUserID = "&suserid&"AND NOT ParentID = 0 AND Datediff(d, Dateandtime, "&SqlNowString&") < "&DelDate)
	Else
		Set Rs=Dvbbs.Execute("SELECT Announceid, rootid, boardid, dateandtime FROM "&DelBBS&" WHERE (NOT boardid IN (444, 777)) AND PostUserID = "&suserid&" AND NOT ParentID = 0 AND Datediff('d', Dateandtime, "&SqlNowString&") < "&DelDate)
	End if
	If Not Rs.EOF Then
		Sql=Rs.GetRows(-1)
		Rs.Close:Set Rs=Nothing
		For j=0 To Ubound(Sql,2)
			ii=ii+1
			'更新发贴用户数据
			Dvbbs.Execute("UPDATE [dv_user] SET userpost = userpost - 1, UserDel = UserDel - 1, UserWealth = UserWealth - "&Dvbbs.Forum_user(3)&", userEP = userEP - "&Dvbbs.Forum_user(8)&", userCP = userCP - "&Dvbbs.Forum_user(13)&" WHERE userid = "&suserid)
			Dvbbs.Execute("UPDATE "&DelBBS&" SET locktopic = boardid, boardid = 444 WHERE Announceid = "&Sql(0,j))
			isEndReply Sql(1,j),Sql(0,j),1

			If DateDiff("d",Sql(3,j),Now())=0 Then
				todaynum=1
			Else
				todaynum=0
			End If
			'得到该帖子所属论坛的上级论坛ID
			Set trs=Dvbbs.Execute("SELECT ParentStr, ParentID FROM dv_board WHERE boardid = "&Sql(2,j))
			If trs(1)=0 Then
				UpdateBoardID=Sql(2,j)
			Else
				UpdateBoardID=trs(0) & "," & Sql(2,j)
			End If
			'更新版面数据
			Dvbbs.Execute("UPDATE dv_board SET postnum = postnum - 1, todayNum = todayNum - "&todayNum&" WHERE boardid IN ("&UpdateBoardID&")")
			'更新版面缓存
			SplitBoardID=Split(UpdateBoardID,",")
			For i=0 To Ubound(SplitBoardID)
				Dvbbs.ReloadBoardCache SplitBoardID(i),-1,9,1
				Dvbbs.ReloadBoardCache SplitBoardID(i),-todayNum,12,1
			Next
			'更新总数据
			Dvbbs.Execute("UPDATE dv_setup SET Forum_TodayNum = Forum_todayNum - "&todaynum&", Forum_postNum = Forum_postNum - 1")
			'更新总设置缓存
			Dvbbs.ReloadSetupCache CLng(Dvbbs.CacheData(8,0))-1,8
			Dvbbs.ReloadSetupCache CLng(Dvbbs.CacheData(9,0))-todaynum,9
			'更新最后回复数据
			LastCount Sql(2,j),Sql(1,j),Sql(0,j),2
		Next
	End if
	Set Rs=Nothing
	Set Trs=Nothing
	Sql="insert into Dv_log (l_touser,l_username,l_content,l_ip,l_type) values ('"&username&"','"&Dvbbs.membername&"','用户操作：删除该用户回复 "&ii&" 条。','"&ip&"',6)"
	Dvbbs.Execute(SQL)
	Dvbbs.Dvbbs_suc("<li>删除该用户回复 "&ii&" 条。")
End Sub

Sub boardlist()
	Dim trs
	Dim dispboard
	If Not AdminUserPer Then
		Response.redirect "showerr.asp?ErrCodes=<li>您没有权限执行此操作。&action=OtherErr"
	Else
		dispboard=True 
	End If
	If Not IsNumeric (request("userid")) then Response.redirect "showerr.asp?ErrCodes=<li>错误的用户参数。&action=OtherErr"

	Response.Write "<table style=""width:70%"" cellspacing=1 cellpadding=3 align=center class=tableborder1>"
	Response.Write "<tr><th colspan=6 height=23 align=left>编辑"&request("username")&"论坛权限（红色表示该用户在该版面有自定义权限）</th></tr>"
	Response.Write "<tr><td colspan=6 class=tablebody1 height=25><a href=?action=userBoardPermission&boardid=0&userid="&request("userid")&">编辑该用户在其它页面的权限</a>（主要针对短信部分设置）</td></tr>"
	'----------------------boardinfo--------------------
	Response.Write "<tr><td colspan=6 class=tablebody1><B>点击论坛名称进入编辑状态</B><BR>"
	Dim reBoard_Setting,FBoardMaster
	FBoardMaster=False 
	sql="select * from dv_board order by rootid,orders"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	Do while Not rs.eof
	reBoard_Setting=split(rs("Board_Setting"),",")
		if rs("parentID")>0 and rs("depth")>0 then
			set trs=Dvbbs.Execute("select boardmaster from dv_board where boardid in ("&rs("ParentStr")&") order by orders")
			Do While Not trs.EOF 
			If Cint(reBoard_Setting(40))=1 and not FBoardMaster Then 
				If InStr("|"&trs(0)&"|","|"&Dvbbs.membername&"|")>0 Then 
					FBoardMaster=True 
				Else
					FBoardMaster=False 
				End If
			End  If 
			i=i+1
			If FBoardMaster Then  Exit Do
			if i>9 Then Exit Do
			trs.MoveNext 
			Loop 
		End If 
		If Dvbbs.boardmaster And InStr("|"&rs("BoardMaster")&"|","|"&Dvbbs.membername&"|")=0 then dispboard=False
		If FBoardMaster Then dispboard=True
		If Dvbbs.master or Dvbbs.superboardmaster Then dispboard=True 
		if dispboard Then
			If Rs("depth")>0 Then 
				for i=1 to rs("depth")
					Response.Write "&nbsp;"
				Next
			End If
			If rs("child")>0 Then
				Response.Write "<img src=""skins/default/plus.gif"">"
			Else
				Response.Write "<img src=""skins/default/nofollow.gif"">"
			End If 
%>
<a href="?action=userBoardPermission&boardid=<%=rs("boardid")%>&userid=<%=request("userid")%>&name=<%=request("name")%>">
<%
			set trs=Dvbbs.Execute("select uc_userid from dv_UserAccess where uc_Boardid="&rs("boardid")&" and uc_userid="&request("userid"))
			if not (trs.eof and trs.bof) then Response.Write "<font color=red>"
			if rs("parentid")=0 then Response.Write "<b>"
			Response.Write rs("boardtype")
			if rs("parentid")=0 then Response.Write "</b>"
			if rs("child")>0 then Response.Write "("&rs("child")&")"
		end if
		Response.Write "</font></a><br>"
		FBoardMaster=false
		dispboard=false
	rs.MoveNext
	loop
	set trs=nothing
	set rs=nothing
	Response.Write "</td></tr></table>"

End Sub

Sub GetUserPermission()
	If Not AdminUserPer Then Response.redirect "showerr.asp?ErrCodes=<li>您没有权限执行此操作。&action=OtherErr"
	Dim usertitle
	If Not IsNumeric(request("userid")) then Response.redirect "showerr.asp?ErrCodes=<li>错误的用户参数。&action=OtherErr"
	If Not IsNumeric(request("boardid")) then Response.redirect "showerr.asp?ErrCodes=<li>错误的用户参数。&action=OtherErr"
	Dim UserGroupID,membername,boardtype
	Response.Write "<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>"
	set rs=Dvbbs.Execute("select u.UserGroupID,ug.title,u.username from [dv_user] u inner join dv_UserGroups UG on u.userGroupID=ug.userGroupID where u.userid="&request("userid"))
	UserGroupID=rs(0)
	usertitle=rs(1)
	membername=rs(2)
	set rs=Dvbbs.Execute("select boardtype from dv_board where boardid="&request("boardid"))
	if rs.eof and rs.bof then
		boardtype="论坛其他页面"
	else
		boardtype=rs(0)
	end if
	Response.Write "<tr><th colspan=6 height=23 align=left>编辑 "&membername&" 在 "&boardtype&" 权限</th></tr>"
	Response.Write "<tr><td colspan=6 height=25 class=tablebody1>注意：该用户属于 <B>"&usertitle&"</B> 用户组中，如果您设置了他的自定义权限，则该用户权限将以自定义权限为主</td></tr>"

	Dim reGroupSetting
	Dim FoundGroup,FoundUserPermission,FoundGroupPermission
	FoundGroup=false
	FoundUserPermission=false
	FoundGroupPermission=false

	set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
	If Not (rs.eof and rs.bof) Then
		reGroupSetting=rs("uc_Setting")
		FoundGroup=true
		FoundUserPermission=true
	End If 
	if not foundgroup then
		set rs=Dvbbs.Execute("select * from dv_BoardPermission where boardid="&request("boardid")&" and groupid="&usergroupid)
		if not(rs.eof and rs.bof) then
			reGroupSetting=rs("PSetting")
			FoundGroup=true
			FoundGroupPermission=true
		end if
	end if
	if not foundgroup then
		set Rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&usergroupid)
		If rs.eof And rs.bof Then
			Response.Write "未找到该用户组！"
			response.end
		Else
			FoundGroup=true
			FoundGroupPermission=true
			reGroupSetting=rs("GroupSetting")
		End If 
	End If
	
%>
<FORM METHOD=POST ACTION="?action=saveuserpermission">
<input type=hidden name="userid" value="<%=request("userid")%>">
<input type=hidden name="BoardID" value="<%=request("boardid")%>">
<input type=hidden name="username" value="<%=membername%>">
<input type=hidden name="name" value="<%=membername%>">
<tr> 
<td height="23" colspan="3" class=tablebody1><input type=radio name="isdefault" value="1" <%if FoundGroupPermission then%>checked<%end if%>><B>使用用户组默认值</B> (注意: 这将删除任何之前所做的自定义设置)</td>
</tr>
<tr> 
<td height="23" colspan="3"  class=tablebody1><input type=radio name="isdefault" value="0" <%if FoundUserPermission then%>checked<%end if%>><B>使用自定义设置</B> </td>
</tr>
<%
GroupPermission(reGroupSetting)
%>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
End Sub

Sub SaveUserPermission()
	if not AdminUserPer then Response.redirect "showerr.asp?ErrCodes=<li>您没有权限执行此操作。&action=OtherErr"
	if not isnumeric(request("userid")) then Response.redirect "showerr.asp?ErrCodes=<li>错误的用户参数。&action=OtherErr"
	if not isnumeric(request("boardid")) then Response.redirect "showerr.asp?ErrCodes=<li>错误的版面参数。&action=OtherErr"
	Dim trs
	'最后一次进行验证，主要为验证版主是否所操作的版面版主
	Dim reboard_setting,fboardmaster
	fboardmaster=false

	if Dvbbs.boardmaster and (Not Dvbbs.FoundUserPer) and (not (Dvbbs.master or Dvbbs.superboardmaster)) Then 
		set rs=Dvbbs.Execute("select boardmaster,parentid,depth,Board_Setting from dv_board where boardid="&request("boardid"))
		if not(rs.eof and rs.bof) then
			reBoard_Setting=split(rs(3),",")
			if rs("parentID")>0 and rs("depth")>0 then
				set trs=Dvbbs.Execute("select boardmaster from dv_board where boardid in ("&rs("ParentStr")&") order by orders")
				do while not trs.eof
				if Cint(reBoard_Setting(40))=1 And Not FBoardMaster Then 
					if instr("|"&trs(0)&"|","|"&Dvbbs.membername&"|")>0 Then
						FBoardMaster=True 
					else
						FBoardMaster=False
					end if
				end if
				i=i+1
				if FBoardMaster then exit do
				if i>9 then exit do
				trs.movenext
				loop
			end if
			if not fboardmaster then
				if instr("|"&rs(0)&"|","|"&membername&"|")=0 then Dvbbs.AddErrMsg "您不是该版面的版主，无权进行用户权限操作。"
			end if
		Else
			Response.redirect "showerr.asp?ErrCodes=<li>您不是该版面的版主，无权进行用户权限操作。&action=OtherErr"
		End If
	End If

	Dim myGroupSetting
	Dim IsGroupSetting,MyIsGroupSetting,FoundSetting
	myGroupSetting=GetGroupPermission
	if request("isdefault")=1 then
		Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
		Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
		FoundSetting=Rs(0)
		If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
		Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&request("boardid"))
		If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
			MyIsGroupSetting = ""
		Else
			IsGroupSetting = "," & Rs(0) & ","
			If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0","")
			IsGroupSetting = Split(IsGroupSetting,",")
			For i=1 To Ubound(IsGroupSetting)-1
				If i=1 Then
					MyIsGroupSetting = IsGroupSetting(i)
				Else
					MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
				End If
			Next
		End If
		Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&request("boardid"))
		Set Rs=Nothing
	else
		set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
		if rs.eof and rs.bof Then
			Dvbbs.Execute("insert into Dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&request("boardid")&",'"&myGroupSetting&"')")						
		else
			Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
		End If
		Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&request("boardid"))
		If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
			MyIsGroupSetting = 0
		Else
			IsGroupSetting = "," & Rs(0) & ","
			IsGroupSetting = Replace(IsGroupSetting,",0","")
			IsGroupSetting = IsGroupSetting & "0,"
			IsGroupSetting = Split(IsGroupSetting,",")
			For i=1 To Ubound(IsGroupSetting)-1
				If i=1 Then
					MyIsGroupSetting = IsGroupSetting(i)
				Else
					MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
				End If
			Next
		End If
		Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&request("boardid"))
		Set Rs=Nothing
	end If 
	sql="insert into Dv_log (l_touser,l_username,l_content,l_ip,l_type) values ('"&username&"','"&Dvbbs.membername&"','用户操作：设置权限。','"&ip&"',5)"
	Dvbbs.Execute(SQL)
	Dvbbs.Dvbbs_suc("<li>设置用户权限成功。您的操作已经记录")
End Sub

'更新指定论坛信息，取主题和最后回复作者及时间
'flag=1为根据主题ID判断，2为根据回复ID判断是否版面最后回复
Function LastCount(boardid,topicid,ireplyid,flag)
	LastCount=false
	Dim LastTopic,LastRootid,LastPostTime,LastPostUser
	Dim LastPost,uploadpic_n,Lastpostuserid,Lastid
	Dim trs
	Dim LastPostInfo
	set trs=Dvbbs.Execute("select LastPost from dv_board where boardid="&boardid)
	if not (trs.eof and trs.bof) then
		LastPostInfo=split(trs(0),"$")
		if ubound(LastPostInfo)=7 then
			if flag=1 then
				if clng(LastPostInfo(6))=topicid then LastCount=true
			elseif flag=2 then
				if clng(LastPostInfo(1))=ireplyid then LastCount=true
			end if
		end if
	end if
	if LastCount then
		set trs=Dvbbs.Execute("select top 1 T.title,b.Announceid,b.dateandtime,b.username,b.postuserid,b.rootid from "&Dvbbs.NowUseBBS&" b inner join dv_Topic T on b.rootid=T.TopicID where b.boardid="&boardid&" order by b.announceid desc")
		if not(trs.eof and trs.bof) then
			Lasttopic=replace(left(trs(0),15),"$","")
			LastRootid=trs(1)
			LastPostTime=trs(2)
			LastPostUser=trs(3)
			LastPostUserid=trs(4)
			Lastid=trs(5)
		else
			LastTopic="无"
			LastRootid=0
			LastPostTime=now()
			LastPostUser="无"
			LastPostUserid=0
			Lastid=0
		end if
		LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & BoardID
		Dim SplitUpBoardID,SplitLastPost
		SplitUpBoardID=split(UpdateBoardID,",")
		For i=0 to ubound(SplitUpBoardID)
			set trs=Dvbbs.Execute("select LastPost from dv_board where boardid="&SplitUpBoardID(i))
			if not (trs.eof and trs.bof) then
				SplitLastPost=split(trs(0),"$")
				if ubound(SplitLastPost)=7 and clng(LastRootID)<>clng(SplitLastPost(1)) then
					Dvbbs.Execute("update dv_board set LastPost='"&LastPost&"' where boardid="&SplitUpBoardID(i))
					Dvbbs.ReloadBoardCache SplitUpBoardID(i),LastPost,14,0
				end if
			end if
		Next
	End if
	set trs=Nothing
End function

'删除回复过程中判断该回复是否最后发言，如果是则更新主题LastPost数据并回复-1
'入口：
'TopicID--主题ID
'iReplyID--回复ID
'isUpdate--是否更新时间
'出口：isEndReply=true/false
Function IsEndReply(TopicID,iReplyID,isUpdate)
	isEndReply=false
	Dim trs
	Dim LastPostInfo,iTotalUseTable
	Dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
	Dim LastPost,uploadpic_n,LastPostUserID,LastID,istop
	set trs=Dvbbs.Execute("select LastPost,PostTable,istop from dv_Topic where Topicid="&Topicid)
	if not (trs.eof and trs.bof) then
		LastPostInfo=split(trs(0),"$")
		iTotalUseTable=trs(1)
		istop=trs(2)
		if ubound(LastPostInfo)=7 then
			if Clng(LastPostInfo(1))=Clng(iReplyID) then isEndReply=true
		end if
	end if
	if isEndReply and isUpdate=1 then
		set trs=Dvbbs.Execute("select top 1 topic,body,Announceid,dateandtime,username,PostUserid,rootid,boardid from "&iTotalUseTable&" where  (not boardid in (444,777)) and rootid="&TopicID&" order by Announceid desc")
		if not(trs.eof and trs.bof) then
			body=trs(1)
			LastRootid=trs(2)
			LastPostTime=trs(3)
			LastPostUser=replace(trs(4),"$","")
			LastTopic=left(replace(body,"$",""),20)
			LastPostUserID=trs(5)
			LastID=trs(6)
			BoardID=trs(7)
		else
			LastTopic="无"
			LastRootid=0
			LastPostTime=now()
			LastPostUser="无"
			LastPostUserID=0
			LastID=0
			Boardid=0
		end if
		if istop=1 then LastPostTime=dateadd("d",100,LastPostTime)
		LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & replace(left(LastTopic,20),"$","") & "$" & LastPostInfo(4) & "$" & LastPostUserID & "$" & LastID & "$" & BoardID
		Dvbbs.Execute("update dv_topic set LastPost='"&LastPost&"',child=child-1,LastPostTime='"&LastPostTime&"' where topicid="&TopicID)
	elseif isupdate=1 and not isendreply then
		Dvbbs.Execute("update dv_topic set child=child-1 where topicid="&topicid)
	end if
	set trs=nothing
End Function
%>