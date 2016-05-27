<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<!--#include file="inc/ubblist.asp"-->
<%
Head()
Server.ScriptTimeout=9999999
dim admin_flag 
admin_flag="12,19"
if not Dvbbs.master or instr(","&session("flag")&",",",12,")=0 or instr(","&session("flag")&",",",19,")=0 then
	Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	dvbbs_error()
End If
	dim tmprs
	dim body
	call main()
	Footer()

sub main()
%>
<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<tr>
<th align=left colspan=2 height=23>论坛数据处理</th>
</tr>
<tr>
<td width="20%" class="forumrow" height=25>注意事项</td>
<td width="80%" class="forumrow">下面有的操作可能将非常消耗服务器资源，而且更新时间很长，请仔细确认每一步操作后执行。</td>
</tr>
<%
	If request("action")="updat" Then
		If request("submit")="更新论坛数据" Then
			call updateboard()
		ElseIf request("submit")="修 复" Then
			call fixtopic()
		Else
			call updateall()
		End If
		If founderr Then
			response.write errmsg
		Else
			response.write body
		End If
	ElseIf request("action")="fix"  Then
		Call Fixbbs()
		If founderr Then
			response.write errmsg
		Else
			response.write body
		End If
	ElseIf request("action")="delboard" then
		if isnumeric(request("boardid")) then
		Dvbbs.Execute("update dv_topic set boardid=444 where boardid="&request("boardid"))
		for i=0 to ubound(AllPostTable)
		Dvbbs.Execute("update "&AllPostTable(i)&" set boardid=444 where boardid="&request("boardid"))
		next
		end if
		response.write "<tr><td align=left colspan=2 height=23 class=forumrow>清空论坛数据成功，请返回更新帖子数据！</td></tr>"
	elseif request("action")="updateuser" then
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>更新用户数据</th>
</tr>
<tr>
<td width="20%" class="forumrow">重新计算用户发贴</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>发贴重新计算所有用户发表帖子数量。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="重新计算用户发贴"></td>
</tr>
</form>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="forumrow" valign=top>更新用户等级</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户发贴数量和论坛的等级设置重新计算用户等级，本操作不影响等级为贵宾、版主、总版主的数据。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户等级"></td>
</tr>
</form>

<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr>
<td width="20%" class="forumrow" valign=top>更新用户金钱/经验/魅力</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户的发贴数量和论坛的相关设置重新计算用户的金钱/经验/魅力，本操作也将重新计算贵宾、版主、总版主的数据<BR>注意：不推荐用户进行本操作，本操作在数据很多的时候请尽量不要使用，并且本操作对各个版面删除帖子等所扣相应分值不做运算，只是按照发贴和总的论坛分值设置进行运算，请大家慎重操作，<font color=red>而且本项操作将重置用户因为奖励、惩罚等原因管理员对用户分值的修改。</font></td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="100" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户金钱/经验/魅力"></td>
</tr>
</FORM>
<%
	elseif request("action")="updateuserinfo" then
		if request("submit")="重新计算用户发贴" then
		call updateTopic()
		elseif request("submit")="更新用户等级" then
		call updategrade()
		else
		call updatemoney()
		end if
		if founderr then
		response.write errmsg
		else
		response.write body
		end if
	else
%>
<tr> 
<th align=left colspan=2 height=23>更新论坛数据</th>
</tr>

<form action="admin_update.asp?action=updat" method=post>
<tr>
<td width="20%" class="forumrow">更新分论坛数据</td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新论坛数据"><BR><BR>这里将重新计算每个论坛的帖子主题和回复数，今日帖子，最后回复信息等，建议每隔一段时间运行一次。</td>
</tr>
<tr>
<td width="20%" class="forumrow">更新总论坛数据</td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新论坛总数据"><BR><BR>这里将重新计算整个论坛的帖子主题和回复数，今日帖子，最后加入用户等，建议每隔一段时间运行一次。</td>
</tr>
<tr> 
<th align=left colspan=2 height=23>修复帖子(修复指定范围内帖子的最后回复数据)</th>
</tr>
<tr>
<td width="20%" class="forumrow">开始的ID号</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;帖子主题ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束的ID号</td>
<td width="80%" class="forumrow"><input type=text name="EndID" value="1000" size=10>&nbsp;将更新开始到结束ID之间的帖子数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="修 复"></td>
</tr>
</form>
<form name=Fix action="admin_update.asp?action=fix" method=post>
<tr> 
<th align=left colspan=2 height=23>修正贴子UBB标签(修复指定范围贴子UBB标签)</th>
</tr>
<tr>
<td width="20%" class="forumrow">开始的ID号</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="1" size=10>&nbsp;帖子主题ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束的ID号</td>
<td width="80%" class="forumrow"><input type=text name="EndID" value="1000" size=10>&nbsp;将更新开始到结束ID之间的帖子数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow">新老贴的标识日期</td>
<td width="80%" class="forumrow"><input type="text" name="updatedate" value="2003-12-1">(格式：YYYY-M-D) 就是论坛升级到v7.0的日期，如果不填写，一律按老贴处理</td>
</tr>
<tr>
<td width="20%" class="forumrow">去掉贴子中的HTML标记</td>
<td width="80%" class="forumrow">是 <input type="radio" name="killhtml" value="1">
  否 <input type="radio" name="killhtml" value="0" checked>&nbsp;<br>选是的话，贴子中的HTML标记将会自动被清除，有利于减少数据库的大小，但是会失去原来的HTML效果。</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="修 正"></td>
</tr>
</form>

<%
	end if
%>
</table><BR><BR>
<%
	end sub

sub updateboard()
	'先按照所有版面ID得出帖子数，然后计算各个有下属论坛的帖子总和
	Dim allarticle
	Dim alltopic
	Dim alltoday
	Dim allboard
	Dim trs,Esql,ars
	Dim Maxid
	Dim LastTopic,LastRootid,LastPostTime,LastPostUser
	Dim LastPost,uploadpic_n,Lastpostuserid,Lastid
	Dim ParentStr
	Dim C,C1,C2
	Dim reBoard_Setting,BoardTopStr,IsGroupSetting
	Dim UserAccessCount,UpGroupSetting,ii
	ii=0


	'获得要更新的总数
	If IsNumeric(request("boardid")) And request("boardid")<>"" Then
		Set Rs=Dvbbs.Execute("Select Count(*) From [Dv_board] Where BoardID="&request("boardid"))
		C1=rs(0)
		If Isnull(C1) Then C1=0
	Else
		Set Rs=Dvbbs.Execute("Select Count(*) From [Dv_board]")
		C1=rs(0)
		If Isnull(C1) Then C1=0
	End If
	Set Rs=Nothing
%>
</table><BR>
<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<tr><td colspan=2 class=forumrow>
下面开始更新论坛版面资料，共有<%=C1%>个版面需要更新
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=ffffff height=9><img src="skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>
<%
	Response.Flush

	'排序按照Child和Orders，以便先更新下级论坛的数据才循环到上级版面，这时上级版面读取的就是下级版面的最新数据
	If IsNumeric(request("boardid")) And request("boardid")<>"" Then
		Set Rs=Dvbbs.Execute("Select BoardID,BoardType,Child,ParentStr,RootID,Board_Setting,BoardTopStr,IsGroupSetting From Dv_Board Where BoardID="&Request("BoardID"))
	Else
		Call Boardchild()	'统计更新下属论坛个数 YZ-2004-2-26注
		Set Rs=Dvbbs.Execute("Select BoardID,BoardType,Child,ParentStr,RootID,Board_Setting,BoardTopStr,IsGroupSetting From Dv_Board Order by Child,RootID,Orders Desc")
	End If
	Dim SQL
	If Not Rs.EOF Then 
		SQL=Rs.GetRows(-1)
		Set Rs=Nothing
		For i=0 to UBound(SQL,2)
	'Do While Not Rs.Eof
		reBoard_Setting=Split(SQL(5,i),",")
		'所有主题和帖子
		Set Trs=Dvbbs.Execute("Select Count(*),Sum(Child) From Dv_Topic Where BoardID="&SQL(0,i))
		AllTopic=Trs(0)
		AllArticle=Trs(1)
		If IsNull(AllTopic) Then AllTopic = 0
		If IsNull(AllArticle) Then AllArticle = 0
		AllArticle = AllArticle + AllTopic
		Set Trs=Nothing
		'所有今日贴
		If IsSqlDataBase = 1 Then
			Set Trs=Dvbbs.Execute("Select Count(*) From "&Dvbbs.NowUseBBS&" Where BoardID="&SQL(0,i)&" and datediff(d,dateandtime,"&SqlNowString&")=0")
		Else
			Set Trs=Dvbbs.Execute("Select Count(*) From "&Dvbbs.NowUseBBS&" Where BoardID="&SQL(0,i)&" and datediff('d',dateandtime,"&SqlNowString&")=0")
		End If
		AllToday=Trs(0)
		Set Trs=Nothing
		If IsNull(AllToday) Then AllToday=0
		'最后回复信息
		Set Trs=Dvbbs.Execute("Select Top 1 LastPost From Dv_Topic Where BoardID="&SQL(0,i)&" Order by LastPostTime Desc")
		If Not (Trs.Eof And Trs.Bof) Then
			LastPost=Replace(Trs(0)&"","'","''")
		Else
			LastPost="无$0$"&Now()&"$无$$$$"
		End If
		Set Trs=Nothing
		'更新当前版面数据
		Dvbbs.Execute("Update [Dv_board] Set PostNum="&AllArticle&",TopicNum="&AllTopic&",TodayNum="&AllToday&",LastPost='"&LastPost&"' Where BoardID="&SQL(0,i))
		'如果当前版面有下属论坛，则更新其数据为下属论坛数据
		If SQL(2,i)>0 Then
			'帖子总数，主题总数，今日贴总数，下属版面数
			If SQL(3,i)="0" Then
				ParentStr=SQL(0,i)
				Set Trs=Dvbbs.Execute("Select Sum(PostNum),Sum(TopicNum),Sum(TodayNum),Count(*) From Dv_board Where (Not BoardID="&SQL(0,i)&") And RootID="&SQL(4,i))
			Else
				ParentStr=SQL(3,i) & "," & SQL(0,i)
				Set Trs=Dvbbs.Execute("Select Sum(PostNum),Sum(TopicNum),Sum(TodayNum),Count(*) From Dv_board Where ParentStr Like '%"&ParentStr&"%'")
			End If
			If Not (Trs.Eof And Trs.Bof) Then
				'如果该版面允许发贴，则帖子数应该是该版面贴数+下属版面帖子数
				If reBoard_Setting(43)="0" Then
					If Not IsNull(Trs(0)) Then AllArticle = Trs(0) + AllArticle
					If Not IsNull(Trs(1)) Then AllTopic = Trs(1) + AllTopic
					If Not IsNull(Trs(2)) Then AllToday = Trs(2) + AllToday
					If Not IsNull(Trs(3)) Then AllBoard = Trs(3) + AllBoard
				Else
					AllArticle=Trs(0)
					AllTopic=Trs(1)
					AllToday=Trs(2)
					AllBoard=Trs(3)
					If IsNull(AllArticle) Then AllArticle=0
					If IsNull(AllTopic) Then AllTopic=0
					If IsNull(AllToday) Then AllToday=0
					If IsNull(AllBoard) Then AllBoard=0
				End If
			End If
			Set Trs=Nothing
			'下属版块ID
			ParentStr = Sql(0,i)
			Set Trs = Dvbbs.Execute("SELECT Boardid FROM Dv_Board WHERE ParentID = "&Sql(0,i))
			If Not (Trs.Eof And Trs.Bof) Then
				Do While Not Trs.Eof
					ParentStr = ParentStr & "," & Trs(0)
					Trs.Movenext
				Loop
			End If
			Set Trs=Nothing
			'最后回复信息
			Set Trs=Dvbbs.Execute("Select Top 1 LastPost From Dv_Topic Where BoardID In ("&ParentStr&") Order by LastPostTime Desc")
			If Not (Trs.Eof And Trs.Bof) Then
				LastPost=Replace(Trs(0),"'","''")
			Else
				LastPost="无$0$"&Now()&"$无$$$$"
			End If
			'更新版面数据
			Dvbbs.Execute("Update [Dv_board] Set PostNum="&AllArticle&",TopicNum="&AllTopic&",TodayNum="&AllToday&",LastPost='"&LastPost&"' Where BoardID="&SQL(0,i))
		End If
		'更新IsGroupSetting
		IsGroupSetting=SQL(7,i)
		Set Trs=Dvbbs.Execute("Select Count(*) From Dv_UserAccess Where uc_BoardID="&SQL(0,i))
		UserAccessCount = Trs(0)
		If IsNull(UserAccessCount) Or UserAccessCount="" Then UserAccessCount=0
		If UserAccessCount>0 Then UpGroupSetting="0"
		Set Trs=Dvbbs.Execute("Select GroupID From Dv_BoardPermission Where BoardID="&SQL(0,i))
		If Not Trs.Eof Then
			Do While Not Trs.Eof
				If UpGroupSetting="" Then
					UpGroupSetting = Trs(0)
				Else
					UpGroupSetting = UpGroupSetting & "," & Trs(0)
				End If
			Trs.MoveNext
			Loop
		End If
		'更新和清理固顶贴数据(固顶和区域固顶)
		'Set Trs=Dvbbs.Execute("Select TopicID From Dv_Topic Where BoardID="&Rs(0)&" And IsTop In (1,2)")
		If Not IsNull(SQL(6,i)) And SQL(6,i)<>"" Then
		Set Trs=Dvbbs.Execute("Select TopicID,BoardID,IsTop From Dv_Topic Where TopicID In ("&SQL(6,i)&")")
		If tRs.Eof And tRs.Bof Then
			BoardTopStr=""
		Else
			Do While Not Trs.Eof
				If Trs(1)<>444 And Trs(1)<>777 And Trs(2)>0 And Trs(2)<>3 Then
					If BoardTopStr="" Then
						BoardTopStr = Trs(0)
					Else
						BoardTopStr = BoardTopStr & "," & Trs(0)
					End If
				End If
			Trs.MoveNext
			Loop
		End If
		End If
		Dvbbs.Execute("Update Dv_Board Set BoardTopStr='"&BoardTopStr&"',IsGroupSetting='"&UpGroupSetting&"' Where BoardID="&SQL(0,i))
		UserAccessCount=""
		IsGroupSetting=""
		UpGroupSetting=""
		BoardTopStr=""
		ii=ii+1
		'If (i mod 100) = 0 Then
			Response.Write "<script>img2.width=" & Fix((ii/C1) * 400) & ";" & VbCrLf
			Response.Write "txt2.innerHTML=""" & FormatNumber(ii/C1*100,4,-1) & """;" & VbCrLf
			Response.Write "img2.title=""" & SQL(0,i) & "(" & ii & ")"";</script>" & VbCrLf
			Response.Flush
		'End If
		body="<table cellpadding=0 cellspacing=0 border=0 width=95% class=tableBorder align=center><tr><td colspan=2 class=forumrow>更新论坛数据成功，"&SQL(1,i)&"共有"&AllArticle&"篇贴子，"&AllTopic&"篇主题，今日有"&AllToday&"篇帖子。</td></tr></table>"
		Response.Write body
		Response.Flush
	'Rs.MoveNext
	'Loop
	Next
	Set Trs=Nothing
	
	End If 
	body=""
	Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
	Dvbbs.ReloadAllBoardInfo()
	Dvbbs.Name="setup"
	Dvbbs.ReloadSetup
	Dvbbs.CacheData=Dvbbs.value
	Dim Forum_Boards
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
End Sub

Rem 统计下属论坛函数 2004-5-3 Dvbbs.YangZheng
Sub Boardchild()
	Dim cBoardNum, cBoardid
	Dim Trs
	Dim Bn
	Dvbbs.Execute("UPDATE Dv_Board SET Child = 0")
	Set Rs = Dvbbs.Execute("SELECT Boardid, Rootid, ParentID, Depth, Child, ParentStr FROM Dv_Board ORDER BY Boardid DESC")
	If Not (Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		For Bn = 0 To Ubound(Sql,2)
			If Isnull(Sql(4,Bn)) And Cint(Sql(3,Bn)) > 0 Then
				Dvbbs.Execute("UPDATE Dv_Board SET Child = 0 WHERE Boardid = " & Sql(0,Bn))
			End If
			If Cint(Sql(2,Bn)) = 0 And Cint(Sql(3,Bn)) = 0 Then
				Set Trs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_Board WHERE RootID = " & Sql(1,Bn))
				Cboardnum = Trs(0) - 1
				Trs.Close:Set Trs = Nothing
				If Isnull(Cboardnum) Or Cboardnum < 0 Then Cboardnum = 0
				Dvbbs.Execute("UPDATE Dv_Board SET Child = " & Cboardnum & " WHERE Boardid = " & Sql(0,Bn))
			Elseif Cint(Sql(3,Bn)) > 1 Then
				cBoardid = Split(Sql(5,Bn),",")
				For i = 1 To Ubound(cBoardid)
					Dvbbs.Execute("UPDATE Dv_Board SET Child = Child + 1 WHERE Boardid = " & cBoardid(i))
				Next
			End If
		Next
	End If
End Sub

sub updateall()
Dim AllTopNum
AllTopNum=Forum_AllTopNum()
sql="update Dv_setup set Forum_TopicNum="&topicnum()&",Forum_PostNum="&announcenum()&",Forum_TodayNum="&alltodays()&",Forum_UserNum="&allusers()&",Forum_lastUser='"&newuser()&"',Forum_AllTopNum='"&AllTopNum&"'"
Dvbbs.Execute(sql)
body="<tr><td colspan=2 class=forumrow>更新总论坛数据成功，全部论坛共有"&announcenum()&"篇贴子，"&topicnum()&"篇主题，总固顶主题共 "&UBound(split(AllTopNum,","))+1&" 篇，今日有"&alltodays()&"篇帖子，有"&allusers()&"用户，最新加入为"&newuser()&"。</td></tr>"
Dvbbs.Name="setup"
Dvbbs.ReloadSetup()
end sub

Sub fixtopic()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	exit sub
End If
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if
dim TotalUseTable,Ers
dim username,dateandtime,rootid,announceid,postuserid,lastpost,topic
'set rs=server.createobject("adodb.recordset")
'Dvbbs.Execute("update Dv_topic set PostTable='dv_bbs1'")
Dim C1
C1=clng(request.form("endid"))-clng(request.form("beginid"))
%>
</table>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<tr><td colspan=2 class=forumrow>
下面开始更新论坛帖子资料，预计本次共有<%=C1%>个帖子需要更新
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=ffffff height=9><img src="skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>

<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<%
Response.Flush
sql="select topicid,PostTable from Dv_topic where topicid>="&request.form("beginid")&" and topicid<="&request.form("endid")

set rs=Dvbbs.Execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=forumrow>已经到记录的最尾端，请结束更新！</td></tr>"
	exit sub
end if
do while not rs.eof
	sql="select top 1 username,dateandtime,topic,Announceid,PostUserID,rootid,body,boardid from "&rs(1)&" where rootid="&rs(0)&" order by Announceid desc"
	set ers=Dvbbs.Execute(sql)
	if not (ers.eof and ers.bof) then
		username=Ers("username")
		dateandtime=Ers("dateandtime")
		rootid=Ers("rootid")
		topic=left(Ers("body"),20)
		Announceid=ers("Announceid")
		postuserid=ers("postuserid")
		LastPost=username & "$" & Announceid & "$" & dateandtime & "$" & replace(topic,"$","") & "$$" & postuserid & "$" & rootid & "$" & ers("BoardID")
		LastPost=Dvbbs.Checkstr(LastPost)
		Dvbbs.Execute("update [DV_topic] set LastPost='"&replace(LastPost,"'","")&"' where topicid="&rs(0))
		i=i+1
		'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""更新完"&server.htmlencode(ers(2)&"")&"的数据，正在更新下一个帖子数据，" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & server.htmlencode(eRs(2)&"") & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
		'End If
	end if
rs.movenext
loop
set ers=nothing
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
%>
<form action="admin_update.asp?action=updat" method=post>
<tr> 
<th align=left colspan=2 height=23>继续修复帖子(修复指定范围内帖子的最后回复数据)</th>
</tr>
<tr>
<td width="20%" class="forumrow">开始的ID号</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=5>&nbsp;帖子主题ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束的ID号</td>
<td width="80%" class="forumrow"><input type=text name="EndID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=5>&nbsp;将更新开始到结束ID之间的帖子数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="修 复"></td>
</tr>
</form>
<%
end sub

'分论坛今日帖子
function todays(boardid)
If IsSqlDataBase=1 Then
set tmprs=Dvbbs.Execute("Select count(announceid) from "&Dvbbs.NowUseBBS&" Where boardid="&boardid&" and datediff(day,dateandtime,"&SqlNowString&")=0")
else
set tmprs=Dvbbs.Execute("Select count(announceid) from "&Dvbbs.NowUseBBS&" Where boardid="&boardid&" and datediff('d',dateandtime,"&SqlNowString&")=0")
end if
todays=tmprs(0)
set tmprs=nothing
if isnull(todays) then todays=0
end function

'全部论坛今日帖子
function alltodays()
If IsSqlDataBase=1 Then
	set tmprs=Dvbbs.Execute("Select count(announceid) from "&Dvbbs.NowUseBBS&" Where not boardid in (444,777) and  datediff(day,dateandtime,"&SqlNowString&")=0")
Else
	set tmprs=Dvbbs.Execute("Select count(announceid) from "&Dvbbs.NowUseBBS&" Where not boardid in (444,777) and datediff('d',dateandtime,"&SqlNowString&")=0")
End If
alltodays=tmprs(0)
set tmprs=nothing
if isnull(alltodays) then alltodays=0
end function

'所有注册用户数量
function allusers() 
	set tmprs=Dvbbs.Execute("Select count(userid) from [Dv_user]") 
	allusers=tmprs(0) 
	Set tmprs=nothing 
	If IsNull(allusers) Then allusers=0 
End function
'最新注册用户
Function newuser()
	sql="Select top 1 username from [Dv_user] order by userid desc"
	Set tmprs=Dvbbs.Execute(sql)
	If tmprs.eof and tmprs.bof Then
		newuser="没有会员"
	Else
   		newuser=tmprs("username")
	End If
	Set tmprs=Nothing 
End function 

'所有论坛帖子
function AnnounceNum()
dim AnnNum
AnnNum=0
AnnounceNum=0
For i=0 to ubound(AllPostTable)
	set tmprs=Dvbbs.Execute("Select Count(announceID) from "&AllPostTable(i)&" where not boardid in (444,777)") 
	AnnNum=tmprs(0)
	set tmprs=nothing 
	if isnull(AnnNum) then AnnNum=0
	AnnounceNum=AnnounceNum + AnnNum
next
set tmprs=nothing
end function
'分论坛帖子
function BoardAnnounceNum(boardid)
dim BoardAnnNum
BoardAnnNum=0
BoardAnnounceNum=0
For i=0 to ubound(AllPostTable)
	set tmprs=Dvbbs.Execute("Select Count(announceID) from "&AllPostTable(i)&" where boardid="&boardid) 
	BoardAnnNum=tmprs(0) 
	set tmprs=nothing 
	if isnull(BoardAnnNum) then BoardAnnNum=0
	BoardAnnounceNum=BoardAnnounceNum + BoardAnnNum
next
set tmprs=nothing
end function

'所有论坛主题
function TopicNum() 
set tmprs=Dvbbs.Execute("Select Count(topicid) from DV_topic where not boardid in (444,777)") 
TopicNum=tmprs(0) 
set tmprs=nothing 
if isnull(TopicNum) then TopicNum=0 
end function

'分论坛主题
function BoardTopicNum(boardid) 
set tmprs=Dvbbs.Execute("Select Count(topicid) from [Dv_topic] where boardid="&boardid)
BoardTopicNum=tmprs(0) 
set tmprs=nothing 
if isnull(BoardTopicNum) then BoardTopicNum=0 
end function

'论坛总固顶主题数
function Forum_AllTopNum()
	Set tmprs=Dvbbs.Execute("Select TopicID From Dv_Topic Where IsTop=3 And (Not BoardID In (444,777)) ")
	If tmprs.eof and tmprs.bof Then
		Forum_AllTopNum=""
	Else
		Do While Not tmprs.Eof
			If Forum_AllTopNum="" Then
				Forum_AllTopNum = tmprs(0)
			Else
				Forum_AllTopNum = Forum_AllTopNum & "," & tmprs(0)
			End If
		tmprs.MoveNext
		Loop
	End If
	Set tmprs=Nothing
end function

'更新用户发贴数
sub updateTopic()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if
Dim C1
C1=clng(request.form("endid"))-clng(request.form("beginid"))
%>
</table>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<tr><td colspan=2 class=forumrow>
下面开始更新论坛用户资料，预计本次共有<%=C1%>个用户需要更新
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=ffffff height=9><img src="skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>

<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<%
Response.Flush
dim userTopic,UserPost
sql="select userid,username from [Dv_user] where userid>="&request.form("beginid")&" and userid<="&request.form("endid")
set rs=Dvbbs.Execute(sql)
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=forumrow>已经到记录的最尾端，请结束更新！</td></tr>"
	exit sub
end if
do while not rs.eof
	UserTopic=UserallTopicnum(rs(0))
	userPost=Userallnum(rs(0))
	Dvbbs.Execute("update [Dv_user] set UserPost="&userPost&",UserTopic="&UserTopic&" where userid="&rs(0))
	i=i+1
	'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""更新完"&rs(1)&"的数据，正在更新下一个用户数据，" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Rs(1) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
	'End If
rs.movenext
loop
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>继续更新用户数据</th>
</tr>
<tr>
<td width="20%" class="forumrow">重新计算用户发贴</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>发贴重新计算所有用户发表帖子数量。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="重新计算用户发贴"></td>
</tr>
</form>
<%
end sub

'更新用户金钱/经验/魅力
sub updatemoney()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if
dim userTopic,userReply,userWealth
dim userEP,userCP

Dim C1
C1=clng(request.form("endid"))-clng(request.form("beginid"))
%>
</table>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<tr><td colspan=2 class=forumrow>
下面开始更新论坛用户资料，预计本次共有<%=C1%>个用户需要更新
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=ffffff height=9><img src="skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>

<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<%
Response.Flush
sql="select userlogins,userid,userpost,usertopic,username from [Dv_user] where userid>="&request.form("beginid")&" and userid<="&request.form("endid")
set rs=Dvbbs.Execute(sql)
'shinzeal加入自动提示完成
if rs.eof and rs.bof then
	body="<tr><td colspan=2 class=forumrow>已经到记录的最尾端，请结束更新！</td></tr>"
	exit sub
end if
do while not rs.eof
	'userTopic=UserTopicNum(rs(1))
	'userreply=UserReplyNum(rs(1))
	userwealth=rs(0)*Dvbbs.Forum_user(4) + rs("usertopic")*Dvbbs.Forum_user(1) + (rs("userpost")-rs("usertopic"))*Dvbbs.Forum_user(2)
	userEP=rs(0)*Dvbbs.Forum_user(9) + rs("usertopic")*Dvbbs.Forum_user(6) + (rs("userpost")-rs("usertopic"))*Dvbbs.Forum_user(7)
	userCP=rs(0)*Dvbbs.Forum_user(14) + rs("usertopic")*Dvbbs.Forum_user(11) + (rs("userpost")-rs("usertopic"))*Dvbbs.Forum_user(12)
	if isnull(UserWealth) or not isnumeric(userwealth) then userwealth=0
	if isnull(Userep) or not isnumeric(userep) then userep=0
	if isnull(Usercp) or not isnumeric(usercp) then usercp=0
	Dvbbs.Execute("update [Dv_user] set userWealth="&userWealth&",userep="&userep&",usercp="&usercp&" where userid="&rs(1))
	i=i+1
	'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""更新完"&rs(4)&"的数据，正在更新下一个用户数据，" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Rs(4) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
	'End If
rs.movenext
loop
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>继续更新用户数据</th>
</tr>
<tr>
<td width="20%" class="forumrow" valign=top>更新用户金钱/经验/魅力</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户的发贴数量和论坛的相关设置重新计算用户的金钱/经验/魅力，本操作也将重新计算贵宾、版主、总版主的数据<BR>注意：不推荐用户进行本操作，本操作在数据很多的时候请尽量不要使用，并且本操作对各个版面删除帖子等所扣相应分值不做运算，只是按照发贴和总的论坛分值设置进行运算，请大家慎重操作，<font color=red>而且本项操作将重置用户因为奖励、惩罚等原因管理员对用户分值的修改。</font></td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户金钱/经验/魅力"></td>
</tr>
</form>
<%
end sub

'更新用户等级
sub updategrade()
if not isnumeric(request.form("beginid")) then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	exit sub
end if
if not isnumeric(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	exit sub
end if
if clng(request.form("beginid"))>clng(request.form("endid")) then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	exit sub
end if

Dim oldMinArticle,Rss
oldMinArticle=0
set Rss=Dvbbs.Execute("select userid from [Dv_user] where userid>="&request.form("beginid"))
if Rss.eof and rss.bof then
	body="<tr><td colspan=2 class=forumrow>已经到记录的最尾端，请结束更新！</td></tr>"
	exit sub
end if
Rss.close

SQL = "Select usergroupid From Dv_UserGroups where ParentGID=4"
SET Rss = Conn.Execute(SQL)
	SQL = Rss.GetString(,, "", ",", "")
Rss.close
Set Rss = Nothing
SQL = SQL&"4"

set rs=Dvbbs.Execute("select * from Dv_UserGroups where ParentGID=4 order by MinArticle desc")
do while not rs.eof
	Dvbbs.Execute("update [Dv_user] set userclass='"&rs("usertitle")&"',titlepic='"&rs("grouppic")&"' where usergroupid in ("&SQL&")  and (userid>="&request.form("beginid")&" and userid<="&request.form("endid")&") and (userpost<"&oldMinArticle&" and userpost>="&rs("MinArticle")&" )")
	oldMinArticle=rs("MinArticle")
rs.movenext
loop
rs.close
set rs=nothing
%>
<FORM METHOD=POST ACTION="?action=updateuserinfo">
<tr> 
<th align=left colspan=2 height=23>继续更新用户数据</th>
</tr>
<tr>
<td width="20%" class="forumrow" valign=top>更新用户等级</td>
<td width="80%" class="forumrow">执行本操作将按照<font color=red>当前论坛数据库</font>用户发贴数量和论坛的等级设置重新计算用户等级，本操作不影响等级为贵宾、版主、总版主的数据。</td>
</tr>
<tr>
<td width="20%" class="forumrow">开始用户ID</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=10>&nbsp;用户ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束用户ID</td>
<td width="80%" class="forumrow"><input type=text name="endID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=10>&nbsp;将更新开始到结束ID之间的用户数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="更新用户等级"></td>
</tr>
</form>
<%
end sub

'用户所有主题数
function UserTopicNum(userid)
dim topicnum
topicnum=0
usertopicnum=0
set tmprs=Dvbbs.Execute("select count(*) from dv_topic where not boardid in (444,777) and PostUserID="&userid)
TopicNum=tmprs(0)
if isnull(TopicNum) then TopicNum=0
UserTopicNum=UserTopicNum + TopicNum
set tmprs=nothing
end function
'用户所有回复数
Function UserReplyNum(userid)
dim replynum
replynum=0
userreplynum=0
For i=0 to ubound(AllPostTable)
	set tmprs=Dvbbs.Execute("select count(announceid) from "&AllPostTable(i)&" where not boardid in (444,777) and ParentID>0 and PostUserID="&userid)
	replyNum=tmprs(0)
	if isnull(replyNum) then replyNum=0
	UserReplyNum=UserReplyNum + replynum
next
set tmprs=nothing
end function
'用户所有帖子
function Userallnum(userid)
dim allnum
allnum=0
userallnum=0
For i=0 to ubound(AllPostTable)
	set tmprs=Dvbbs.Execute("select count(announceid) from "&AllPostTable(i)&" where not boardid in (444,777) and PostUserID="&userid)
	allnum=tmprs(0)
	if isnull(allnum) then allnum=0
	userallnum=userallnum+allnum
Next
Set tmprs=nothing
End function

function UserallTopicnum(userid)
dim allnum
allnum=0
UserallTopicnum=0
For i=0 to ubound(AllPostTable)
	set tmprs=Dvbbs.Execute("select count(*) from Dv_Topic where not boardid in (444,777) and PostUserID="&userid)
	allnum=tmprs(0)
	if isnull(allnum) then allnum=0
	UserallTopicnum=UserallTopicnum+allnum
Next
Set tmprs=nothing
End function

Sub fixbbs()
Dim killhtml,updatedate
updatedate=Request("updatedate")
killhtml=Request("killhtml")
If Not IsNumeric(request.form("beginid")) Then
	body="<tr><td colspan=2 class=forumrow>错误的开始参数！</td></tr>"
	Exit Sub
End If
If Not IsNumeric(request.form("endid")) Then
	body="<tr><td colspan=2 class=forumrow>错误的结束参数！</td></tr>"
	Exit Sub
End If
If CLng(request.form("beginid"))>clng(request.form("endid")) Then
	body="<tr><td colspan=2 class=forumrow>开始ID应该比结束ID小！</td></tr>"
	Exit Sub
End If
Dim C1
C1=clng(request.form("endid"))-clng(request.form("beginid"))
%>
</table>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<tr><td colspan=2 class=forumrow>
下面开始更新论坛帖子资料，预计本次共有<%=C1%>个帖子需要更新
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=ffffff height=9><img src="skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table>
<span id=txt3 name=txt3 style="font-size:9pt;color:red;"></span>
<span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>

<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<%
Response.Flush
Dim TotalUseTable,Ers,SQL1
Dim vBody,isagree,re
Dim Maxid
Set Rs=Dvbbs.Execute("select Max(topicid) from dv_topic")
Maxid=Rs(0)
Set Rs=Nothing
If Maxid< CLng(request.form("beginid")) Then
	body="<tr><td colspan=2 class=forumrow>已经到记录的最尾端，请结束更新！</td></tr>"
	set rs=nothing
	Exit Sub
End If
sql="select topicid,PostTable from dv_topic where topicid>="&CLng(request.form("beginid"))&" and topicid<="&CLng(request.form("endid"))
Set Rs=Dvbbs.Execute(sql)
Set  ERs=server.createobject("adodb.recordset")
Do While Not rs.eof
	SQl1 ="select Body,isagree,Ubblist,topic,DateAndTime,Announceid from "&Rs(1)&" where Rootid="&CLng(Rs(0))&""
	Ers.open SQL1,Conn,1,3
	If Not(eRs.eof OR ers.BOF) Then
		Do While Not ers.eof
			vbody=eRs(0)
			isagree=eRs(1)
			If IsNull(isagree) Then isagree=""
			If isagree <>""  Then
				isagree=Replace(isagree,"[isubb]","")
			Else
				isagree=ers(1)
			End If
			If killhtml="1" Then
				eRs(0)=vbody
			End If
			ers(1)=	isagree&""
			If IsDate(updatedate) And IsDate(ers(4)) Then
				If updatedate > ers(4) Then
					ers(2)=Ubblist(vbody)
				Else
					If killhtml="1" Then
						Set re=new RegExp
						re.IgnoreCase =true
						re.Global=True
						re.Pattern="<br>"
						vbody=re.Replace(vbody,"[br]")
						re.Pattern="<(.[^>]*)>"
						vbody=re.Replace(vbody,"")
						Set re=Nothing
					End If 
					ers(2)=UbblistOLD(vbody)
				End If
			Else
				ers(2)=UbblistOLD(vbody)
			End If
		ers.update
			Response.Write "<script>txt3.innerHTML=""更新完编号为"&eRs(5)&"的帖子数据，"";</script>"
			Response.Flush
		eRs.movenext 
		Loop
	End If
	ERs.close
	i=i+1
	'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 100) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""更新完编号为"&Rs(0)&"的数据，正在更新下一个帖子数据，" & FormatNumber(i/C1*25,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Rs(0) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
	'End If
Rs.movenext 
Loop
Set ers=nothing
set rs=nothing
Response.Write "<script>img2.width=400;txt3.innerHTML='';txt2.innerHTML=""100"";</script>"
%>
<form name=Fix action="admin_update.asp?action=fix" method=post>
<tr> 
<th align=left colspan=2 height=23>继续修复帖子(修复指定范围贴子UBB标签)</th>
</tr>
<tr>
<td width="20%" class="forumrow">开始的ID号</td>
<td width="80%" class="forumrow"><input type=text name="beginID" value="<%=request.form("endid")+1%>" size=5>&nbsp;帖子主题ID，可以填写您想从哪一个ID号开始进行修复</td>
</tr>
<tr>
<td width="20%" class="forumrow">结束的ID号</td>
<td width="80%" class="forumrow"><input type=text name="EndID" value="<%=request.form("endid")+(request.form("endid")-request.form("beginid"))+1%>" size=5>&nbsp;将更新开始到结束ID之间的帖子数据，之间的数值最好不要选择过大</td>
</tr>
<tr>
<td width="20%" class="forumrow">新老贴的标识日期</td>
<td width="80%" class="forumrow"><input type="text" name="updatedate" value="<%=updatedate%>">(格式：YYYY-M-D) 就是论坛升级到v7.0的日期，如果不填写，一律按老贴处理</td>
</tr>
<tr>
<td width="20%" class="forumrow">去掉贴子中的HTML标记</td>
<td width="80%" class="forumrow"><input type="radio" name="killhtml" value="1" 
 <%
  If killhtml="1" Then 
  %>
  checked 
  <%
  End If 
  %>
> 是
  <input type="radio" name="killhtml" value="0" 
  <%
  If killhtml="0" Then 
  %>
  checked 
  <%
  End If 
  %>
  > 否 &nbsp;<br>选是的话，贴子中的HTML标记将会自动被清除，有利于减少数据库的大小，但是会失去原来的HTML效果。</td>
</tr>
<tr>
<td width="20%" class="forumrow"></td>
<td width="80%" class="forumrow"><input type="submit" name="Submit" value="修 正"></td>
</tr>
</form>
<%
End Sub
%>
