<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dvbbs.Loadtemplates("")
Dvbbs.stats="帖子审核"
Dvbbs.Nav
If Dvbbs.BoardID=0 Then
	Dvbbs.Head_var 2,0,"",""
Else
	Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
End If
Dim currentPage,Rs,SQl,i
Dim AdminLockTopic
Dim p,announceIDRange1,announceIDRange2,tableclass
Dim bBoardEmpty
bBoardEmpty=False
AdminLockTopic=False 
If (Dvbbs.master or Dvbbs.superboardmaster or Dvbbs.boardmaster) And Cint(Dvbbs.GroupSetting(36))=1 Then
	AdminLockTopic=True 
Else 
	AdminLockTopic=False 
End If
If Cint(Dvbbs.GroupSetting(36))=1 And Dvbbs.UserGroupID>3 Then
	AdminLockTopic=True 
End If 
If Dvbbs.FoundUserPer And Cint(Dvbbs.GroupSetting(36))=1 Then
	AdminLockTopic=true
ElseIf Dvbbs.FoundUserPer And Cint(Dvbbs.GroupSetting(36))=0 Then
	AdminLockTopic=False 
End If 
If Not AdminLockTopic Then Response.redirect "showerr.asp?ErrCodes=<li>您没有在本版面审核帖子的权限。&action=OtherErr"
currentPage=request("page")

If currentpage="" or not IsNumeric(currentpage) Then
	currentpage=1
Else
	currentpage=clng(currentpage)
End If

If request("action")="freetopic" Then
	freetopic()
ElseIf request("action")="dispaudit" Then
	View()
Else
	main()
End If

Dvbbs.activeonline()
Dvbbs.footer()

Sub main()
	Dim totalrec,ii,page_count
	Dim n,pi
	Dim rs1,sql1
%>
<BR>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>
<form action="?action=freetopic" method=post name=batch>
<input type=hidden value="<%=Dvbbs.boardid%>" name=boardid>
<TR align=middle>
<Th height=25 width=32 id=tabletitlelink>选项</th>
<Th width=*>主 题</Th>
<Th width=80>作 者</Th>
</TR>
<%
	Set Rs=server.createobject("adodb.recordset")
	If dvbbs.boardid=0 Then
		sql="select AnnounceID,boardID,UserName,Topic,DateAndTime,RootID,layer,orders,Expression,body,PostUserID,locktopic,parentid from "& Dvbbs.NowUseBBS &" where BoardID=777 Order by AnnounceID Desc"
	Else
		sql="select AnnounceID,boardID,UserName,Topic,DateAndTime,RootID,layer,orders,Expression,body,PostUserID,locktopic,parentid from "& Dvbbs.NowUseBBS &" where BoardID=777 And locktopic="&dvbbs.boardid&" Order by AnnounceID Desc"
	End If
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1
	If rs.eof And rs.bof Then
		Response.Write "<tr><td colSpan=3 width=100% class=tablebody1 height=25>&nbsp;暂无审核内容</td></tr>"
	Else
		rs.PageSize = cint(Dvbbs.Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		Do While Not Rs.Eof and (not page_count = rs.PageSize)
			page_count=page_count+1
			If rs("layer")= 1 Then
				tableclass="tablebody1"
			Else
				tableclass="tablebody2"
			End If
			Response.Write "<TR align=middle><TD class=tablebody2 width=32 height=27 class="&tableclass&">"
			Response.Write "<input type=checkbox name=Announceid value="""&rs("Announceid")&""">"
			Response.Write "</TD><TD align=left class=tablebody1 width=* class="&tableclass&">"

			Response.Write "<img src=skins/default/topicface/"&rs("Expression")&"> "
			If Rs("ParentID")>0 Then GetTopic(Rs("RootID"))
			Response.Write "<a href='accesstopic.asp?action=dispaudit&boardID="& Dvbbs.boardID &"&ID="&cstr(rs("RootID"))&"&replyID="&Cstr(rs("announceID"))&"' target=_blank>"

			If Rs("topic")="" or isnull(rs("topic")) Then
				If Len(rs("body"))>50 Then
					Response.Write Dvbbs.htmlencode(replace(left(rs("body"),50),chr(10),""))
				Else
					Response.Write Dvbbs.htmlencode(replace(rs("body"),chr(10),""))
				End If 
			Else
				If len(rs("Topic"))>50 Then 
					Response.Write Dvbbs.htmlencode(left(rs("Topic"),50))
				Else 
					Response.Write Dvbbs.htmlencode(rs("Topic"))
				End If
			End If
			Response.Write "&nbsp;("&rs("dateandtime")&")</TD>"
			Response.Write "<TD class=tablebody2 width=80 class="&tableclass&"><a href=""dispuser.asp?id="& rs("postuserid") &""" target=_blank>"& Dvbbs.htmlencode(rs("username")) &"</a></TD>"

			Response.Write "</TR>"
		Rs.MoveNext   
		Loop
	End If

	Rs.Close
	Set Rs=Nothing
	If totalrec mod Dvbbs.Forum_Setting(11)=0 Then
		n= totalrec \ Dvbbs.Forum_Setting(11)
	Else
   		n= totalrec \ Dvbbs.Forum_Setting(11)+1
	End If
	If currentpage-1 mod 10=0 Then
		p=(currentpage-1) \ 10
	Else
		p=(currentpage-1) \ 10
	End If
	Dim pagelist,pagelistbit
%>
<TR align=middle>
<Td height=25 class=tablebody2 colspan=3>&nbsp;请选择要操作的内容：<input name="actiontype" value="1" type=radio checked>通过审核&nbsp;<input name="actiontype" value="2" type=radio>删除帖子&nbsp;<input name=submit value="执行" type=submit onclick="{if(confirm('您确定执行的操作吗?')){return true;}return false;}"></Td>
</TR>
</table>

<table border=0 cellpadding=0 cellspacing=3 width="<%=Dvbbs.mainsetting(0)%>" align="center">
</form>
<form method=post action="accesstopic.asp">
<tr>
<td valign=middle>页次：<b><%= currentPage %></b>/<b><%= n %></b>页 每页<b><%= Dvbbs.Forum_Setting(11) %></b> 主题数<b><%= totalrec %></b></td>
<td valign=middle><div align=right >分页：
<%
	If currentPage=1 Then
	Response.Write "<font face=webdings color="&Dvbbs.mainsetting(1)&">9</font>   "
	Else
	Response.Write "<a href='?boardid="&Dvbbs.boardid&"&page=1&action="&request("action")&"' title=首页><font face=webdings>9</font></a>   "
	End If
	If p*10>0 Then Response.Write "<a href='?boardid="&Dvbbs.boardid&"&page="&Cstr(p*10)&"&action="&request("action")&"' title=上十页><font face=webdings>7</font></a>   "
	Response.Write "<b>"
	for ii=p*10+1 to P*10+10
		   If ii=currentPage Then
	          Response.Write "<font color="&Dvbbs.mainsetting(1)&">"+Cstr(ii)+"</font> "
		   Else
		      Response.Write "<a href='?boardid="&Dvbbs.boardid&"&page="&Cstr(ii)&"&action="&request("action")&"'>"+Cstr(ii)+"</a>   "
		   End If
		If ii=n Then exit for
		'p=p+1
	next
	Response.Write "</b>"
	If ii<n Then Response.Write "<a href='?boardid="&Dvbbs.boardid&"&page="&Cstr(ii)&"&action="&request("action")&"' title=下十页><font face=webdings>8</font></a>   "
	If currentPage=n Then
	Response.Write "<font face=webdings color="&Dvbbs.mainsetting(1)&">:</font>   "
	Else
	Response.Write "<a href='?boardid="&Dvbbs.boardid&"&page="&Cstr(n)&"&action="&request("action")&"' title=尾页><font face=webdings>:</font></a>   "
	End If
%>
转到:<input type=text name=Page size=3 maxlength=10  value='<%= currentpage %>'><input type=submit value=Go name=submit>
</div></td></tr>
<input type=hidden name=BoardID value='<%= Dvbbs.BoardID %>'>
</form></table>
<%
End sub

Function GetTopic(TopicID)
	Dim Trs
	Set Trs=Dvbbs.Execute("Select Title,BoardID From Dv_Topic Where TopicID="&TopicID)
	If Not(TRs.Eof And TRs.Bof) Then
		Response.Write "[<a href=dispbbs.asp?boardid="&trs(1)&"&id="&TopicID&" target=_blank>主题帖："&Dvbbs.HtmlEncode(Left(Trs(0),16))&"</a>] "
	Else
		Response.Write "[未找到相关主题] "
	End If
	Set Trs=Nothing
End Function

Sub freetopic()

	If request.form("announceid")="" Then Response.redirect "showerr.asp?ErrCodes=<li>请指定相关帖子。&action=OtherErr"
	Dim id,trs,ars
	Dim FoundID,MyID
	Dim bbsnum,topicnum,todaynum
	Dim haveaudit
	bbsnum=0
	topicnum=0
	todaynum=0
	for i=1 to request.form("Announceid").count
		ID=replace(request.form("Announceid")(i),"'","")
		If Not IsNumeric(ID) Then
			ID = 0
		Else
			ID = Clng(ID)
		End If
		'删除
		If request("actiontype")=2 Then
			Set Rs=Dvbbs.Execute("select rootid from "&Dvbbs.NowUsebbs&" where parentid=0 And Announceid="&id)
			If not (rs.eof And rs.bof) Then
				ID=Rs(0)
				Set Rs=Nothing 
				Dvbbs.Execute("delete from dv_topic where topicid="&ID)
				Dvbbs.Execute("delete from "&Dvbbs.NowUsebbs&" where rootid="&ID)
				FoundID=ID
			Else
				Dvbbs.Execute("delete from "&Dvbbs.NowUsebbs&" where Announceid="&id)
				FoundID=0
			End If
		'通过审核
		ElseIf cint(request("actiontype"))=1 Then
			Set Rs=Dvbbs.Execute("select rootid,dateandtime,PostUserID from "&Dvbbs.NowUsebbs&" where parentid=0 And Announceid="&id)
			If not (rs.eof And rs.bof) Then
				'如果被审核的是主题帖
				bbsnum=bbsnum+1
				topicnum=topicnum+1
				If datediff("d",rs(1),Now())=0 Then todaynum=todaynum+1
				Dvbbs.Execute("update dv_topic set boardid=locktopic,locktopic=0 where topicid="&rs(0))
				Dvbbs.Execute("update "&Dvbbs.NowUsebbs&" set boardid=locktopic,locktopic=0 where Announceid="&id)
				Dvbbs.Execute("update [dv_user] set userpost=userpost+1,userWealth=userWealth+"&Dvbbs.Forum_user(2)&",UserEP=UserEP+"&Dvbbs.Forum_user(7)&",UserCP=UserCP+"&Dvbbs.Forum_user(12)&" where userid="&rs(2))
			Else
				set trs=Dvbbs.Execute("select rootid,dateandtime,PostUserID from "&Dvbbs.NowUsebbs&" where Announceid="&id)
				If not (trs.eof And trs.bof) Then
				'更新主题最后回复数据和回复数
				bbsnum=bbsnum+1
				topicnum=topicnum+1
				If datediff("d",trs(1),Now())=0 Then todaynum=todaynum+1
				Dvbbs.Execute("update "&Dvbbs.NowUsebbs&" set boardid=locktopic,locktopic=0 where Announceid="&id)
				Dvbbs.Execute("update [dv_user] set userpost=userpost+1,userWealth=userWealth+"&Dvbbs.Forum_user(2)&",UserEP=UserEP+"&Dvbbs.Forum_user(7)&",UserCP=UserCP+"&Dvbbs.Forum_user(12)&" where userid="&trs(2))
				IsEndReply(trs(0))
				End If
			End If
		End If
	next
	Set Rs=Nothing
	'更新论坛总数据和版面数据
	If CInt(request("actiontype"))=1 Then update Dvbbs.boardid,bbsnum,topicnum,todaynum
	Dvbbs.Dvbbs_Suc("<li>帖子操作成功.")
End Sub 

'是否最后回复
Function IsEndReply(TopicID)
	isEndReply=false
	Dim trs
	Dim LastPostInfo,iTotalUseTable
	Dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
	Dim LastPost,uploadpic_n,LastPostUserID,LastID,istop
	set trs=Dvbbs.Execute("select LastPost,PostTable,istop from dv_Topic where Topicid="&Topicid)
	If not (trs.eof And trs.bof) Then
		LastPostInfo=split(trs(0),"$")
		iTotalUseTable=trs(1)
		istop=trs(2)
	End If
	set trs=Dvbbs.Execute("select top 1 topic,body,Announceid,dateandtime,username,PostUserid,rootid from "&iTotalUseTable&" where (Not BoardID In (444,777)) And RootID="&TopicID&" order by Announceid desc")
	If not(trs.eof And trs.bof) Then
		body=trs(1)
		LastRootid=trs(2)
		LastPostTime=trs(3)
		LastPostUser=replace(trs(4),"$","")
		LastTopic=left(replace(body,"$",""),20)
		LastPostUserID=trs(5)
		LastID=trs(6)
	Else
		LastTopic="无"
		LastRootid=0
		LastPostTime=now()
		LastPostUser="无"
		LastPostUserID=0
		LastID=0
	End If
	LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & replace(left(replace(LastTopic,"'",""),20),"$","") & "$" & LastPostInfo(4) & "$" & LastPostUserID & "$" & LastID & "$" & Dvbbs.boardid
	If istop=0 Then
		Dvbbs.Execute("update dv_topic set LastPost='"&LastPost&"',child=child+1,LastPostTime='"&LastPostTime&"' where topicid="&TopicID)
	Else
		Dvbbs.Execute("update dv_topic set LastPost='"&LastPost&"',child=child+1 where topicid="&TopicID)
	End If
	set trs=Nothing
End Function


'更新论坛总数据和版面数据
Function update(boardid,bbsnum,topicnum,todaynum)
	Dim lastpost_1,trs
	Dim LastTopic,LastRootid,LastPostTime,LastPostUser
	Dim LastPost,uploadpic_n,Lastpostuserid,Lastid
	Dim UpdateBoardID
	'本论坛和上级论坛ID
	UpdateBoardID=Dvbbs.Board_Data(3,0) & "," & Dvbbs.BoardID
	'版面最后回复数据
	set trs=Dvbbs.Execute("select top 1 T.title,b.Announceid,b.dateandtime,b.username,b.postuserid,b.rootid from "&Dvbbs.NowUsebbs&" b inner join Dv_Topic T on b.rootid=T.TopicID where b.boardid="&Dvbbs.boardid&" order by b.announceid desc")
	If not(trs.eof And trs.bof) Then
		Lasttopic=replace(left(replace(trs(0),"'",""),15),"$","")
		LastRootid=trs(1)
		LastPostTime=trs(2)
		LastPostUser=trs(3)
		LastPostUserid=trs(4)
		Lastid=trs(5)
	Else
		LastTopic="无"
		LastRootid=0
		LastPostTime=now()
		LastPostUser="无"
		LastPostUserid=0
		Lastid=0
	End If
	set trs=Nothing
	LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & Dvbbs.boardid
	'总版面最后回复数据
	set trs=Dvbbs.Execute("select top 1 T.title,b.Announceid,b.dateandtime,b.username,b.postuserid,b.rootid from "&Dvbbs.NowUsebbs&" b inner join Dv_Topic T on b.rootid=T.TopicID order by b.announceid desc")
	If not(trs.eof And trs.bof) Then
		Lasttopic=replace(left(replace(trs(0),"'",""),15),"$","")
		LastRootid=trs(1)
		LastPostTime=trs(2)
		LastPostUser=trs(3)
		LastPostUserid=trs(4)
		Lastid=trs(5)
	Else
		LastTopic="无"
		LastRootid=0
		LastPostTime=now()
		LastPostUser="无"
		LastPostUserid=0
		Lastid=0
	End If
	LastPost_1=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & Dvbbs.boardid

	Dim SplitUpBoardID,SplitLastPost
	SplitUpBoardID=split(UpdateBoardID,",")
	For i=0 to ubound(SplitUpBoardID)
		set trs=Dvbbs.Execute("select LastPost from dv_board where boardid="&SplitUpBoardID(i))
		If not (trs.eof And trs.bof) Then
			SplitLastPost=split(trs(0),"$")
			If isnull(SplitLastPost(1)) Then SplitLastPost(1)=0
			If ubound(SplitLastPost)=7 And clng(LastRootID)<>clng(SplitLastPost(1)) Then
				Dvbbs.Execute("update dv_board set LastPost='"&LastPost&"' where boardid="&SplitUpBoardID(i))
			End If
		End If
	Next
	Dvbbs.Execute("update dv_board set PostNum=PostNum+"&bbsnum&",TopicNum=TopicNum+"&TopicNum&",TodayNum=TodayNum+"&todaynum&" where boardid in ("&UpdateBoardID&")")
	Dvbbs.Execute("update dv_setup set  Forum_PostNum=Forum_PostNum+"&bbsnum&",Forum_TopicNum=Forum_TopicNum+"&TopicNum&",Forum_TodayNum=Forum_TodayNum+"&todaynum&",Forum_LastPost='"&LastPost_1&"'")
	set trs=Nothing
End Function


Sub View()
	dim AnnounceID,replyid
	dim username
	If request("id")="" Then
		Response.redirect "showerr.asp?ErrCodes=<li>请指定所需参数。&action=OtherErr"
	ElseIf Not IsNumeric(request("id")) Then
		Response.redirect "showerr.asp?ErrCodes=<li>请指定所需参数。&action=OtherErr"
	Else
		AnnounceID=request("id")
	End If
	If request("replyid")="" Then
		Response.redirect "showerr.asp?ErrCodes=<li>请指定所需参数。&action=OtherErr"
	ElseIf Not IsNumeric(request("replyid")) Then
		Response.redirect "showerr.asp?ErrCodes=<li>请指定所需参数。&action=OtherErr"
	Else
		replyid=request("replyid")
	End If
	Set Rs=server.createobject("adodb.recordset")
	set rs=dvbbs.execute("select posttable from dv_topic where topicid="&announceid)
	If rs.eof and rs.bof Then
		Response.redirect "showerr.asp?ErrCodes=<li>没有找到相关信息&action=OtherErr"
	end if
	dim tablename
	tablename=rs(0)
	set rs=dvbbs.execute("select * from "&tablename&" where announceid="&replyid)
	if rs.eof and rs.bof then
		Response.redirect "showerr.asp?ErrCodes=<li>没有找到相关信息&action=OtherErr"
	end if
%>

<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
<TBODY> 
<TR align=middle> 
<Th height=24><%=Dvbbs.htmlencode(rs("topic"))%></Th>
</TR>
<TR> 
<TD height=24 class=tablebody1>
<p align=center><a href="dispuser.asp?name=<%=Dvbbs.htmlencode(rs("username"))%>" target=_blank><%=Dvbbs.htmlencode(rs("username"))%></a> 发布于 <%=rs("dateandtime")%></p>
    <blockquote>   
      <br>   
<%
response.Write server.htmlencode(rs("body"))
%>
    </blockquote>
</TD>
</TR>
<TR align=middle> 
<TD height=24 class=tablebody2> </TD>
</TR>
</TBODY>
</TABLE>
    </td>
  </tr>
</table>
<%
rs.close
Set rs=nothing
End Sub
%>