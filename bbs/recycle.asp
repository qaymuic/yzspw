<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<!-- #include file="inc/dv_ubbcode.asp" -->
<%
Rem 修改记录：2004-5-3 行396 Dvbbs.YangZheng
Server.ScriptTimeOut=999999
dim bBoardEmpty
dim totalrec
dim n,RowCount
dim p
dim currentpage,page_count,Pcount
dim tablename
Dvbbs.stats="论坛回收站"
Dvbbs.nav()
If Not Dvbbs.master Then Response.redirect "showerr.asp?ErrCodes=<li>您没有权限浏览本页。&action=OtherErr"
Dvbbs.ShowErr()
Dvbbs.Head_var 2,0,"",""
Dim EmotPath
EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)
Dim dv_ubb,abgcolor
Set dv_ubb=new Dvbbs_UbbCode
If Request("reaction")="manage" Then
	dim topicid
	Dim trs,UpdateBoardID
	Dim fixid
	Manage_Main()
ElseIf Request("reaction")="view" Then
	View()
Else
	Main()
End If

Call Dvbbs.activeonline()
call Dvbbs.footer()

Sub Main()
	currentPage=request("page")
	If currentpage="" or not IsNumerIc(currentpage) Then
		currentpage=1
	Else
		currentpage=clng(currentpage)
	End If
	call AnnounceList1()
	call listPages3()
End Sub

Sub Manage_Main()
	topicid=replace(request("topicid"),"'","")
	topicid=replace(topicid,";","")
	topicid=replace(topicid,"--","")
	If request("action")<>"清空回收站" Then
		If topicid="" or IsNull(topicid) Then
			Response.redirect "showerr.asp?ErrCodes=<li>请选择相关帖子后进行操作。&action=OtherErr"
		End If
		fixid=replace(topicid,",","")
		fixid=Trim(replace(fixid," ",""))
		If Not IsNumeric(fixid) Then
			Response.redirect "showerr.asp?ErrCodes=<li>参数错误&action=OtherErr"
		End If
	End If 
	If request("tablename")="dv_topic" Then 
		tablename="dv_topic"
	ElseIf InStr(request("tablename"),"bbs")>0 Then
		tablename=Trim(request("tablename"))
		If Len(tablename)>8 Then
			Response.redirect "showerr.asp?ErrCodes=<li>错误的系统参数！&action=OtherErr"
		End If
	Else
		Response.redirect "showerr.asp?ErrCodes=<li>错误的系统参数！&action=OtherErr"
	End If

	If request("action")="删除" Then 
		call delete()
	ElseIf request("action")="还原" Then 
		call redel()
	ElseIf  request("action")="清空回收站" Then
		call Alldel()
	Else
		Response.redirect "showerr.asp?ErrCodes=<li>请指定所需参数。&action=OtherErr"
	End If
End Sub

Sub View()
	dim AnnounceID
	dim username
	If  request("id")="" Then
		Response.redirect "showerr.asp?ErrCodes=<li>请指定所需参数。&action=OtherErr"
	ElseIf Not IsNumeric(request("id")) Then
		Response.redirect "showerr.asp?ErrCodes=<li>请指定所需参数。&action=OtherErr"
	Else
		AnnounceID=request("id")
	End If
	tablename=Trim(request("tablename"))
	If Len(tablename)>8 Then
		Response.redirect "showerr.asp?ErrCodes=<li>请指定所需参数。&action=OtherErr"
	End If
	Set Rs=server.createobject("adodb.recordset")
	If InStr(tablename,"bbs")>0 Then
		sql="select Announceid,topic,body,username,dateandtime from "&replace(tablename,"'","")&" where AnnounceID="&AnnounceID
		tablename=request("tablename")
	Else
		sql="select topicid,title,title as body,postusername,dateandtime,posttable from dv_topic where topicID="&AnnounceID
		tablename="dv_topic"
	End  If
	If Not IsObject(Conn) Then ConnectionDatabase
	rs.open sql,conn,1,1
	If rs.eof and rs.bof Then
		Response.redirect "showerr.asp?ErrCodes=<li>没有找到相关信息&action=OtherErr"
	Else 
%>

<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
<TBODY> 
<TR align=middle> 
<Th height=24><%=Dvbbs.htmlencode(rs(1))%></Th>
</TR>
<TR> 
<TD height=24 class=tablebody1>
<p align=center><a href="dispuser.asp?name=<%=Dvbbs.htmlencode(rs(3))%>" target=_blank><%=Dvbbs.htmlencode(rs(3))%></a> 发布于 <%=rs(4)%></p>
    <blockquote>   
      <br>   
<%
'shinzeal允许在回收站浏览中察看主题帖子内容
if tablename="dv_topic" then
dim rrs
Set rrs=Dvbbs.Execute("select body from "&rs("posttable")&" where rootID="&rs("topicid")&"")
response.Write dv_ubb.Dv_UbbCode(rrs(0),1,2,1)
set rrs=nothing
else
response.Write dv_ubb.Dv_UbbCode(rs(2),1,2,1)
end if%>
    </blockquote>
</TD>
</TR>
<TR align=middle> 
<TD height=24 class=tablebody2><a href="recycle.asp?reaction=manage&TopicID=<%=rs(0)%>&action=删除&tablename=<%=tablename%>">『 直接删除 』</a></TD>
</TR>
</TBODY>
</TABLE>
    </td>
  </tr>
</table>
<%
End If
rs.close
Set rs=nothing
End Sub

Sub AnnounceList1()
	response.write "<table cellpadding=0 cellspacing=0 border=0 width="&Dvbbs.mainsetting(0)&" align=center><tr>"&_
	"<td align=center width=2 valign=middle> </td>"&_
	"<td align=left valign=middle> 本页面只有系统管理员可进行操作，请选择需要的操作，列出所有：<a href=?tablename=dv_topic>主题表数据</a>"
	For i=0 to ubound(allposttable)
	response.write " | <a href=?tablename="&allposttable(i)&">帖子"&allposttablename(i)&"</a>"
	Next
	response.write "<BR><BR>注意：还原主题表数据将连跟贴（帖子）表数据一起还原，删除主题表数据将连跟帖（帖子）表一起删除</td>"&_
	"<td align=right> </td></tr></table><BR>"
	
		If instr(lcase(request("tablename")),"bbs")>0 then
			sql="select AnnounceID,boardID,UserName,Topic,body,DateAndTime from "&replace(request("tablename"),"'","")&" where boardid=444 and not parentid=0 order by announceid desc"
			tablename=request("tablename")
		Else
	 		sql="select topicID,boardID,PostUserName,Title,title as body,DateAndTime from dv_topic where boardid=444 order by topicid desc"
			tablename="dv_topic"
	End If 
	set rs=server.createobject("adodb.recordset")
	If Not IsObject(Conn) Then ConnectionDatabase
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		'论坛无内容
		call showEmptyBoard1()
	else
		rs.PageSize = cint(Dvbbs.Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
  	    	totalrec=rs.recordcount
		call showPageList1() 
	end if
End sub
	
	
Rem 显示贴子列表	
Sub showPageList1()
	dim body
	dim vrs
	dim votenum,votenum_1
	dim pnum
	i=0
	response.write "<form name=recycle action=recycle.asp?reaction=manage method=post><input type=hidden value="&tablename&" name=tablename>"&_
				"<TABLE cellPadding=1 cellSpacing=1 align=center class=tableborder1>"&_
  				"<TBODY>"&_
				"<TR align=middle>"&_
				"<Th height=25 width=32>状态</Th>"&_ 
				"<Th width=*>主 题</Th>"&_
				"<Th width=80>作 者</Th>"&_
				"<Th width=195>最后更新 | 回复人</Th>"&_
				"</TR>"
	While (not rs.eof) and (not page_count = rs.PageSize)
	response.write "<TR align=middle>"&_
				"<TD class=tablebody2 width=32 height=27>"

	response.write "<input type=checkbox name=topicid value="&rs(0)&">"

	response.write "</TD>"&_
				"<TD align=left class=tablebody1 width=*>"

	body=replace(replace(Dvbbs.htmlencode(left(rs(4),20)),"<BR>",""),"</P><P>","")
	
	response.write "<a href=""?reaction=view&id="&rs(0)&"&tablename="&tablename&""">"
	if rs(3)="" or isnull(rs(3)) then
		response.write body
	else
		if len(rs(3))>26 then
			response.write ""&left(Dvbbs.htmlencode(rs(3)),26)&"..."
		else
			response.write Dvbbs.htmlencode(rs(3))
		end if
	end if
	response.write "</a>"

	response.write "</TD>"&_
	"<TD class=tablebody2 width=80><a href=dispuser.asp?name="& rs(2) &">"& rs(2) &"</a></TD>"

	response.write "<TD align=left class=tablebody1 width=195>&nbsp;"

	'on error resume next
	Response.Write "&nbsp;"&_
	FormatDateTime(rs(5),2)&"&nbsp;"&FormatDateTime(rs(5),4)&_
	"&nbsp;|&nbsp;------"

	response.write "</TD></TR>"

	 page_count = page_count + 1
         rs.movenext
        wend
End sub


Sub listPages3()
	
	dim endpage
	'on error resume next
	Pcount=rs.PageCount
	if Dvbbs.master and totalrec >0 then
	
		response.write "<TR><TD height=27 width=""100%"" class=tablebody2 colspan=4><input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">选中所有显示帖子&nbsp;<input type=submit name=action onclick=""{if(confirm('确定还原选定的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=还原>"

		If session("flag")<>"" then
		response.write "&nbsp;<input type=submit name=action onclick=""{if(confirm('确定删除选定的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=删除>&nbsp;<input type=submit name=action onclick=""{if(confirm('确定清除回收站所有的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=清空回收站>"
		End if
		response.write "</TD></TR></form>"
	end if
	Response.Write "</TBODY></TABLE>"
	Response.Write "<table border=0 cellpadding=0 cellspacing=3 width="""&Dvbbs.mainsetting(0)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
			"每页<b>"&Dvbbs.Forum_Setting(11)&"</b> 贴数<b>"&totalrec&"</b></td>"&_
			"<td valign=middle nowrap><div align=right><p>分页： <b>"

	if currentpage > 4 then
	response.write "<a href=""?tablename="&tablename&"&page=1"">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color="&dvbbs.mainsetting(1)&">["&i&"]</font>"
		else
        response.write " <a href=""?tablename="&tablename&"&page="&i&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?tablename="&tablename&"&page="&Pcount&""">["&Pcount&"]</a></b>"
	end if
	response.write "</p></div></td></tr></table>"
	rs.close
	set rs=Nothing
End sub 

Sub showEmptyBoard1()
	Response.Write "<TABLE class=tableborder1 cellPadding=4 cellSpacing=1 align=center>"&_
			"<TBODY>"&_
			"<TR align=middle>"&_
			"<Th height=25>状态</Th>"&_
			"<Th>主 题</Th>"&_
			"<Th>作 者</Th> "&_
			"<Th>回复/人气</Th> "&_
			"<Th>最新回复</Th></TR> "&_
			"<tr><td colSpan=5 width=100% class=tablebody1>论坛回收站暂无内容。</td></tr>"&_
			"</TBODY></TABLE>"
End Sub

'删除回收站内容
sub delete()
	If InStr(tablename,"bbs")>0 Then 
		Dvbbs.Execute("delete from "&tablename&" where boardid=444 and Announceid in ("&TopicID&")")
	ElseIf tablename="dv_topic" then
		For i=0 to UBound(AllPostTable)
			Dvbbs.Execute("delete from "&allposttable(i)&" where boardid=444 and rootid in ("&TopicID&")")
		Next 
		Dvbbs.Execute("delete from dv_topic where boardid=444 and topicid in ("&TopicID&")")
	End If
	Dvbbs.Dvbbs_Suc("<li>帖子操作成功。<li>您的操作信息已经记录在案.")
End sub

'还原回收站内容
Sub redel()
	Dim tempnum,todaynum,lasttime,myrs,upchild,F_Announceid
	upchild=false
	If InStr(tablename,"bbs")>0 then
		sql="update "&tablename&" set boardid=locktopic,locktopic=0 where Announceid in ("&TopicID&")"
		Dvbbs.Execute(sql)
		'如果该回复帖对应主题已删除，则同时还原该主题帖
		set rs=Dvbbs.Execute("select topicid,posttable,boardid from dv_topic where boardid=444 and topicid in (select distinct rootid from "&tablename&" where Announceid in ("&TopicID&"))")
		do while not rs.eof
			Dvbbs.Execute("update "&rs(1)&" set boardid=locktopic,locktopic=0 where parentid=0 and rootid="&rs(0))
			set trs=Dvbbs.Execute("select count(*) from "&rs(1)&" where (not boardid=444) and not parentid=0 and rootid="&rs(0))
			set myrs=Dvbbs.Execute("select top 1 dateandtime from "&rs(1)&" where (not boardid=444) and rootid="&rs(0)&" order by announceid desc")
			Dvbbs.Execute("update dv_topic set boardid=locktopic,child="&trs(0)&",locktopic=0,lastposttime='"&myrs(0)&"' where topicid="&rs(0))
			Dvbbs.Execute("update dv_board set TopicNum=TopicNum+1 where boardid="&rs(2))
			upchild=true
		rs.movenext
		loop
		set rs=Dvbbs.Execute("select PostUserID,BoardID,DateAndtime,ParentID,rootid,Announceid,IsUpload from "&tablename&" where Announceid in ("&TopicID&")")
		do while not rs.eof
			sql="update [dv_user] set userpost=userpost+1,userWealth=userWealth+"&Dvbbs.Forum_user(3)&",userEP=userEP+"&Dvbbs.Forum_user(8)&",userdel=userdel-1 where userid="&rs(0)
			Dvbbs.Execute(sql)
			if not upchild then
			'更新被删除跟贴的主题回复数和最后更新时间
			set trs=dvbbs.execute("select top 1 dateandtime from "&tablename&" where (not boardid=444) and rootid="&rs("rootid")&" order by announceid desc")
			dvbbs.execute("update dv_topic set child=child+1,lastposttime='"&trs(0)&"' where topicid="&rs("rootid"))
			end if
			set trs=Dvbbs.Execute("select ParentStr,LastPost from dv_board where boardid="&rs(1))
			lasttime=split(trs(1),"$")(2)
			if not isdate(lasttime) then lasttime=dateadd("d",-3,Now())
			if datediff("d",rs(2),lasttime)=0 then
				todaynum=1
			else
				todaynum=0
			end if
			call AllboardNumAdd(todayNum,1,0)
			UpdateBoardID=trs(0) & "," & rs(1)
			Dvbbs.Execute("update dv_board set todaynum=todaynum+"&todaynum&",PostNum=PostNum+1 where boardid in ("&UpdateBoardID&")")
			LastCount(rs(1))
			'还原所含附件数据
			If rs(6)=1 Then
				F_Announceid=rs(4)&"|"&rs(5)
				Dvbbs.Execute("update DV_Upfile set F_flag=0 where F_boardid="&rs(1)&" and F_announceID = '"& F_Announceid&"'")
			End If
		rs.movenext
		loop
		set rs=nothing
	ElseIf tablename="dv_topic" then
		dim TotalUseTable,LastPost_a
		i=0
		todaynum=0
		sql="update dv_topic set boardid=locktopic,locktopic=0 where topicid in ("&TopicID&")"
		Dvbbs.Execute(sql)
		set rs=Dvbbs.Execute("select topicid,posttable,boardid from dv_topic where topicid in ("&topicid&")")
		do while not rs.eof
			set trs=Dvbbs.Execute("select ParentStr,LastPost from dv_board where boardid="&rs(2))
			UpdateBoardID=trs(0) & "," & rs(2)
			LastPost_a=split(trs(1),"$")(2)
			Dvbbs.Execute("update "&rs(1)&" set boardid=locktopic,locktopic=0 where rootid="&rs(0))
			set trs=Dvbbs.Execute("select postuserid,dateandtime from "&rs(1)&" where rootid="&rs(0))
			do while not trs.eof
				i=i+1
				sql="update [dv_user] set userpost=userpost+1,userWealth=userWealth+"&Dvbbs.Forum_user(3)&",userEP=userEP+"&Dvbbs.Forum_user(8)&",userdel=userdel-1 where userid="&trs(0)
				Dvbbs.Execute(sql)
				if datediff("d",trs(1),now())=0 then
					todaynum=todaynum+1
				else
					todaynum=todaynum
				end if
			trs.movenext
			loop
			call AllboardNumAdd(todayNum,i,1)
			Dvbbs.Execute("update dv_board set todaynum=todaynum+"&todaynum&",postNum=postNum+"&i&",TopicNum=TopicNum+1 where boardid in ("&UpdateBoardID&")")
			LastCount(rs(2))
			i=0
			todaynum=0
		rs.movenext
		loop
		set rs=nothing
		set trs=nothing
	End If 
	Dvbbs.Dvbbs_Suc("<li>帖子操作成功。<li>您的操作信息已经记录在案.")
End Sub 

'全部删除回收站内容 2004-5-3 Dvbbs.YangZheng
Sub AllDel()
	Dim Bn
	Sql = "SELECT TopicId, PostTable From Dv_Topic Where BoardId = 444 ORDER BY TopicId"
	Set Rs = Dvbbs.Execute(Sql)
	If Not (Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		For i = 0 To Ubound(Sql,2)
			Dvbbs.Execute("DELETE FROM " & Sql(1,i) & " WHERE RootId = " & Sql(0,i))
			Dvbbs.Execute("DELETE From Dv_Topic WHERE TopicId = " & Sql(0,i))
		Next
	End If
	For i = 0 To Ubound(Allposttable)
		Sql = "SELECT AnnounceId From " & Allposttable(i) & " WHERE BoardId = 444 ORDER BY AnnounceId"
		Set Rs = Dvbbs.Execute(Sql)
		If Not (Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
			Rs.Close:Set Rs = Nothing
			For Bn = 0 To Ubound(Sql,2)
				Dvbbs.Execute("DELETE FROM " & Allposttable(i) & " WHERE AnnounceId = " & Sql(0,Bn))
			Next
		End If
	Next
	Dvbbs.Dvbbs_Suc("<li>帖子操作成功。<li>您的操作信息已经记录在案.")
End Sub

Function LastCount(boardid)
	Dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
	Dim LastPost,uploadpic_n,Lastpostuserid,Lastid
	set trs=Dvbbs.Execute("select top 1 T.title,b.Announceid,b.dateandtime,b.username,b.postuserid,b.rootid from "&DVbbs.NowUseBBS&" b inner join dv_Topic T on b.rootid=T.TopicID where b.boardid="&boardid&" order by b.announceid desc")
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
	set trs=nothing

	LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & BoardID
	Dim SplitUpBoardID,SplitLastPost
	SplitUpBoardID=split(UpdateBoardID,",")
	For i=0 to ubound(SplitUpBoardID)
		set trs=Dvbbs.Execute("select LastPost from dv_board where boardid="&SplitUpBoardID(i))
		if not (trs.eof and trs.bof) then
			SplitLastPost=split(trs(0),"$")
			if ubound(SplitLastPost)=7 and clng(LastRootID)<>clng(SplitLastPost(1)) then
				Dvbbs.Execute("update dv_board set LastPost='"&LastPost&"' where boardid="&SplitUpBoardID(i))
			end if
		end if
	Next
	Set Trs=Nothing
End Function


'当还原时,所有论坛发帖数增加
Function AllboardNumAdd(todayNum,postNum,topicNum)
	sql="update dv_setup set forum_TodayNum=forum_todayNum+"&todaynum&",forum_postNum=forum_postNum+"&postNum&",forum_TopicNum=forum_topicNum+"&TopicNum
	Dvbbs.Execute(sql)
End Function
%>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>