<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<%
dim sql
If Dvbbs.BoardID = 0 Then
	Response.Write "参数错误"
	Response.End 
End If
Dim announceid,replyid
dim username
dim rootid
Dim topic
dim postbuyuser
Dim TotalUseTable
Dim dv_ubb
Set dv_ubb=new Dvbbs_UbbCode
abgcolor="#FFFFFF"
Dvbbs.LoadTemplates("postjob")
If request("id")="" Then
	Dvbbs.AddErrCode(43)
ElseIf Not Isnumeric(request("id")) Then
	Dvbbs.AddErrCode(30)
Else
	AnnounceID=Clng(request("id"))
End If 
If Dvbbs.GroupSetting(2)="0"  Then Dvbbs.AddErrcode(31)
Dvbbs.ShowErr()
Dim EmotPath
EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)		'em心情路径
Dim abgcolor,bgcolor
abgcolor="tablebody1"
bgcolor="tablebody2"
Call announceinfo()
Dvbbs.ShowErr()
Dvbbs.ActiveOnline
Dvbbs.Footer()
Sub announceinfo()
	Dim rs
	Dim Tempwrite,Templist
    Set rs=Dvbbs.execute("select title,PostTable from Dv_topic where topicID="&AnnounceID)
	If not(rs.bof and rs.eof) then
		TotalUseTable=rs(1)
		topic=rs(0)
	Else
		Dvbbs.AddErrCode(48)
		Exit sub
	End if
	Rs.close:Set rs=Nothing
	Tempwrite=template.html(2)
	Tempwrite=Replace(Tempwrite,"{$tablewidth}",Dvbbs.Mainsetting(0))
	Tempwrite=Replace(Tempwrite,"{$forumname}",Dvbbs.Forum_info(0))
	Tempwrite=Replace(Tempwrite,"{$forumurl}",Dvbbs.Get_ScriptNameUrl)
	Tempwrite=Replace(Tempwrite,"{$boardtype}",Dvbbs.Boardtype)
	Tempwrite=Replace(Tempwrite,"{$boardid}",Dvbbs.boardid)
	Tempwrite=Replace(Tempwrite,"{$topic}",Dvbbs.HtmlEncode(Topic))
	Tempwrite=Replace(Tempwrite,"{$announceid}",announceid)
	
	Sql="Select b.UserName,b.Topic,b.dateandtime,b.body,u.UserGroupID,b.postbuyuser,b.Ubblist from "&TotalUseTable&" b inner join [Dv_user] u on b.PostUserID=u.userid where b.boardid="&dvbbs.boardid&" and b.rootid="&Announceid&" and b.locktopic<>2 and u.lockuser=0 order by b.announceid"
	Set rs=Dvbbs.execute(Sql)
	If rs.eof and rs.bof then
		Dvbbs.AddErrCode(48)
		Exit sub
	Else
		Dim i
		Sql=Rs.GetRows(-1)
		Rs.close:set Rs=nothing
		For i=0 to Ubound(sql,2)
			postbuyuser=Sql(5,i)
			Ubblists=SQL(6,i)
			username=Sql(0,i)
			Templist=Templist&template.html(3)
			Templist=Replace(Templist,"{$username}",username)
			Templist=Replace(Templist,"{$dateandtime}",Sql(2,i))
			Templist=Replace(Templist,"{$topic}",Dvbbs.HtmlEncode(Sql(1,i)))
			Templist=Replace(Templist,"{$body}",SimJsReplace(dv_ubb.Dv_UbbCode(SQL(3,i),SQL(4,i),1,1)))
		Next
		Tempwrite=Replace(Tempwrite,"{$bbslist}",Templist)
		Dvbbs.stats=Dvbbs.HtmlEncode(Sql(1,0))
		Dvbbs.head()
		Response.write Tempwrite
	End if	
End sub
Function SimJsReplace(str)
	If IsNull(str) Or str="" Then Exit Function
	str=Replace(str,"\","\\")
	str=Replace(str,"'","\'")
	SimJsReplace=str
End Function
%>