<!--#include file="Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<!--#include file="inc/email.asp"-->
<%
'杨铮2003-11-29修改
Dim announceid
dim username
dim rootid
dim topic
dim mailbody
dim useremail
dim TotalUseTable
dim PostBuyUser,replyid
Dim Sql,rs
abgcolor="#FFFFFF"
Dvbbs.LoadTemplates("postjob")
dvbbs.stats=template.Strings(5)
If Cint(dvbbs.GroupSetting(15))=0 then
	Dvbbs.AddErrCode(49)
End if

If request("id")="" then
	Dvbbs.AddErrCode(43)
Elseif not Isnumeric(request("id")) then
	Dvbbs.AddErrCode(30)
Else
	AnnounceID=Clng(request("id"))
End if

Dvbbs.nav()
Dvbbs.ShowErr()

Dvbbs.head_var 1,Dvbbs.Board_Data(4,0),"",""
Dim dv_ubb,abgcolor
Set dv_ubb=new Dvbbs_UbbCode
Dim EmotPath
EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)		'em心情路径
If request("action")="sendmail" Then
	If IsValidEmail(trim(Request.Form("mail")))=false Then
		Dvbbs.AddErrCode(50)
		Dvbbs.ShowErr()
	Else
		useremail=trim(Request.Form("mail"))
	End If
	call announceinfo()
	Dvbbs.ShowErr()
	If Dvbbs.Forum_Setting(2)=0 Then
		Dvbbs.AddErrCode(51)
	ElseIf Dvbbs.Forum_Setting(2)=1 Then
		Call jmail(useremail,topic,mailbody)
	ElseIf Dvbbs.Forum_Setting(2)=2 Then
		Call Cdonts(useremail,topic,mailbody)
	ElseIf Dvbbs.Forum_Setting(2)=3 Then
		call aspemail(useremail,topic,mailbody)
	End If
	If SendMail="False" Then
		Dvbbs.AddErrCode(51)
	End If
	Dvbbs.ShowErr()
	Dvbbs.Dvbbs_suc("<li>"&template.Strings(6))
Else
	call pag()
End If	
Dvbbs.ActiveOnline
Dvbbs.Footer()

Sub Announceinfo()
	Dim Tempwrite,Templist
    Set Rs = Dvbbs.execute("SELECT Title, PostTable, PostUserid FROM Dv_Topic WHERE TopicID = " & AnnounceID)
	If Not(Rs.Bof And Rs.Eof) Then
		Topic = Rs(0)
		TotalUseTable=rs(1)
		If Dvbbs.Userid <> Rs(2) And Dvbbs.GroupSetting(2) = "0" Then
			Dvbbs.AddErrCode(31)
			Exit Sub
		End If
	Else
		Dvbbs.AddErrCode(48)
		exit sub
	End if
	rs.close
	mailbody=mailbody &template.html(4)
	Tempwrite=template.html(2)
	Tempwrite=Replace(Tempwrite,"{$tablewidth}",Dvbbs.Mainsetting(0))
	Tempwrite=Replace(Tempwrite,"{$forumname}",Dvbbs.Forum_info(0))
	Tempwrite=Replace(Tempwrite,"{$forumurl}",Dvbbs.Get_ScriptNameUrl)
	Tempwrite=Replace(Tempwrite,"{$boardtype}",Dvbbs.Boardtype)
	Tempwrite=Replace(Tempwrite,"{$boardid}",Dvbbs.boardid)
	Tempwrite=Replace(Tempwrite,"{$topic}",Dvbbs.HtmlEncode(Topic))
	Tempwrite=Replace(Tempwrite,"{$announceid}",announceid)

	Sql="Select b.UserName,b.Topic,b.dateandtime,b.body,u.UserGroupID,b.postbuyuser,b.ubblist from "&TotalUseTable&" b inner join [Dv_user] u on b.PostUserID=u.userid where b.boardid="&dvbbs.boardid&" and b.rootid="&Announceid&"  and b.locktopic<>2 and u.lockuser=0 order by b.announceid"
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
	End if
	mailbody=mailbody&Tempwrite
	mailbody=mailbody &"<div align=center>"&dvbbs.Forum_Copyright&"&nbsp;&nbsp;"&dvbbs.Forum_Version&"</div>"
'	response.write mailbody
'	mailbody=""
end sub

Sub pag()
	Dim Tempwrite
	Tempwrite=template.html(5)
	Tempwrite=Replace(Tempwrite,"{$announceid}",announceid)
	Tempwrite=Replace(Tempwrite,"{$boardid}",Dvbbs.boardid)
	Response.write Tempwrite
End sub
Function SimJsReplace(str)
	If IsNull(str) Or str="" Then Exit Function
	str=Replace(str,"\","\\")
	str=Replace(str,"'","\'")
	SimJsReplace=str
End Function
%>