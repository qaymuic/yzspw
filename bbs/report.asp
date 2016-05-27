<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<%
'Ñîï£2003-11-18ÐÞ¸Ä
Dim announceid
Dim username
Dim rootid
Dim topic
Dim mailbody
Dim email
Dim content
Dim postname
Dim incepts
Dim announce
Dim sql,Yrs
Dvbbs.LoadTemplates("postjob")
Dvbbs.stats=template.Strings(0)
Dvbbs.Nav()
If Dvbbs.userid=0 Then
	Dvbbs.AddErrCode(6)
End If
If not Isnumeric(request("id")) then
	Dvbbs.AddErrCode(30)
Else
	AnnounceID=clng(request("id"))
End If
Dvbbs.ShowErr()
Dvbbs.head_var 1,Dvbbs.Board_Data(4,0),"",""
If request("action")="send" then
		call announceinfo()
Else
		Call pag()
End If
Dvbbs.Showerr()

Sub announceinfo()
	dim body
	dim writer
	dim incept
	dim topic,topic_1
	body=Dvbbs.checkStr(request.Form("content"))
	writer=Dvbbs.membername
	incept=Dvbbs.checkStr(Request.Form("boardmaster"))
	sql="select title from Dv_topic where TopicID="&AnnounceID
	set Yrs=Dvbbs.Execute(sql)
	If not(Yrs.bof and Yrs.eof) Then
		topic_1=Yrs(0)
		topic=template.Strings(0)
		body=body&template.Strings(2)
		body=Replace(body,"{$dvbbsurl}","http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"report.asp",""))
		body=Replace(body,"{$boardid}",Dvbbs.boardid)
		body=Replace(body,"{$announceid}",Announceid)
	Else
		Dvbbs.AddErrCode(48)
		Exit sub
	End If
	Yrs.close
	Sql="insert into Dv_message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept&"','"&Dvbbs.membername&"','"&topic&"','"&body&"',"&SqlNowString&",0,1)"
	Dvbbs.execute(sql)
	update_user_msg(incept)
	Dvbbs.ActiveOnline
	Dvbbs.Dvbbs_suc("<li>"&template.Strings(3))
	Set Yrs=Nothing
	Dvbbs.Footer()
End sub
Sub pag()
	Dim i
	dim MainTable,Boardmasterlist
	Dim Boardmasterl,Boardmastersp
	Boardmasterlist=Template.Html(1)
	Sql="select boardmaster from Dv_board where boardID="&cstr(Dvbbs.boardid)
	Set Yrs=Dvbbs.execute(Sql)
	If Yrs.eof and Yrs.bof then
		Dvbbs.AddErrCode(29)
		Exit sub
	ElseIf Yrs(0)="" or isnull(Yrs(0)) then
		Boardmasterl=Replace(Boardmasterlist,"{$boardmaster}",template.Strings(1))
	Else
		Boardmastersp=Split(Yrs(0),"|")
		For i=0 to Ubound(Boardmastersp)
			Boardmasterl=Boardmasterl&Replace(Boardmasterlist,"{$boardmaster}",Boardmastersp(i))
		Next
	End if
	MainTable=Template.Html(0)
	MainTable=Replace(MainTable,"{$boardid}",Dvbbs.boardid)
	MainTable=Replace(MainTable,"{$announceid}",AnnounceID)
	MainTable=Replace(MainTable,"{$boardmasterlist}",Boardmasterl)
	Response.write MainTable
	Dvbbs.ActiveOnline
	Dvbbs.Footer()
End Sub
Function update_user_msg(username)
	Dim msginfo
	if newincept(username)>0 then
		msginfo=newincept(username) & "||" & inceptid(1,username) & "||" & inceptid(2,username)
	Else
		msginfo="0||0||null"
	End if
	Dvbbs.execute("update [Dv_user] set UserMsg='"&dvbbs.CheckStr(msginfo)&"' where username='"&dvbbs.CheckStr(username)&"'")
End function
'Í³¼ÆÁôÑÔ
Function newincept(iusername)
	Set Yrs=Dvbbs.execute("Select Count(id) From Dv_Message Where flag=0 and issend=1 and delR=0 And incept='"& iusername &"'")
    newincept=Yrs(0)
	set Yrs=nothing
	if isnull(newincept) then newincept=0
End function
Function inceptid(stype,iusername)
	Set Yrs=Dvbbs.execute("Select top 1 id,sender From Dv_Message Where flag=0 and issend=1 and delR=0 And incept='"& iusername &"'")
	If stype=1 then
		inceptid=Yrs(0)
	Else
		inceptid=Yrs(1)
	End if
	set Yrs=nothing
End function
%>
