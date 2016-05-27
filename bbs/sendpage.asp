<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<!-- #include file="inc/chkinput.asp" -->
<!-- #include file="inc/email.asp" -->
<%
'Ñîï£2003-11-20ÐÞ¸Ä
dim announceid
dim username
dim rootid
dim topic
dim mailbody
dim email
dim content
dim postname
dim incepts
dim announce
Dim Sql,rs
Dvbbs.LoadTemplates("postjob")
dvbbs.stats=template.Strings(9)
Dvbbs.Nav()
Dvbbs.ShowErr()
If Cint(dvbbs.GroupSetting(15))=0 Then
	Dvbbs.AddErrCode(65)
End If

If request("id")="" Then
	Dvbbs.AddErrCode(43)
ElseIf Not Isnumeric(request("id")) Then
	Dvbbs.AddErrCode(30)
Else
	AnnounceID=Clng(request("id"))
End If 
Dvbbs.ShowErr()

Dvbbs.head_var 1,Dvbbs.Board_Data(4,0),"",""
	If request("action")="sendmail" Then
		If IsValidEmail(trim(Request.Form("mail")))=False Then
			Dvbbs.AddErrCode(50)
		Else
			email=trim(Request.Form("mail"))
		End If
		If request("postname")="" Then
			Dvbbs.AddErrCode(66)
		Else 
			postname=request("postname")
		End If
		If request("incept")="" Then
			Dvbbs.AddErrCode(67)
		Else
			incepts=request("incept")
		End If
		If request("content")="" Then
			Dvbbs.AddErrCode(68)
		Else
			content=Dvbbs.HtmlEnCode(request("content"))
		End If
		Dvbbs.ShowErr()
		Call announceinfo()
		Dvbbs.ShowErr()
		if Dvbbs.Forum_Setting(2)=0 Then
			Dvbbs.AddErrCode(51)
		ElseIf Dvbbs.Forum_Setting(2)=1 Then
			call jmail(email,topic,mailbody)
		ElseIf Dvbbs.Forum_Setting(2)=2 Then
			call Cdonts(email,topic,mailbody)
		ElseIf Dvbbs.Forum_Setting(2)=3 Then
			call aspemail(email,topic,mailbody)
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

Sub announceinfo()
    Set rs=Dvbbs.execute("select title from Dv_topic where topicID="&AnnounceID)
	If Not(rs.bof and rs.eof) then
		topic=Dvbbs.HtmlEnCode(rs(0))
		rs.close:set rs=nothing
	Else
		Dvbbs.AddErrCode(48)
		Exit  Sub 
	End If
	mailbody=template.html(4)&template.html(6)
	mailbody=Replace(mailbody,"{$incepts}",incepts)
	mailbody=Replace(mailbody,"{$postname}",postname)
	mailbody=Replace(mailbody,"{$bbsname}",Dvbbs.Forum_Info(0))
	mailbody=Replace(mailbody,"{$boardtype}",Dvbbs.Boardtype)
	mailbody=Replace(mailbody,"{$topic}",topic)
	mailbody=Replace(mailbody,"{$content}",content)
	mailbody=Replace(mailbody,"{$bbsurl}",Dvbbs.Get_ScriptNameUrl)
	mailbody=Replace(mailbody,"{$boardid}",Dvbbs.Boardid)
	mailbody=Replace(mailbody,"{$announceid}",announceid)
	mailbody=Replace(mailbody,"{$copyright}",Dvbbs.Forum_Copyright)
	mailbody=Replace(mailbody,"{$version}",Dvbbs.Forum_Version)
'	response.write mailbody
'	mailbody=""
End Sub 

Sub pag()
	Dim Tempwrite
	Tempwrite=template.html(7)
	Tempwrite=Replace(Tempwrite,"{$bbsname}",Dvbbs.Forum_info(0))
	Tempwrite=Replace(Tempwrite,"{$forumurl}",Dvbbs.Get_ScriptNameUrl)
	Tempwrite=Replace(Tempwrite,"{$announceid}",announceid)
	Tempwrite=Replace(Tempwrite,"{$boardid}",Dvbbs.boardid)
	Response.write Tempwrite
End Sub
%>
