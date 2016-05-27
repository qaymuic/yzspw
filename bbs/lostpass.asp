<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/email.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim topic,mailbody,sendmsg,useremail
dim username,password,repassword,Rs,SQl
dim answer
Dvbbs.LoadTemplates("login")
Dvbbs.Stats=template.Strings(2)
Dvbbs.Nav()
Dvbbs.Head_var 0,"",template.Strings(0),""
If request("action")="step1" Then
	call step1()
ElseIf request("action")="step2" Then
	call step2()
ElseIf  request("action")="step3" Then
	call step3()
Else	
	Call main()
End If
Dvbbs.activeonline()
Dvbbs.footer()

Sub step1()
	If Dvbbs.chkpost=False Then 
   		showerr template.Strings(10)
		Exit Sub
	End If
	If request.Form("username")="" Then
		showerr template.Strings(6)
		Exit Sub
	Else
		username=replace(request("username"),"'","")
	End If
	If Dvbbs.forum_setting(81)="1"  Then
		If Not Dvbbs.CodeIsTrue() Then
			 Response.redirect "showerr.asp?ErrCodes=<li>验证码校验失败，请返回刷新页面后再输入验证码。&action=OtherErr"
		End If
	End If
	If Dvbbs.Forum_Setting(2)<>"0" Then
		Set Rs=Dvbbs.execute("Select UserQuesion,userAnswer,Username,Usergroupid from [Dv_user] where username='"&username&"'")
	Else
		Set Rs=Dvbbs.execute("Select UserQuesion,userAnswer,Username,Usergroupid from [Dv_user] where username='"&username&"' and UserGroupID>3")
	End If
	If rs.eof and rs.bof then
		showerr template.Strings(8)
		Exit Sub 
	ElseIf  rs(3) < 4 then
		showerr template.Strings(7)
		Exit Sub 
	Else 
		If  rs(0)="" or isnull(rs(0)) Then 
			showerr template.Strings(9)
			Exit Sub 
		Else
			
			template.html(6)=Replace(template.html(6),"{$Quesion}",Rs(0))
			template.html(6)=Replace(template.html(6),"{$username}",username)
			If Dvbbs.forum_setting(81)="0"  Then
				template.html(6)=Replace(template.html(6),"{$getcode}","")
			Else
				template.html(6)=Replace(template.html(6),"{$getcode}"," 验证码："&Dvbbs.GetCode())
			End If 
			Response.Write template.html(6)		
		End If
	End If
	Rs.Close  
	Set Rs=Nothing 
End Sub 

Sub step2()
	Dim answer,UserToday,UpUserToday
	If request.Form("username")="" Then
		showerr template.Strings(6)
		Exit Sub
	Else
		username=replace(request("username"),"'","")
	End If
	If Dvbbs.chkpost=False Then 
   		showerr template.Strings(10)
		Exit Sub
	End If
	If request.form("answer")="" then
		showerr template.Strings(11)
		Exit Sub
	Else
		answer=md5(request("answer"),16)
	End If
	If Dvbbs.forum_setting(81)="1"  Then
		If Not Dvbbs.CodeIsTrue() Then
			 Response.redirect "showerr.asp?ErrCodes=<li>验证码校验失败，请返回刷新页面后再输入验证码。&action=OtherErr"
		End If
	End If
	sql="select useranswer,userquesion,useranswer,UserToday,Lastlogin from [Dv_user] where username='"&username&"'"
	Set Rs=Dvbbs.execute(sql)
	If Rs.EOF And Rs.BOF Then
		showerr template.Strings(12)
		Exit Sub
	Else
		If datediff("h",Rs(4),now())<24 Then
			showerr "您取回密码次数已超出系统限制，请24小时后再使用取回密码功能！"
			Exit Sub
		End If
		If Rs(2)<>answer Then
			If Cint(Dvbbs.forum_setting(84))<>0 Then
				If Ubound(Split(Rs(3),"|"))<3 Then
					UpUserToday=Rs(3)&"|0"
				Else
					UpUserToday=Rs(3)
				End IF
				UserToday=Split(UpUserToday,"|")
				UserToday(3)=Cint(UserToday(3))+1
				UpUserToday=Clng(UserToday(0))&"|"&Clng(UserToday(1))&"|"&Clng(UserToday(2))&"|"&Clng(UserToday(3))
				Dvbbs.Execute("update [Dv_user] Set UserToday='"&UpUserToday&"' Where username='"&username&"'")
				If UserToday(3)>Cint(Dvbbs.forum_setting(84)) Then
					UpUserToday=Clng(UserToday(0))&"|"&Clng(UserToday(1))&"|"&Clng(UserToday(2))
					Dvbbs.Execute("update [Dv_user] Set Lastlogin="&SqlNowString&",UserToday='"&UpUserToday&"' Where username='"&username&"'")
					showerr "您取回密码次数已超出系统限制，请24小时后再使用取回密码功能！"
					Exit Sub
				End If
			End If
			showerr template.Strings(12)
			Exit Sub
		Else
			If Dvbbs.Forum_Setting(2)<>"0" Then
				template.html(7)=Replace(template.html(7),"{$readme}",template.Strings(3))
			Else
				template.html(7)=Replace(template.html(7),"{$readme}",template.Strings(4))
			End If
			template.html(7)=Replace(template.html(7),"{$Quesion}",Rs(1))
			template.html(7)=Replace(template.html(7),"{$answer}",request.form("answer"))
			template.html(7)=Replace(template.html(7),"{$username}",username)
			Response.Write template.html(7)
		End If
	End If

	Rs.Close 
	Set Rs=Nothing
	
End Sub

Sub step3()
	If Dvbbs.chkpost=False Then 
   		showerr template.Strings(10)
		Exit Sub
	End If
	If request.Form("username")="" Then 
		showerr template.Strings(6)
		Exit Sub 
	Else 
		username=replace(request("username"),"'","")
	End If
	If request.Form("answer")="" Then
		showerr template.Strings(11)
		Exit Sub
	Else
		answer=md5(request("answer"),16)
	End  If 
	if request.Form("password")="" or Len(request("password"))>10 or len(request("password"))<6 then
		showerr template.Strings(13)
		Exit  Sub 
	ElseIf request.Form("repassword")="" Then
		showerr template.Strings(14)
		Exit Sub 
	ElseIf request.Form("password")<>request("repassword") Then
		showerr template.Strings(15)
		Exit Sub
	Else 
		password=md5(request.Form("password"),16)
	End If 
	set rs=server.createobject("adodb.recordset")
	sql="select userpassword,useremail,userquesion,userclass,UserGroupID from [Dv_user] where username='"&username&"' and useranswer='"&Dvbbs.checkStr(answer)&"'"
	If Not IsObject(Conn) Then ConnectionDatabase
	rs.open sql,conn,1,3
	If rs.eof and rs.bof Then 
		showerr template.Strings(16)
		Exit Sub 
	Else
		If Dvbbs.Forum_Setting(2)>0 Then
			repassword=request.form("password")
			answer=request.form("answer")
			password=rs("userpassword")
			useremail=rs("useremail")
			call sendusermail()
			If SendMail="OK" Then
				sendmsg=template.Strings(17)
			ElseIf Rs("UserGroupID")<4 Then
				sendmsg=template.Strings(18)
			Else
				rs("userpassword")=md5(repassword,16)
				Rs.Update
				sendmsg=template.Strings(19)
			End If
		Else
			Rs("userpassword")=password
			Rs.Update 
		End If
		template.html(8)=Replace(template.html(8),"{$Quesion}",Rs(2))
		template.html(8)=Replace(template.html(8),"{$answer}",request.form("answer"))
		template.html(8)=Replace(template.html(8),"{$password}",request.form("password"))
		If Dvbbs.Forum_Setting(2)<>"0" Then
			template.html(8)=Replace(template.html(8),"{$readme}",sendmsg)
		Else
			template.html(8)=Replace(template.html(8),"{$readme}",template.Strings(5))
		End If 
		Response.Write template.html(8)	
	End if
	Rs.Close
	Set Rs=Nothing 
End Sub 

Sub  main()
	If Dvbbs.Forum_ChanSetting(0)="1" Then  Response.Write template.html(4)
	If Dvbbs.forum_setting(81)="0"  Then
		template.html(5)=Replace(template.html(5),"{$getcode}","")
	Else
		template.html(5)=Replace(template.html(5),"{$getcode}"," 验证码："&Dvbbs.GetCode())
	End If 
	Response.Write template.html(5)
End  Sub 
Sub showerr(errmsg)
	template.html(9)=Replace(template.html(9),"{$Errmsg}",errmsg)
	Response.Write template.html(9)
End Sub 
Sub sendusermail()
	on error resume Next
	Dim activepassurl
	activepassurl=Dvbbs.Get_ScriptNameUrl()&"activepass.asp?username="&Dvbbs.htmlencode(username)&"&pass="&password&"&repass="&repassword&"&answer="&answer
	template.html(12)=Replace(template.html(12),"{$Forumname}",Dvbbs.Forum_info(0))
	topic=template.html(12)
	template.html(10)=Replace(template.html(10),"{$username}",Dvbbs.htmlencode(username))
	template.html(10)=Replace(template.html(10),"{$Copyright}",Dvbbs.Forum_Copyright)
	template.html(10)=Replace(template.html(10),"{$activepassurl}","<a href="&activepassurl&">"&activepassurl&"</a>")
	template.html(10)=Replace(template.html(10),"{$Version}","Dvbbs"&Dvbbs.Forum_Version)
	mailbody=template.html(10)
	select case Dvbbs.Forum_Setting(2)
	case 0
	sendmsg=template.Strings(20)&"<a href="&activepassurl&"><B>"&template.Strings(21)&"</B></a>"
	case 1
	call jmail(useremail,topic,mailbody)
	case 2
	call Cdonts(useremail,topic,mailbody)
	case 3
	call aspemail(useremail,topic,mailbody)
	case else
	sendmsg=template.Strings(20)&"<a href="&activepassurl&"><B>"&template.Strings(21)&"</B></a>"
	end select
End Sub
%>