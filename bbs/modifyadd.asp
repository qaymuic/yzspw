<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/chkinput.asp"-->
<%
Dvbbs.LoadTemplates("Usermanager")
Dvbbs.Stats=Dvbbs.MemberName&template.Strings(3)
Dvbbs.Nav()
Dvbbs.Head_var 0,0,template.Strings(0),"Usermanager.asp"
Dim ErrCodes
If Dvbbs.Userid=0 Then
	Dvbbs.AddErrCode(6)
End If
If Cint(Dvbbs.GroupSetting(16))=0 Then
	Dvbbs.AddErrCode(28)
End If
Dvbbs.Showerr()
Response.write template.html(0)
If request("action")="updat" Then
	Call update()
	If ErrCodes<>"" Then Response.redirect "showerr.asp?ErrCodes="&ErrCodes&"&action=OtherErr"
	Dvbbs.Showerr()
	Dvbbs.Dvbbs_Suc("<li>"+template.Strings(26))
Else
	Call Userinfo()
	Dvbbs.Showerr()
End If
Dvbbs.ActiveOnline()
Dvbbs.Footer()

Sub userinfo()
	Dim Rs,Sql,tempstr,userim
	tempstr=template.html(10)
	sql="Select Userid,UserEmail,UserIM from [Dv_User] where Userid="&Dvbbs.Userid
	Set Rs=Dvbbs.Execute(Sql)
	If Rs.eof And Rs.bof Then
		Dvbbs.AddErrCode(32)
		Exit Sub
	Else
		tempstr=Replace(tempstr,"{$user_id}",Rs(0))
		tempstr=Replace(tempstr,"{$user_email}",Rs(1)&"")
		If rs(2)="" or isnull(rs(2)) Then
			tempstr=Replace(tempstr,"{$user_homepage}","")
			tempstr=Replace(tempstr,"{$user_oicq}","")
			tempstr=Replace(tempstr,"{$user_icq}","")
			tempstr=Replace(tempstr,"{$user_Msn}","")
			tempstr=Replace(tempstr,"{$user_Yahoo}","")
			tempstr=Replace(tempstr,"{$user_Aim}","")
			tempstr=Replace(tempstr,"{$user_UC}","")
		Else
			userim=split(rs(2),"|||")
			tempstr=Replace(tempstr,"{$user_homepage}",userim(0))
			tempstr=Replace(tempstr,"{$user_oicq}",userim(1))
			tempstr=Replace(tempstr,"{$user_icq}",userim(2))
			tempstr=Replace(tempstr,"{$user_Msn}",userim(3))
			tempstr=Replace(tempstr,"{$user_Aim}",userim(4))
			tempstr=Replace(tempstr,"{$user_Yahoo}",userim(5))
			tempstr=Replace(tempstr,"{$user_UC}",userim(6))
		End If
		Response.write tempstr
	End If
	Rs.Close:Set Rs =Nothing
End sub

Sub update()
	Dim Rs,Sql
	Dim Email,NewUserIM
	If Dvbbs.chkpost=False Then
		Dvbbs.AddErrCode(16)
		Exit Sub
	End If
	Dim userpassword
	userpassword=Request.form("password")
	If userpassword="" Then
		Dvbbs.AddErrCode(11)
		Exit Sub
	Else
		userpassword=md5(userpassword,16)	
	End If
	'校验密码，
	SQL="Select userpassword from dv_user where userid="&Dvbbs.UserID&""
	
	Set Rs=Dvbbs.Execute(SQL)
	If Not Rs.eof Then
		If Rs(0)<> userpassword Then
			Response.redirect "showerr.asp?ErrCodes=您输入的密码错误&action=OtherErr"
		End If
	Else
		Response.redirect "showerr.asp?ErrCodes=您输入的密码错误&action=OtherErr"
	End If
	Set rs=Nothing 
	If Not Dvbbs.FoundIsChallenge Then
		If IsValidEmail(Request.form("Email"))=false Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(31)		'Dvbbs.AddErrmsg "您的Email有错误。"
			Exit Sub
		Else
			If Not IsNull(Dvbbs.forum_setting(52)) And Dvbbs.forum_setting(52)<>"" And Dvbbs.forum_setting(52)<>"0" Then
				Dim SplitUserEmail,i
				SplitUserEmail=split(Dvbbs.forum_setting(52),"|")
				For i=0 to ubound(SplitUserEmail)
					If instr(Request.form("email"),SplitUserEmail(i))>0 Then
						ErrCodes=ErrCodes+"<li>"+template.Strings(32)		'Dvbbs.AddErrmsg "您填写的Email地址含有系统禁止字符。"
						Exit Sub
					End If
				Next
			End If
			Email=Dvbbs.checkstr(Request.form("Email"))
		End If
	Else
		Email=Dvbbs.checkstr(Request.form("Email"))
	End If
	If Request.form("Oicq")<>"" Then
		If Not isnumeric(Request.form("Oicq")) or len(Request.form("Oicq"))>12 Then
			Dvbbs.AddErrCode(18)
			Exit Sub
		End If
	End If
	'HomePage,UserOicq,UserIcq,UserMsn,UserAim,UserYahoo,UserUC
	NewUserIM=Request.form("homepage") &"|||"& Request.form("Oicq") &"|||"& Request.form("Icq") &"|||"& Request.form("Msn") &"|||"& Request.form("Yahoo") &"|||"& Request.form("UserAim") &"|||"& Request.form("UC")
	NewUserIM=Dvbbs.checkstr(NewUserIM)
	'update data
	sql="update [Dv_User] set UserEmail='"&Email&"',UserIM='"&NewUserIM&"' where Userid="&Dvbbs.Userid
	Set Rs=Dvbbs.Execute(Sql)
End Sub
%>