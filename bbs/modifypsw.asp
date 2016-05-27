<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.LoadTemplates("Usermanager")
Dvbbs.Stats=Dvbbs.MemberName&template.Strings(2)
Dvbbs.Nav()
Dvbbs.Head_var 0,0,template.Strings(0),"Usermanager.asp"
Dim psw,password,oldpassword,quesion,answer,Usercookies,ErrCodes

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

sub Userinfo()
If Dvbbs.chkpost=False Then
	Dvbbs.AddErrCode(16)
	Exit Sub
End If
Dim Rs,Sql,tempstr
tempstr=template.html(9)
sql="Select Userid,UserAnswer,UserQuesion from [Dv_User] where Userid="&Dvbbs.Userid
Set Rs=Dvbbs.Execute(Sql)
If Rs.eof And Rs.bof Then
	Dvbbs.AddErrCode(32)
	Exit Sub
Else
	tempstr=Replace(tempstr,"{$user_id}",Rs(0))
	tempstr=Replace(tempstr,"{$user_answer}",Rs(1) & "")
	tempstr=Replace(tempstr,"{$user_quesion}",Rs(2) & "")
	tempstr=Replace(tempstr,"{$color}",Dvbbs.mainsetting(1))
	Response.write tempstr
End If
Rs.Close:Set Rs=Nothing
End Sub

Sub Update()
Dim Rs,sql
sql="Select Userpassword from [Dv_User] where Userid="&Dvbbs.Userid
Set Rs=Dvbbs.Execute(Sql)
if Rs.eof and Rs.bof then
	Dvbbs.AddErrCode(32)
Else
	if Request.Form("oldpsw")="" then
	  	ErrCodes=ErrCodes+"<li>"+template.Strings(27)'Dvbbs.AddErrMsg "请输入您的旧密码,才能完成修改。"
	ElseIf md5(trim(Request.Form("oldpsw")),16)<>trim(RS("Userpassword")) then
	  	ErrCodes=ErrCodes+"<li>"+template.Strings(28)'Dvbbs.AddErrMsg "输入的旧密码错误，请重新输入。"
	Else
		oldpassword=Request.Form("oldpsw")
	End If
	If Request.Form("psw")<>"" then
		password=md5(Request.Form("psw"),16)
	Else
		password=RS("Userpassword")
	End If
	If Request.Form("quesion")="" then
	  	ErrCodes=ErrCodes+"<li>"+template.Strings(29)'Dvbbs.AddErrMsg "请输入密码提示问题。"
	Else
		quesion=Request.Form("quesion")
	End If
	If Request.Form("answer")="" then
	  	ErrCodes=ErrCodes+"<li>"+template.Strings(30)'Dvbbs.AddErrMsg "请输入密码提示问题答案。"
	ElseIf Request.Form("answer")=Request.Form("oldanswer") then
		answer=Request.Form("answer")
	Else
		answer=md5(Request.Form("answer"),16)
	End If
End If

If ErrCodes<>"" Then Exit sub
Dvbbs.Showerr()

set rs=server.createobject("adodb.recordset")
sql="Select * from [Dv_User] where Userid="&Dvbbs.Userid
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	Dvbbs.AddErrCode(32)
	Exit Sub
Else
	'If Not Dvbbs.FoundIsChallenge Then
	Rs("Userpassword")=password
	rs("UserQuesion")=quesion
	rs("UserAnswer")=answer
	rs.Update
End If
rs.close:set rs=nothing
End Sub 
%>