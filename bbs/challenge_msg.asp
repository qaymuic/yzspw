<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Dim ErrCodes
Dvbbs.LoadTemplates("")
Dvbbs.Stats="站内手机留言"
Dvbbs.Head()
Response.Write "<div topmargin=0 leftmargin=0 onkeydown=""if(event.keyCode==13 && event.ctrlKey)messager.submit()"">"

If Dvbbs.UserID=0 then ErrCodes=ErrCodes+"<li>您还没有登录，请登录后进行操作。"

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(6)=1) Then ErrCodes=ErrCodes+"<li>本论坛没有开启论坛手机短信互发功能。"

If ErrCodes<>"" Then Showerr

Select case request("action")
	case "submobile"
		Dvbbs.Stats="提交资料"
		call sendmsg_2()
	case "submessage"
		Dvbbs.Stats="提交资料"
		call sendmsg_3()
	case "redir"
		Dvbbs.Stats="提交资料"
		call redir()
	case else
		Dvbbs.Stats="输入资料"
		call sendmsg_1()
End Select 
If ErrCodes<>"" Then Showerr
Response.Write "</div>"
Dvbbs.Footer()

Sub sendmsg_1()
	dim Rs
	set Rs=Dvbbs.execute("select username from [dv_user] where userid="&Dvbbs.userid&" and IsChallenge=1")
	If Rs.EOF And Rs.BOF Then
		ErrCodes=ErrCodes+"<li>您还不是本站的阳光会员，不能使用此功能，请<a href=challenge_up.asp>升级为阳光会员</a>。"
		Exit Sub
	Else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr><th align=center colspan=2><form action="challenge_msg.asp?action=submobile" method=post>站内手机留言</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>请输入对方用户名</B>：<BR>对方必须是本站高级用户方可接收手机留言</td>
	<td class=tablebody1 width="60%">
<input type=text size=30 name="username" value="<%=rs("username")%>">
	</td></tr>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="下一步"></td></form></tr>
</table>
<%
	End If
	Set Rs=nothing
end sub

Sub sendmsg_2()
	Dim susername
	Dim Rs
	If request("username")="" Then
		If request("touser")="" Then
			ErrCodes=ErrCodes+"<li>请输入要发送站内手机留言的用户名。"
			Exit Sub
		Else
			susername=dvbbs.CheckStr(trim(request("touser")))
		End if
	Else
		susername=dvbbs.CheckStr(trim(request("username")))
	End If
	Set Rs=dvbbs.execute("select UserMobile from [dv_user] where username='"&susername&"' and IsChallenge=1")
	If rs.eof and rs.bof then
		ErrCodes=ErrCodes+"<li>您要发送的对象不是本站的阳光会员。"
		dvbbs.execute("insert into dv_message (incept,sender,title,content,sendtime,flag,issend) values ('"&susername&"','"&dvbbs.membername&"','站内手机短信发送失败通知','由于您不是本站的阳光会员，用户"&dvbbs.membername&"给您发送站内手机短信失败，您可以升级成为本站的手机阳光会员，以便您不在线的时候网友能够直接和您联系。',"&SqlNowString&",0,1)")
		UPDATE_User_Msg(susername)
		Exit Sub
	Else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr><th align=center colspan=2><form action="challenge_msg.asp?action=submessage" method=post>站内手机留言</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>对方用户名</B>：<BR>对方必须是本站高级用户方可接收手机留言</td>
	<td class=tablebody1 width="60%">
<%=susername%>
<input type=hidden name="username" value="<%=susername%>">
	</td></tr>
	<tr><td class=tablebody1 align=right width="40%"><B>请输入手机留言内容</B>：最多只能输入114个字符，多出来的系统将自动截断并分条发送<BR></td>
	<td class=tablebody1 width="60%">
<textarea cols=70 rows=6 name="message"><%=request("message")%></textarea>
	</td></tr>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="提 交"></td></form></tr>
</table>
<%
	End If
	set rs=nothing
End Sub

Sub sendmsg_3()
	Dim susername,message
	Dim mymobile,tomobile
	Dim imymobile,itomobile
	Dim rs
	If request("username")="" then
		If request("touser")="" Then
			ErrCodes=ErrCodes+"<li>请输入要发送站内手机留言的用户名。"
			Exit Sub
		Else
			susername=dvbbs.CheckStr(trim(request("touser")))
		End If
	Else
		susername=dvbbs.CheckStr(trim(request("username")))
	End If
	
	If request("message")="" Then
		ErrCodes=ErrCodes+"<li>请输入要发送站内手机留言的信息。"
		Exit Sub
	Else
		message=dvbbs.CheckStr(trim(request("message")))
	End If

	Set Rs=Dvbbs.execute("select UserMobile from [dv_user] where userid="&dvbbs.userid&" and IsChallenge=1")
	If Rs.EOF And Rs.BOF Then
		ErrCodes=ErrCodes+"<li>您还不是本站的阳光会员，不能使用此功能，<a href=challenge_up.asp>请升级为阳光会员</a>。"
		Exit Sub
	End If
	mymobile=rs(0)
	imymobile=mid(mymobile,3,1)
	If imymobile="0" or imymobile="1" or imymobile="2" or imymobile="3" or imymobile="4" Then
		imymobile=0
	Else 
		imymobile=1
	End If

	set rs=dvbbs.execute("select UserMobile from [dv_user] where username='"&susername&"' and IsChallenge=1")
	if rs.eof and rs.bof Then
		ErrCodes=ErrCodes+"<li>您要发送的对象不是本站的阳光会员。<br><li>您发送的手机短信已经自动转为站内短信发送给对方。"
		dvbbs.execute("insert into dv_message (incept,sender,title,content,sendtime,flag,issend) values ('"&susername&"','"&dvbbs.membername&"','站内手机短信发送失败通知','由于您不是本站的阳光会员，用户"&dvbbs.membername&"给您发送站内手机短信失败，您可以升级成为本站的手机阳光会员，以便您不在线的时候网友能够直接和您联系。"&chr(10)&"以下是用户"&dvbbs.membername&"给您发送的短信："&chr(10)&""&message&"',"&SqlNowString&",0,1)")
		UPDATE_User_Msg(susername)
		Exit Sub
	End If
	tomobile=rs(0)
	itomobile=mid(tomobile,3,1)
	If itomobile="0" or itomobile="1" or itomobile="2" or itomobile="3" or itomobile="4" Then
		itomobile=0
	Else 
		itomobile=1
	End If

	If imymobile<>itomobile Then
		ErrCodes=ErrCodes+"<li>对不起，根据相关政策，移动、联通手机不能互发短信。"
		Exit Sub
	End If

	Set Rs=Dvbbs.execute("select top 1 * from Dv_ChallengeInfo")
	Dim MyForumID,MouseID
	MyForumID=rs("D_ForumID")
	MouseID=rs("D_username")
%>
正在提交数据，请稍后……
<form name="redir" action="http://bbs.ray5198.com/send_message.jsp" method="post">
<INPUT type=hidden name="from" value="<%=mymobile%>">
<INPUT type=hidden name="to" value="<%=tomobile%>">
<INPUT type=hidden name="message" value="<%=message%>">
<INPUT type=hidden name="forumId" value="<%=MyForumID%>">
<INPUT type=hidden name="mouseId" value="<%=mouseid%>">
<input type=hidden name="forumUrl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<INPUT type=hidden name="sender" value="<%=dvbbs.membername%>">
<input type=hidden value="challenge_msg.asp?action=redir" name="dirPage">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
Set rs=nothing
End Sub

Sub redir()
	Dim ErrorCode,ErrorMsg

	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))

	Select Case ErrorCode
		Case 100
		Case 101
			ErrCodes=ErrCodes+"<li>您在论坛发送站内短信失败。"
			Exit Sub
		Case Else
			ErrCodes=ErrCodes+"<li>非法的提交过程。"
		Exit Sub
	End Select
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>成功：您成功的发送了站内短信</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li><a href="challenge_msg.asp">返回发送页面</a></li></ul>
</td></tr>
</table>
<%
End Sub

'更新用户短信通知信息（新短信条数||新短讯ID||发信人名）
Sub UPDATE_User_Msg(username)
	Dim msginfo,i,UP_UserInfo,newmsg
	newmsg=newincept(username)
	If newmsg>0 Then
		msginfo=newincept(username) & "||" & inceptid(1,username) & "||" & inceptid(2,username)
	Else
		msginfo="0||0||null"
	End If
	Dvbbs.execute("UPDATE [Dv_User] Set UserMsg='"&Dvbbs.CheckStr(msginfo)&"' WHERE username='"&Dvbbs.CheckStr(username)&"'")
	If username=Dvbbs.MemberName Then
		UP_UserInfo=Session(Dvbbs.CacheName & "UserID")
		UP_UserInfo(30)=msginfo
		Session(Dvbbs.CacheName & "UserID")=UP_UserInfo
	Else
		Call Dvbbs.NeedUpdateList(username,1)
	End If
End Sub

'统计留言
Function newincept(iusername)
Dim Rs
Rs=Dvbbs.execute("SELECT Count(id) FROM Dv_Message WHERE flag=0 And issend=1 And DelR=0 And incept='"& iusername &"'")
    newincept=Rs(0)
	Set Rs=nothing
	If isnull(newincept) Then newincept=0
End Function

Function inceptid(stype,iusername)
	Dim rs
	Set Rs=Dvbbs.execute("SELECT top 1 id,sender FROM Dv_Message WHERE flag=0 And issend=1 And DelR=0 And incept='"& iusername &"'")
	If not rs.eof Then
		If stype=1 Then
			inceptid=Rs(0)
		Else
			inceptid=Rs(1)
		End If
	Else
		If stype=1 Then
			inceptid=0
		Else
			inceptid="null"
		End If
	End If
	Set Rs=nothing
End Function
'-------------------------------------------------------------------------------------------------------------
'显示错误信息
Sub Showerr()
Dim Show_Errmsg
	If ErrCodes<>"" Then 
		Show_Errmsg=Dvbbs.mainhtml(14)
		ErrCodes=Replace(ErrCodes,"{$color}",Dvbbs.mainSetting(1))
		Show_Errmsg=Replace(Show_Errmsg,"{$color}",Dvbbs.mainSetting(1))
		Show_Errmsg=Replace(Show_Errmsg,"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		Show_Errmsg=Replace(Show_Errmsg,"{$action}",Dvbbs.Stats)
		Show_Errmsg=Replace(Show_Errmsg,"{$ErrString}",ErrCodes)
	End If
	Response.write Show_Errmsg
End Sub
%>