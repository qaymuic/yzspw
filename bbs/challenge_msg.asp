<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Dim ErrCodes
Dvbbs.LoadTemplates("")
Dvbbs.Stats="վ���ֻ�����"
Dvbbs.Head()
Response.Write "<div topmargin=0 leftmargin=0 onkeydown=""if(event.keyCode==13 && event.ctrlKey)messager.submit()"">"

If Dvbbs.UserID=0 then ErrCodes=ErrCodes+"<li>����û�е�¼�����¼����в�����"

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(6)=1) Then ErrCodes=ErrCodes+"<li>����̳û�п�����̳�ֻ����Ż������ܡ�"

If ErrCodes<>"" Then Showerr

Select case request("action")
	case "submobile"
		Dvbbs.Stats="�ύ����"
		call sendmsg_2()
	case "submessage"
		Dvbbs.Stats="�ύ����"
		call sendmsg_3()
	case "redir"
		Dvbbs.Stats="�ύ����"
		call redir()
	case else
		Dvbbs.Stats="��������"
		call sendmsg_1()
End Select 
If ErrCodes<>"" Then Showerr
Response.Write "</div>"
Dvbbs.Footer()

Sub sendmsg_1()
	dim Rs
	set Rs=Dvbbs.execute("select username from [dv_user] where userid="&Dvbbs.userid&" and IsChallenge=1")
	If Rs.EOF And Rs.BOF Then
		ErrCodes=ErrCodes+"<li>�������Ǳ�վ�������Ա������ʹ�ô˹��ܣ���<a href=challenge_up.asp>����Ϊ�����Ա</a>��"
		Exit Sub
	Else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr><th align=center colspan=2><form action="challenge_msg.asp?action=submobile" method=post>վ���ֻ�����</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>������Է��û���</B>��<BR>�Է������Ǳ�վ�߼��û����ɽ����ֻ�����</td>
	<td class=tablebody1 width="60%">
<input type=text size=30 name="username" value="<%=rs("username")%>">
	</td></tr>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="��һ��"></td></form></tr>
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
			ErrCodes=ErrCodes+"<li>������Ҫ����վ���ֻ����Ե��û�����"
			Exit Sub
		Else
			susername=dvbbs.CheckStr(trim(request("touser")))
		End if
	Else
		susername=dvbbs.CheckStr(trim(request("username")))
	End If
	Set Rs=dvbbs.execute("select UserMobile from [dv_user] where username='"&susername&"' and IsChallenge=1")
	If rs.eof and rs.bof then
		ErrCodes=ErrCodes+"<li>��Ҫ���͵Ķ����Ǳ�վ�������Ա��"
		dvbbs.execute("insert into dv_message (incept,sender,title,content,sendtime,flag,issend) values ('"&susername&"','"&dvbbs.membername&"','վ���ֻ����ŷ���ʧ��֪ͨ','���������Ǳ�վ�������Ա���û�"&dvbbs.membername&"��������վ���ֻ�����ʧ�ܣ�������������Ϊ��վ���ֻ������Ա���Ա��������ߵ�ʱ�������ܹ�ֱ�Ӻ�����ϵ��',"&SqlNowString&",0,1)")
		UPDATE_User_Msg(susername)
		Exit Sub
	Else
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr><th align=center colspan=2><form action="challenge_msg.asp?action=submessage" method=post>վ���ֻ�����</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>�Է��û���</B>��<BR>�Է������Ǳ�վ�߼��û����ɽ����ֻ�����</td>
	<td class=tablebody1 width="60%">
<%=susername%>
<input type=hidden name="username" value="<%=susername%>">
	</td></tr>
	<tr><td class=tablebody1 align=right width="40%"><B>�������ֻ���������</B>�����ֻ������114���ַ����������ϵͳ���Զ��ضϲ���������<BR></td>
	<td class=tablebody1 width="60%">
<textarea cols=70 rows=6 name="message"><%=request("message")%></textarea>
	</td></tr>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="�� ��"></td></form></tr>
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
			ErrCodes=ErrCodes+"<li>������Ҫ����վ���ֻ����Ե��û�����"
			Exit Sub
		Else
			susername=dvbbs.CheckStr(trim(request("touser")))
		End If
	Else
		susername=dvbbs.CheckStr(trim(request("username")))
	End If
	
	If request("message")="" Then
		ErrCodes=ErrCodes+"<li>������Ҫ����վ���ֻ����Ե���Ϣ��"
		Exit Sub
	Else
		message=dvbbs.CheckStr(trim(request("message")))
	End If

	Set Rs=Dvbbs.execute("select UserMobile from [dv_user] where userid="&dvbbs.userid&" and IsChallenge=1")
	If Rs.EOF And Rs.BOF Then
		ErrCodes=ErrCodes+"<li>�������Ǳ�վ�������Ա������ʹ�ô˹��ܣ�<a href=challenge_up.asp>������Ϊ�����Ա</a>��"
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
		ErrCodes=ErrCodes+"<li>��Ҫ���͵Ķ����Ǳ�վ�������Ա��<br><li>�����͵��ֻ������Ѿ��Զ�תΪվ�ڶ��ŷ��͸��Է���"
		dvbbs.execute("insert into dv_message (incept,sender,title,content,sendtime,flag,issend) values ('"&susername&"','"&dvbbs.membername&"','վ���ֻ����ŷ���ʧ��֪ͨ','���������Ǳ�վ�������Ա���û�"&dvbbs.membername&"��������վ���ֻ�����ʧ�ܣ�������������Ϊ��վ���ֻ������Ա���Ա��������ߵ�ʱ�������ܹ�ֱ�Ӻ�����ϵ��"&chr(10)&"�������û�"&dvbbs.membername&"�������͵Ķ��ţ�"&chr(10)&""&message&"',"&SqlNowString&",0,1)")
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
		ErrCodes=ErrCodes+"<li>�Բ��𣬸���������ߣ��ƶ�����ͨ�ֻ����ܻ������š�"
		Exit Sub
	End If

	Set Rs=Dvbbs.execute("select top 1 * from Dv_ChallengeInfo")
	Dim MyForumID,MouseID
	MyForumID=rs("D_ForumID")
	MouseID=rs("D_username")
%>
�����ύ���ݣ����Ժ󡭡�
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
			ErrCodes=ErrCodes+"<li>������̳����վ�ڶ���ʧ�ܡ�"
			Exit Sub
		Case Else
			ErrCodes=ErrCodes+"<li>�Ƿ����ύ���̡�"
		Exit Sub
	End Select
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>�ɹ������ɹ��ķ�����վ�ڶ���</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li><a href="challenge_msg.asp">���ط���ҳ��</a></li></ul>
</td></tr>
</table>
<%
End Sub

'�����û�����֪ͨ��Ϣ���¶�������||�¶�ѶID||����������
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

'ͳ������
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
'��ʾ������Ϣ
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