<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!--#include file="inc/md5.asp"-->
<%
dim AnnounceID
Dvbbs.Loadtemplates("")
Dvbbs.stats="��������"
Dvbbs.Nav()
GetBoardPermission
if Dvbbs.UserID=0 then Response.redirect "showerr.asp?ErrCodes=<li>����û�е�¼�����¼����в�����&action=OtherErr"
If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(7)=1) Then Response.redirect "showerr.asp?ErrCodes=<li>����̳û�п����ֻ����Ŷ������⹦�ܡ�&action=OtherErr"

If request("id")="" Then
	Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ������Ӳ�����&action=OtherErr"
ElseIf Not IsNumeric(request("id")) Then
	Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ������Ӳ�����&action=OtherErr"
Else
	AnnounceID=Clng(request("id"))
End If

Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""

Select Case request("action")
	Case "subinfo"
		call subinfo()
	Case "redir"
		call redir()
	Case Else
		Call inputmyinfo()
End select

dvbbs.footer()

Sub inputmyinfo()
Dim Rs
Set Rs=dvbbs.execute("select IsChallenge,usermobile from [dv_user] where userid="&dvbbs.userid)
If Rs("IsChallenge")<>1 Or IsNull(Rs("IsChallenge"))Then Response.redirect "showerr.asp?ErrCodes=<li>�������������Ա����������Ϊ�����Ա��&action=OtherErr"

%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr><th align=center colspan=2><form action="pay_topic.asp?action=subinfo" method=post>�߼��û���������</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>�����ֻ�������</B>��</td>
<td class=tablebody1 width="60%">
<%=Rs("UserMobile")%>
<input type=hidden value="<%=Dvbbs.BoardID%>" name="boardid">
<input type=hidden value="<%=AnnounceID%>" name="id">
</td></tr>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="�� ��"></td></form></tr>
</table>
<%

Set rs=nothing
End Sub

Sub subinfo()
Dim Rs
Set Rs=dvbbs.execute("select IsChallenge,usermobile from [dv_user] where userid="&dvbbs.userid)
If Rs("IsChallenge")<>1 Or IsNull(Rs("IsChallenge"))Then Response.redirect "showerr.asp?ErrCodes=<li>�������������Ա����������Ϊ�����Ա��&action=OtherErr"
Dim mobile
mobile=rs("usermobile")
Set rs=nothing
Dim topic
Set rs=dvbbs.execute("select * from dv_topic where BoardID="&Dvbbs.BoardID&" And topicid="&announceid)
If Rs.EOF And Rs.BOF Then
	Response.redirect "showerr.asp?ErrCodes=<li>��Ҫ���ĵ����Ⲣ�����ڡ����Ѿ���ɾ������������ȷ�����ύ����Ϣ�Ƿ���ȷ��&action=OtherErr"
Else
	topic=rs("title")
End If

'��ս�����
Dim MaxUserID,MaxLength
MaxLength=12
set Rs=dvbbs.execute("select Max(userid) from [dv_user]")
MaxUserID=rs(0)

Dim num1,rndnum
Randomize
Do While Len(rndnum)<4
	num1=CStr(Chr((57-48)*rnd+48))
	rndnum=rndnum&num1
loop
MaxUserID=rndnum & MaxUserID
MaxLength=MaxLength-len(MaxUserID)
select case MaxLength
case 7
	MaxUserID="0000000" & MaxUserID
case 6
	MaxUserID="000000" & MaxUserID
case 5
	MaxUserID="00000" & MaxUserID
case 4
	MaxUserID="0000" & MaxUserID
case 3
	MaxUserID="000" & MaxUserID
case 2
	MaxUserID="00" & MaxUserID
case 1
	MaxUserID="0" & MaxUserID
case 0
	MaxUserID=MaxUserID
end select
Session("challengeWord")=MaxUserID

session("challengeWord_key")=md5(Session("challengeWord") & ":" & Dvbbs.CacheData(21,0),32)

Set Rs=Dvbbs.Execute("select top 1 * from Dv_ChallengeInfo")
Dim MyForumID,MyForumUrl,MyAdminID
MyForumID=rs("D_ForumID")
MyForumUrl=rs("D_Forumurl")
MyAdminID=rs("D_Username")

%>
�����ύ���ݣ����Ժ󡭡�
<form name="redir" action="http://bbs.ray5198.com/sub.jsp" method="post">
<INPUT type=hidden name="mobile" value="<%=mobile%>">
<INPUT type=hidden name="subjectId" value="<%=announceid%>">
<INPUT type=hidden name="subjectName" value="<%=topic%>">
<INPUT type=hidden name="forumId" value="<%=MyForumID%>">
<INPUT type=hidden name="forumUrl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<input type=hidden value="<%=MaxUserID%>" name="chanWord">
<input type=hidden value="pay_topic.asp?boardid=<%=dvbbs.boardid%>&id=<%=announceid%>&action=redir" name="dirPage">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
set rs=nothing
end sub


sub redir()
dim rs
dim ErrorCode,ErrorMsg
dim remobile,rechallengeWord,retokerWord,resubjectid
dim challengeWord_key,rechallengeWord_key

ErrorCode=trim(request("ErrorCode"))
ErrorMsg=trim(request("ErrorMsg"))
remobile=trim(Dvbbs.CheckStr(request("mobile")))
resubjectid=trim(Dvbbs.CheckStr(request("subjectid")))
rechallengeWord=trim(Dvbbs.CheckStr(request("chanWord")))
retokerWord=trim(request("tokenWord"))

'if not isnumeric(resubjectid) then
'	founderr=true
'	Dvbbs.AddErrmsg "�Ƿ��Ĳ���1��" & request("subjectid")
'	exit sub
'end if

dim smsuserlist
select case ErrorCode
case 100
	challengeWord_key=session("challengeWord_key")
	If challengeWord_key=retokerWord Then
		'type=1�������⣬type=2������̳
		Set Rs=Dvbbs.Execute("select username from [dv_user] where usermobile='"&remobile&"' and IsChallenge=1")
		If rs.eof and rs.bof Then
			Response.redirect "showerr.asp?ErrCodes=<li>����������ʧ�ܡ�&action=OtherErr"
		Else
			dvbbs.membername=rs(0)
		End If
		Set rs=dvbbs.execute("select * from dv_topic where topicid="&AnnounceID)
		If Not (rs.eof and rs.bof) Then
			smsuserlist=rs("smsuserlist")
			If IsNull(smsuserlist) or smsuserlist="" Then
				smsuserlist=dvbbs.membername
			Else
				If InStr("$" & lcase(smsuserlist) & "$","$" & lcase(dvbbs.membername) & "$")=0 Then
					smsuserlist=smsuserlist & "$" & dvbbs.membername
				End If
			End If
			Dvbbs.Execute("update dv_topic set smsuserlist='"&smsuserlist&"',issmstopic=1 where topicid="&AnnounceID)
		Else
			Response.redirect "showerr.asp?ErrCodes=<li>����������ʧ�ܡ�&action=OtherErr"
		end if
		Set Rs=Nothing
	Else
		Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ����ύ���̡�&action=OtherErr"
	End If
case 201
	Response.redirect "showerr.asp?ErrCodes=<li>���������������"&ErrorMsg&"��&action=OtherErr"
case 202
	Response.redirect "showerr.asp?ErrCodes=<li>���ظ�������������⣬"&ErrorMsg&"��&action=OtherErr"
case 203
	Response.redirect "showerr.asp?ErrCodes=<li>����������ʧ�ܣ�"&ErrorMsg&"��&action=OtherErr"
case Else
	Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ����ύ���̣�"&ErrorMsg&"��&action=OtherErr"
End Select
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>���붩������ɹ�</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li>��������ɹ��������������˻ظ�ʱ��ϵͳ�����Զ����ֻ�����֪ͨ��<br><li><a href="index.asp">����������</a></li></ul>
</td></tr>
</table>
<%
End Sub
%>