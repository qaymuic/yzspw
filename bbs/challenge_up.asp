<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.Stats="�û�����"
Dvbbs.Loadtemplates("")
Dvbbs.Nav()
If Dvbbs.UserID=0 then Response.redirect "showerr.asp?ErrCodes=<li>����û�е�¼�����¼����в�����&action=OtherErr"

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(9)=1) Then Response.redirect "showerr.asp?ErrCodes=<li>����̳û�п��������Աע�ᡢ�޸����Ϻ�����Ĺ��ܡ�&action=OtherErr"

Select case request("action")
	case "submobile"
		dvbbs.stats="�ύ����"
		Dvbbs.Head_var 0,0,"��ͨ�û�����","challenge_up.asp"
		call reg_2()
	Case "redir"
		dvbbs.stats="�ύ����"
		Dvbbs.Head_var 0,0,"��ͨ�û�����","challenge_up.asp"
		call redir()
	Case else
		dvbbs.stats="��������"
		Dvbbs.Head_var 0,0,"��ͨ�û�����","challenge_up.asp"
		call reg_1()
End Select

Dvbbs.Footer()

sub reg_1()
dim rs
set rs=dvbbs.execute("select IsChallenge from [dv_user] where userid="&Dvbbs.userid)
if rs(0)=1 Then Response.redirect "showerr.asp?ErrCodes=<li>���Ѿ��Ǹ߼��û��������Ҫ<a href=challenge_mod.asp>�޸�����������������</a>��&action=OtherErr"
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr><th align=center colspan=2><form action="challenge_up.asp?action=submobile" method=post>��ͨ�û�����Ϊ�߼��û�</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>������������̳����</B>��</td>
	<td class=tablebody1 width="60%">
<input type=password size=30 name="password">
	</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>�����������ֻ�����</B>��</td>
	<td class=tablebody1 width="60%">
<input type=text size=30 name="mobile">
	</td></tr>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="�� ��"></td></form></tr>
</table>
<%
end sub

sub reg_2()
dim rs
if request("mobile")="" then Response.redirect "showerr.asp?ErrCodes=<li>�����������ֻ��š�&action=OtherErr"

if request("password")="" then Response.redirect "showerr.asp?ErrCodes=<li>������������̳���롣&action=OtherErr"

'��ս�����
Dim MaxUserID,MaxLength
MaxLength=12
set rs=dvbbs.execute("select Max(userid) from [dv_user]")
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

set rs=dvbbs.execute("select top 1 * from Dv_ChallengeInfo")
Dim MyForumID
MyForumID=rs("D_ForumID")

Set  Rs=dvbbs.execute("select * from [dv_user] where Usermobile='"&Dvbbs.CheckStr(request("mobile"))&"' and IsChallenge=1")
If Not (rs.bof and rs.bof) Then Response.redirect "showerr.asp?ErrCodes=<li>��ʹ�õ��ֻ��ű����Ѿ��ڱ���̳ʹ�ã���ȷ�����ڱ���̳�Ƿ��������ʺ�ʹ�ñ��ֻ��š�&action=OtherErr"

Dim UserIM
set rs=dvbbs.execute("select * from [dv_user] where userid="&Dvbbs.userid)
if md5(trim(request("password")),16) <> rs("userpassword") then Response.redirect "showerr.asp?ErrCodes=<li>���������̳���벻��ȷ�����������롣&action=OtherErr"
dim sex
if cint(rs("usersex"))=1 then
	sex="F"
else
	sex="M"
end if
UserIM = Split(Rs("UserIM"),"|||")
%>
�����ύ���ݣ����Ժ󡭡�
<form name="redir" action="http://bbs.ray5198.com/user_upgrade.jsp" method="post">
<INPUT type=hidden name="username" value="<%=Dvbbs.membername%>">
<INPUT type=hidden name="forumPwd" value="<%=request("password")%>">
<INPUT type=hidden name="mobile" value="<%=request("mobile")%>">
<INPUT type=hidden name="sex" value="<%=sex%>">
<INPUT type=hidden name="qq" value="<%=UserIM(1)%>">
<INPUT type=hidden name="email" value="<%=rs("useremail")%>">
<INPUT type=hidden name="forumId" value="<%=MyForumID%>">
<input type=hidden name="forumUrl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<input type=hidden value="<%=MaxUserID%>" name="challengeWord">
<input type=hidden value="challenge_up.asp?action=redir" name="dirPage">
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
dim remobile,rechallengeWord,retokerWord
dim challengeWord_key,rechallengeWord_key

ErrorCode=trim(request("ErrorCode"))
ErrorMsg=trim(request("ErrorMsg"))
remobile=trim(Dvbbs.CheckStr(request("mobile")))
rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
retokerWord=trim(request("tokenWord"))

select case ErrorCode
case 100
	challengeWord_key=session("challengeWord_key")
	if challengeWord_key=retokerWord then
		Dvbbs.Execute("update [dv_user] set UserMobile='"&remobile&"',IsChallenge=1 where userid="&Dvbbs.UserID)
	else
		Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ����ύ���̡�&action=OtherErr"
	end if
case 101
	Response.redirect "showerr.asp?ErrCodes=<li>��������Ϊ�����Աע��ʧ�ܡ�&action=OtherErr"
case else
	Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ����ύ���̡�&action=OtherErr"
end select
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>ע��ɹ���<%=Dvbbs.Forum_Info(0)%>��ӭ���ĵ���</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li>���ڱ�վ�ɹ���ע���Ϊ�߼��û�<br><li><a href="index.asp">����������</a></li></ul>
</td></tr>
</table>
<%
Session(Dvbbs.CacheName & "UserID")=Empty
end sub
%>