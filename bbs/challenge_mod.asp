<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.stats="�����Ա�޸���Ϣ"
Dvbbs.Loadtemplates("")
Dvbbs.nav()
If Dvbbs.UserID=0 then Response.redirect "showerr.asp?ErrCodes=<li>����û�е�¼�����¼����в�����&action=OtherErr"

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(9)=1) Then Response.redirect "showerr.asp?ErrCodes=<li>����̳û�п��������Աע�ᡢ�޸����Ϻ�����Ĺ��ܡ�&action=OtherErr"

Select case request("action")
	case "submod"
		dvbbs.stats="�ύ����"
		Dvbbs.Head_var 0,0,"�����Ա�޸���Ϣ","challenge_mod.asp"
		call reg_2()
	Case "redir"
		dvbbs.stats="�ύ����"
		Dvbbs.Head_var 0,0,"�����Ա�޸���Ϣ","challenge_mod.asp"
		call redir()
	Case Else
		dvbbs.stats="�޸�����"
		Dvbbs.Head_var 0,0,"�����Ա�޸���Ϣ","challenge_mod.asp"
		call reg_1()
End Select 

Dvbbs.footer()

sub reg_1()
Dim rs,UserIM
set rs=dvbbs.execute("select * from [dv_user] where userid="&Dvbbs.userid)
If rs("IsChallenge")=0 Or Isnull(rs("IsChallenge")) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������Ա����������Ϊ�����Ա��&action=OtherErr"
UserIM = Split(Rs("UserIM"),"|||")
%>
<FORM name=theForm action="challenge_mod.asp?action=submod" method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY>
<TR align=middle>
<Th colSpan=2 height=24>�����û��޸�����</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>�û���</B>��</TD>
<TD width=60%  class=tablebody1>
<%=rs("username")%></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>ԭ��̳����</B>��<BR>
�������룡
</TD>
<TD width=60% class=tablebody1>
<INPUT type=password maxLength=16 size=30 name="psw">
</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>����̳����</B>��<BR>���޸�������</TD>
<TD class=tablebody1>
<INPUT type=password maxLength=16 size=30 name="pswc">
</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>ȷ��������</B>��<BR>���޸�������</TD>
<TD class=tablebody1>
<INPUT type=password maxLength=16 size=30 name="pswc2">
</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>�ֻ�����</B>��</TD>
<TD class=tablebody1>
<%=rs("usermobile")%>
</TD>
</TR>
<TR>
<TD width=40%  class=tablebody1><B>�Ա�</B>��<BR>��ѡ�������Ա�</font></TD>
<TD width=60%  class=tablebody1> <input type="radio" value="1" name="Sex" <%if rs("usersex")=1 Then Response.Write "checked"%>>���
<input type="radio" name="Sex" value="0" <%if rs("usersex")=0 Then Response.Write "checked"%>>����</TD>
</TR>
<TR>
<TD width=40%  class=tablebody1><B>�ϣɣã�</B>��</TD>
<TD class=tablebody1>
<INPUT type=text size=30 name="oicq" value="<%=UserIM(1)%>">
</TD>
</TR>
<TR>
<TD width=40%  class=tablebody1><B>Email��ַ</B>��<BR>��������Ч���ʼ���ַ���⽫ʹ�����õ���̳�е����й���</font></TD>
<TD width=60%  class=tablebody1>
<INPUT maxLength=50 size=30 name="email" value="<%=rs("useremail")%>"></TD>
</TR>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="�� ��"></td></form></tr>
</tbody>
</table>
</form>
<%
Set Rs=Nothing
end sub

sub reg_2()
dim rs
if request("email")="" then Response.redirect "showerr.asp?ErrCodes=<li>�����������ʼ���ַ��&action=OtherErr"
if request("psw")="" then Response.redirect "showerr.asp?ErrCodes=<li>������������̳���롣&action=OtherErr"
if request("pswc") <> request("pswc2") then Response.redirect "showerr.asp?ErrCodes=<li>��������������벻һ�£����������롣&action=OtherErr"

set rs=dvbbs.execute("select * from [dv_user] where userid="&dvbbs.userid)
if md5(trim(request("psw")),16)<>rs("userpassword") then Response.redirect "showerr.asp?ErrCodes=<li>���������̳���벻��ȷ�����������롣&action=OtherErr"
if rs("IsChallenge")=0 or isnull(rs("IsChallenge")) then Response.redirect "showerr.asp?ErrCodes=<li>�������������Ա����������Ϊ�����Ա��&action=OtherErr"
dim mobile
mobile=rs("UserMobile")

dim newpsw
if request("pswc")="" then
	newpsw=request("psw")
else
	newpsw=request("pswc")
end if
dim sex
if cint(request("sex"))=1 then
	sex="F"
else
	sex="M"
end if

Session("challenge_mod_temp")=checkreal(request("psw")) & "|||" & checkreal(request("pswc")) & "|||" & checkreal(request("mobile")) & "|||" & checkreal(request("sex")) & "|||" & checkreal(request("oicq")) & "|||" & checkreal(request("email"))


'��ս�����
Dim MaxUserID,MaxLength
MaxLength=12
set rs=Dvbbs.Execute("select Max(userid) from [dv_user]")
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

%>
�����ύ���ݣ����Ժ󡭡�
<form name="redir" action="http://bbs.ray5198.com/user_update.jsp" method="post">
<INPUT type=hidden name="username" value="<%=dvbbs.membername%>">
<INPUT type=hidden name="forumPwd" value="<%=checkreal(newpsw)%>">
<INPUT type=hidden name="oldPwd" value="<%=checkreal(request("psw"))%>">
<INPUT type=hidden name="mobile" value="<%=mobile%>">
<INPUT type=hidden name="sex" value="<%=sex%>">
<INPUT type=hidden name="qq" value="<%=checkreal(request("oicq"))%>">
<INPUT type=hidden name="email" value="<%=checkreal(request("email"))%>">
<INPUT type=hidden name="forumId" value="<%=MyForumID%>">
<input type=hidden name="forumUrl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<input type=hidden value="<%=MaxUserID%>" name="challengeWord">
<input type=hidden value="challenge_mod.asp?action=redir" name="dirPage">
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
dim ErrorCode,ErrorMsg
dim remobile,rechallengeWord,retokerWord
dim challengeWord_key,rechallengeWord_key
dim challenge_mod_temp,newpsw

ErrorCode=trim(request("ErrorCode"))
ErrorMsg=trim(request("ErrorMsg"))
remobile=trim(dvbbs.CheckStr(request("mobile")))
rechallengeWord=trim(dvbbs.CheckStr(request("challengeWord")))
retokerWord=trim(request("tokenWord"))

select case ErrorCode
case 100
	challengeWord_key=session("challengeWord_key")
	if challengeWord_key=retokerWord then
		challenge_mod_temp=split(dvbbs.CheckStr(Session("challenge_mod_temp")),"|||")
		if trim(challenge_mod_temp(1))="" then
			newpsw=md5(trim(challenge_mod_temp(0)),16)
		else
			newpsw=md5(trim(challenge_mod_temp(1)),16)
		end if
		Dim UserIM,Rs,MyUserIM
		Set Rs=Dvbbs.Execute("Select UserIM From Dv_User Where UserID = " & Dvbbs.UserID)
		UserIM = Split(Rs("UserIM"),"|||")
		MyUserIM = UserIM(0) & "|||" & challenge_mod_temp(4) & "|||" & UserIM(2) & "|||" & UserIM(3) & "|||" & UserIM(4) & "|||" & UserIM(5) & "|||" & UserIM(6)
		dvbbs.execute("update [dv_user] set userpassword='"&newpsw&"',usersex='"&challenge_mod_temp(3)&"',UserIM='"&Replace(MyUserIM,"'","''")&"',useremail='"&challenge_mod_temp(5)&"',UserMobile='"&remobile&"',IsChallenge=1 where userid="&Dvbbs.UserID)
	else
		Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ����ύ����1��&action=OtherErr"
	end if
case 101
	Response.redirect "showerr.asp?ErrCodes=<li>������̳���������޸���Ϣʧ�ܣ�"&ErrorMsg&"��&action=OtherErr"
case 102
	Response.redirect "showerr.asp?ErrCodes=<li>������̳���������޸���Ϣʧ�ܣ�"&ErrorMsg&"��&action=OtherErr"
case else
	Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ����ύ����2��"&ErrorMsg&"��&action=OtherErr"
end select
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>�޸ĳɹ���<%=Dvbbs.Forum_Info(0)%>��ӭ���ĵ���</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li>���ڱ�վ�ɹ����޸��˳����û�����<br><li><a href="index.asp">����������</a></li></ul>
</td></tr>
</table>
<%
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","����")
	checkreal=w
end if
end function
%>