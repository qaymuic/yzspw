<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.stats="阳光会员修改信息"
Dvbbs.Loadtemplates("")
Dvbbs.nav()
If Dvbbs.UserID=0 then Response.redirect "showerr.asp?ErrCodes=<li>您还没有登录，请登录后进行操作。&action=OtherErr"

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(9)=1) Then Response.redirect "showerr.asp?ErrCodes=<li>本论坛没有开启阳光会员注册、修改资料和密码的功能。&action=OtherErr"

Select case request("action")
	case "submod"
		dvbbs.stats="提交资料"
		Dvbbs.Head_var 0,0,"阳光会员修改信息","challenge_mod.asp"
		call reg_2()
	Case "redir"
		dvbbs.stats="提交资料"
		Dvbbs.Head_var 0,0,"阳光会员修改信息","challenge_mod.asp"
		call redir()
	Case Else
		dvbbs.stats="修改资料"
		Dvbbs.Head_var 0,0,"阳光会员修改信息","challenge_mod.asp"
		call reg_1()
End Select 

Dvbbs.footer()

sub reg_1()
Dim rs,UserIM
set rs=dvbbs.execute("select * from [dv_user] where userid="&Dvbbs.userid)
If rs("IsChallenge")=0 Or Isnull(rs("IsChallenge")) Then Response.redirect "showerr.asp?ErrCodes=<li>您还不是阳光会员，请先升级为阳光会员。&action=OtherErr"
UserIM = Split(Rs("UserIM"),"|||")
%>
<FORM name=theForm action="challenge_mod.asp?action=submod" method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY>
<TR align=middle>
<Th colSpan=2 height=24>超级用户修改资料</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>用户名</B>：</TD>
<TD width=60%  class=tablebody1>
<%=rs("username")%></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>原论坛密码</B>：<BR>
必须输入！
</TD>
<TD width=60% class=tablebody1>
<INPUT type=password maxLength=16 size=30 name="psw">
</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>新论坛密码</B>：<BR>不修改请留空</TD>
<TD class=tablebody1>
<INPUT type=password maxLength=16 size=30 name="pswc">
</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>确认新密码</B>：<BR>不修改请留空</TD>
<TD class=tablebody1>
<INPUT type=password maxLength=16 size=30 name="pswc2">
</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>手机号码</B>：</TD>
<TD class=tablebody1>
<%=rs("usermobile")%>
</TD>
</TR>
<TR>
<TD width=40%  class=tablebody1><B>性别</B>：<BR>请选择您的性别</font></TD>
<TD width=60%  class=tablebody1> <input type="radio" value="1" name="Sex" <%if rs("usersex")=1 Then Response.Write "checked"%>>酷哥
<input type="radio" name="Sex" value="0" <%if rs("usersex")=0 Then Response.Write "checked"%>>靓妹</TD>
</TR>
<TR>
<TD width=40%  class=tablebody1><B>ＯＩＣＱ</B>：</TD>
<TD class=tablebody1>
<INPUT type=text size=30 name="oicq" value="<%=UserIM(1)%>">
</TD>
</TR>
<TR>
<TD width=40%  class=tablebody1><B>Email地址</B>：<BR>请输入有效的邮件地址，这将使您能用到论坛中的所有功能</font></TD>
<TD width=60%  class=tablebody1>
<INPUT maxLength=50 size=30 name="email" value="<%=rs("useremail")%>"></TD>
</TR>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="提 交"></td></form></tr>
</tbody>
</table>
</form>
<%
Set Rs=Nothing
end sub

sub reg_2()
dim rs
if request("email")="" then Response.redirect "showerr.asp?ErrCodes=<li>请输入您的邮件地址。&action=OtherErr"
if request("psw")="" then Response.redirect "showerr.asp?ErrCodes=<li>请输入您的论坛密码。&action=OtherErr"
if request("pswc") <> request("pswc2") then Response.redirect "showerr.asp?ErrCodes=<li>两次输入的新密码不一致，请重新输入。&action=OtherErr"

set rs=dvbbs.execute("select * from [dv_user] where userid="&dvbbs.userid)
if md5(trim(request("psw")),16)<>rs("userpassword") then Response.redirect "showerr.asp?ErrCodes=<li>您输入的论坛密码不正确，请重新输入。&action=OtherErr"
if rs("IsChallenge")=0 or isnull(rs("IsChallenge")) then Response.redirect "showerr.asp?ErrCodes=<li>您还不是阳光会员，请先升级为阳光会员。&action=OtherErr"
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


'挑战随机数
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
正在提交数据，请稍后……
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
		Response.redirect "showerr.asp?ErrCodes=<li>非法的提交过程1。&action=OtherErr"
	end if
case 101
	Response.redirect "showerr.asp?ErrCodes=<li>您在论坛超级联盟修改信息失败，"&ErrorMsg&"。&action=OtherErr"
case 102
	Response.redirect "showerr.asp?ErrCodes=<li>您在论坛超级联盟修改信息失败，"&ErrorMsg&"。&action=OtherErr"
case else
	Response.redirect "showerr.asp?ErrCodes=<li>非法的提交过程2，"&ErrorMsg&"。&action=OtherErr"
end select
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>修改成功：<%=Dvbbs.Forum_Info(0)%>欢迎您的到来</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li>您在本站成功的修改了超级用户资料<br><li><a href="index.asp">进入讨论区</a></li></ul>
</td></tr>
</table>
<%
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","§§§")
	checkreal=w
end if
end function
%>