<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.Stats="高级用户修改密码"	
if Dvbbs.UserID=0 then Response.redirect "showerr.asp?ErrCodes=<li>您还没有登录，请登录后进行操作。&action=OtherErr"

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(9)=1) Then Response.redirect "showerr.asp?ErrCodes=<li>本论坛没有开启阳光会员注册、修改资料和密码的功能。&action=OtherErr"

Dvbbs.Loadtemplates("")
Dvbbs.Nav()

Select case request("action")
	case "submod"
		Dvbbs.stats="提交资料"
		Dvbbs.Head_var 0,0,"阳光会员修改信息","challenge_mod_pw.asp"
		call reg_2()
	Case else
		Dvbbs.stats="修改资料"
		Dvbbs.Head_var 0,0,"阳光会员修改信息","challenge_mod_pw.asp"
		call reg_1()
End select

Dvbbs.Footer()

sub reg_1()
dim rs
set rs=dvbbs.execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 or isnull(rs("IsChallenge")) then Response.redirect "showerr.asp?ErrCodes=<li>您还不是阳光会员，请先升级为阳光会员。&action=OtherErr"

%>
<FORM name=theForm action="challenge_mod_pw.asp?action=submod" method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY>
<TR align=middle>
<Th colSpan=2 height=24>高级用户修改密码资料</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>用户名</B>：</TD>
<TD width=60%  class=tablebody1>
<%=Rs("Username")%></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>手机号码</B>：</TD>
<TD class=tablebody1>
<%=rs("usermobile")%>
</TD>
</TR>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="提 交"></td></form></tr>
</tbody>
</table>
</form>
<%

set rs=nothing
end sub

sub reg_2()
'if request("mobile")="" then
'	founderr=true
'	Dvbbs.AddErrmsg "请输入您的手机号。"
'	exit sub
'end if
dim rs
set rs=Dvbbs.Execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 or isnull(rs("IsChallenge")) then Response.redirect "showerr.asp?ErrCodes=<li>您还不是阳光会员，请先升级为阳光会员。&action=OtherErr"
dim mobile
mobile=rs("UserMobile")


%>
正在提交数据，请稍后……
<form name="redir" action="http://bbs.ray5198.com/user_update_paypwd.jsp" method="post">
<INPUT type=hidden name="mobile" value="<%=mobile%>">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
set rs=nothing
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","§§§")
	checkreal=w
end if
end function
%>