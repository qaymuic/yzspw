<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<%
Dvbbs.LoadTemplates("")
If Dvbbs.UserID=0 Then
	Dvbbs.AddErrCode(24)
	Dvbbs.showerr()
End If
If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(8)=1) Then
	Response.redirect "showerr.asp?ErrCodes=<li>本论坛没有开启VIP收费论坛功能。&action=OtherErr"
End If

dvbbs.stats="交费进入认证论坛"
Dvbbs.Nav()
Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
'GetBoardPermission
Dvbbs.Showerr()
Select Case request("action")
	Case "subinfo"
		call subinfo()
	Case "redir"
		call redir()
	Case Else
		call inputmyinfo()
End Select 
Dvbbs.activeonline()
Dvbbs.footer()

sub inputmyinfo()
dim rs
dim mobile
set rs=dvbbs.execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 Or IsNull(Rs("IsChallenge")) then
	Response.redirect "showerr.asp?ErrCodes=<li>您不是本站的高级用户，请先<a href=challenge_up.asp>升级成为高级用户</a>。&action=OtherErr"
	exit sub
end if
mobile=rs("usermobile")
set rs=nothing
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr><th align=center colspan=2><form action="pay_boardlimited.asp?action=subinfo&boardid=<%=Dvbbs.BoardID%>" method=post>高级用户订阅认证论坛</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>请输入您的手机号码</B>：</td>
	<td class=tablebody1 width="60%">
<%=mobile%>
	</td></tr>
<tr><td align=center class=tablebody1 colspan=2>访问该VIP版面规则为：得到 <B><%=Dvbbs.Board_Setting(46)%></B> 天的访问该VIP版面的权限并花费您 <B><%=Dvbbs.Board_Setting(20)/100%></B> 的魔力水晶球</td></tr>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="确 认"></td></form></tr>
</table>
<%
end sub

sub subinfo()
dim mobile
dim rs
set rs=dvbbs.execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 Or IsNull(Rs("IsChallenge")) then
	Response.redirect "showerr.asp?ErrCodes=<li>您不是本站的高级用户，请先<a href=challenge_up.asp>升级成为高级用户</a>。&action=OtherErr"
	exit sub
end if
mobile=rs("usermobile")

'挑战随机数
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
Dim MyForumID,MyForumUrl,MyAdminID
MyForumID=rs("D_ForumID")
MyForumUrl=rs("D_Forumurl")
MyAdminID=rs("D_Username")

dim vipid
vipid=Dvbbs.Board_Setting(47) & Dvbbs.BoardID
vipid=md5(vipid,32)

%>
正在提交数据，请稍后……
<form name="redir" action="http://bbs.ray5198.com/rayvipforum_magicgarden/vipforum/vipmember_resign.jsp" method="post">
<INPUT type=hidden name="usermobile" value="<%=mobile%>">
<INPUT type=hidden name="vipusetime" value="<%=dvbbs.Board_Setting(46)%>">
<INPUT type=hidden name="challengword" value="<%=Session("challengeWord")%>">
<INPUT type=hidden name="forumid" value="<%=MyForumID%>">
<INPUT type=hidden name="vipid" value="<%=vipid%>">
<INPUT type=hidden name="viptransurl" value="<%=Dvbbs.Get_ScriptNameUrl()%>">
<input type=hidden value="pay_boardlimited.asp?boardid=<%=dvbbs.boardid%>&action=redir" name="viptranspage">
<INPUT type=hidden name="vipproveurl" value="<%=Dvbbs.Get_ScriptNameUrl()%>">
<input type=hidden value="Challenge_Scan_Board.asp?BoardID=<%=Dvbbs.BoardID%>" name="vipprovepage">
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
dim remobile,rechallengeWord,retokerWord,repayid
dim challengeWord_key,rechallengeWord_key

ErrorCode=trim(request("errcode"))
ErrorMsg=trim(request("errmessage"))
remobile=trim(Dvbbs.CheckStr(request("mobile")))
'repayid=trim(Dvbbs.CheckStr(request("payid")))
rechallengeWord=trim(Dvbbs.CheckStr(request("challengWord")))
retokerWord=trim(request("tokenword"))

select case ErrorCode
case 1000
	challengeWord_key=session("challengeWord_key")
	if challengeWord_key=retokerWord then
		'type=1订阅主题，type=2订阅论坛
		'apply=1交易成功,主服务器issuc=0正在审核中
		'dvbbs.execute("insert into DV_ChanOrders (O_type,O_mobile,O_Username,O_isApply,O_issuc,O_PayMoney,O_Paycode,O_BoardID) values (2,'"&remobile&"','"&dvbbs.membername&"',1,1,"&dvbbs.Board_Setting(20)&",'"&repayid&"',"&dvbbs.boardid&")")
		'Response.Write "<script language=""javascript"">Dvbbs_suc('申请成功，<a href=list.asp?boardid="&dvbbs.boardid&">点击连接进入VIP版面</a>。',-3)</script>"
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>申请订阅认证论坛成功</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li><a href="list.asp?boardid=<%=Dvbbs.BoardID%>">进入讨论区</a></li></ul>
</td></tr>
</table>
<%
		exit sub
	else
		Response.redirect "showerr.asp?ErrCodes=<li>非法的提交过程。&action=OtherErr"
		exit sub
	end if
case 1001
	Response.redirect "showerr.asp?ErrCodes=<li>VIP没有审核通过，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1002
	Response.redirect "showerr.asp?ErrCodes=<li>用户积分不够，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1003
	Response.redirect "showerr.asp?ErrCodes=<li>用户不是阳光会员，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1004
	Response.redirect "showerr.asp?ErrCodes=<li>论坛数据不合法，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1005
	Response.redirect "showerr.asp?ErrCodes=<li>论坛ID不存在，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1006
	Response.redirect "showerr.asp?ErrCodes=<li>站长用户名或密码不正确，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1007
	Response.redirect "showerr.asp?ErrCodes=<li>VIP论坛已经申请正处于使用状态，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1008
	Response.redirect "showerr.asp?ErrCodes=<li>不是有效的论坛，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1009
	Response.redirect "showerr.asp?ErrCodes=<li>VIP论坛申请失败，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1010
	Response.redirect "showerr.asp?ErrCodes=<li>数据操作失败，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1011
	Response.redirect "showerr.asp?ErrCodes=<li>不明原因与管理员联系，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1012
	Response.redirect "showerr.asp?ErrCodes=<li>积分超过上限，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1013
	Response.redirect "showerr.asp?ErrCodes=<li>提供的挑战随机数是空，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case 1014
	Response.redirect "showerr.asp?ErrCodes=<li>你已经申请过VIP论坛,但是VIP论坛没有被激活，请等待，"&ErrorMsg&"。&action=OtherErr"
	exit sub
case else
	Response.redirect "showerr.asp?ErrCodes=<li>非法的提交过程，"&ErrorMsg&"。&action=OtherErr"
	exit sub
end select
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>申请订阅认证论坛成功</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li>您的VIP论坛访问资格正在审批中，请等待通知<br><li><a href="index.asp">进入讨论区</a></li></ul>
</td></tr>
</table>
<%
end sub
%>