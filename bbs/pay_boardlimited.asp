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
	Response.redirect "showerr.asp?ErrCodes=<li>����̳û�п���VIP�շ���̳���ܡ�&action=OtherErr"
End If

dvbbs.stats="���ѽ�����֤��̳"
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
	Response.redirect "showerr.asp?ErrCodes=<li>�����Ǳ�վ�ĸ߼��û�������<a href=challenge_up.asp>������Ϊ�߼��û�</a>��&action=OtherErr"
	exit sub
end if
mobile=rs("usermobile")
set rs=nothing
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr><th align=center colspan=2><form action="pay_boardlimited.asp?action=subinfo&boardid=<%=Dvbbs.BoardID%>" method=post>�߼��û�������֤��̳</td></tr>
<tr><td class=tablebody1 align=right width="40%"><B>�����������ֻ�����</B>��</td>
	<td class=tablebody1 width="60%">
<%=mobile%>
	</td></tr>
<tr><td align=center class=tablebody1 colspan=2>���ʸ�VIP�������Ϊ���õ� <B><%=Dvbbs.Board_Setting(46)%></B> ��ķ��ʸ�VIP�����Ȩ�޲������� <B><%=Dvbbs.Board_Setting(20)/100%></B> ��ħ��ˮ����</td></tr>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="ȷ ��"></td></form></tr>
</table>
<%
end sub

sub subinfo()
dim mobile
dim rs
set rs=dvbbs.execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 Or IsNull(Rs("IsChallenge")) then
	Response.redirect "showerr.asp?ErrCodes=<li>�����Ǳ�վ�ĸ߼��û�������<a href=challenge_up.asp>������Ϊ�߼��û�</a>��&action=OtherErr"
	exit sub
end if
mobile=rs("usermobile")

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
Dim MyForumID,MyForumUrl,MyAdminID
MyForumID=rs("D_ForumID")
MyForumUrl=rs("D_Forumurl")
MyAdminID=rs("D_Username")

dim vipid
vipid=Dvbbs.Board_Setting(47) & Dvbbs.BoardID
vipid=md5(vipid,32)

%>
�����ύ���ݣ����Ժ󡭡�
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
		'type=1�������⣬type=2������̳
		'apply=1���׳ɹ�,��������issuc=0���������
		'dvbbs.execute("insert into DV_ChanOrders (O_type,O_mobile,O_Username,O_isApply,O_issuc,O_PayMoney,O_Paycode,O_BoardID) values (2,'"&remobile&"','"&dvbbs.membername&"',1,1,"&dvbbs.Board_Setting(20)&",'"&repayid&"',"&dvbbs.boardid&")")
		'Response.Write "<script language=""javascript"">Dvbbs_suc('����ɹ���<a href=list.asp?boardid="&dvbbs.boardid&">������ӽ���VIP����</a>��',-3)</script>"
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>���붩����֤��̳�ɹ�</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li><a href="list.asp?boardid=<%=Dvbbs.BoardID%>">����������</a></li></ul>
</td></tr>
</table>
<%
		exit sub
	else
		Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ����ύ���̡�&action=OtherErr"
		exit sub
	end if
case 1001
	Response.redirect "showerr.asp?ErrCodes=<li>VIPû�����ͨ����"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1002
	Response.redirect "showerr.asp?ErrCodes=<li>�û����ֲ�����"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1003
	Response.redirect "showerr.asp?ErrCodes=<li>�û����������Ա��"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1004
	Response.redirect "showerr.asp?ErrCodes=<li>��̳���ݲ��Ϸ���"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1005
	Response.redirect "showerr.asp?ErrCodes=<li>��̳ID�����ڣ�"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1006
	Response.redirect "showerr.asp?ErrCodes=<li>վ���û��������벻��ȷ��"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1007
	Response.redirect "showerr.asp?ErrCodes=<li>VIP��̳�Ѿ�����������ʹ��״̬��"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1008
	Response.redirect "showerr.asp?ErrCodes=<li>������Ч����̳��"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1009
	Response.redirect "showerr.asp?ErrCodes=<li>VIP��̳����ʧ�ܣ�"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1010
	Response.redirect "showerr.asp?ErrCodes=<li>���ݲ���ʧ�ܣ�"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1011
	Response.redirect "showerr.asp?ErrCodes=<li>����ԭ�������Ա��ϵ��"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1012
	Response.redirect "showerr.asp?ErrCodes=<li>���ֳ������ޣ�"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1013
	Response.redirect "showerr.asp?ErrCodes=<li>�ṩ����ս������ǿգ�"&ErrorMsg&"��&action=OtherErr"
	exit sub
case 1014
	Response.redirect "showerr.asp?ErrCodes=<li>���Ѿ������VIP��̳,����VIP��̳û�б������ȴ���"&ErrorMsg&"��&action=OtherErr"
	exit sub
case else
	Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ����ύ���̣�"&ErrorMsg&"��&action=OtherErr"
	exit sub
end select
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>���붩����֤��̳�ɹ�</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li>����VIP��̳�����ʸ����������У���ȴ�֪ͨ<br><li><a href="index.asp">����������</a></li></ul>
</td></tr>
</table>
<%
end sub
%>