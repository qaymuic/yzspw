<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.Stats="�߼��û��޸�����"	
if Dvbbs.UserID=0 then Response.redirect "showerr.asp?ErrCodes=<li>����û�е�¼�����¼����в�����&action=OtherErr"

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(9)=1) Then Response.redirect "showerr.asp?ErrCodes=<li>����̳û�п��������Աע�ᡢ�޸����Ϻ�����Ĺ��ܡ�&action=OtherErr"

Dvbbs.Loadtemplates("")
Dvbbs.Nav()

Select case request("action")
	case "submod"
		Dvbbs.stats="�ύ����"
		Dvbbs.Head_var 0,0,"�����Ա�޸���Ϣ","challenge_mod_pw.asp"
		call reg_2()
	Case else
		Dvbbs.stats="�޸�����"
		Dvbbs.Head_var 0,0,"�����Ա�޸���Ϣ","challenge_mod_pw.asp"
		call reg_1()
End select

Dvbbs.Footer()

sub reg_1()
dim rs
set rs=dvbbs.execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 or isnull(rs("IsChallenge")) then Response.redirect "showerr.asp?ErrCodes=<li>�������������Ա����������Ϊ�����Ա��&action=OtherErr"

%>
<FORM name=theForm action="challenge_mod_pw.asp?action=submod" method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY>
<TR align=middle>
<Th colSpan=2 height=24>�߼��û��޸���������</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>�û���</B>��</TD>
<TD width=60%  class=tablebody1>
<%=Rs("Username")%></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>�ֻ�����</B>��</TD>
<TD class=tablebody1>
<%=rs("usermobile")%>
</TD>
</TR>
<tr><td align=center class=tablebody2 colspan=2><input type=submit value="�� ��"></td></form></tr>
</tbody>
</table>
</form>
<%

set rs=nothing
end sub

sub reg_2()
'if request("mobile")="" then
'	founderr=true
'	Dvbbs.AddErrmsg "�����������ֻ��š�"
'	exit sub
'end if
dim rs
set rs=Dvbbs.Execute("select * from [dv_user] where userid="&dvbbs.userid)
if rs("IsChallenge")=0 or isnull(rs("IsChallenge")) then Response.redirect "showerr.asp?ErrCodes=<li>�������������Ա����������Ϊ�����Ա��&action=OtherErr"
dim mobile
mobile=rs("UserMobile")


%>
�����ύ���ݣ����Ժ󡭡�
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
	w=replace(v,"|||","����")
	checkreal=w
end if
end function
%>