<!--#include file="Conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.LoadTemplates("")
Dvbbs.Stats="��װ��̳"
Dvbbs.Nav()
Dvbbs.Showerr()
Dim Rs
If Not Dvbbs.Master Then
	Dvbbs.AddErrCode(36)
End If
Dvbbs.Showerr()
If Session("flag")="" Then
	Dvbbs.AddErrCode(36)
End If
Dvbbs.Showerr()
Set Rs=Dvbbs.Execute("select * from dv_Setup")
If rs("Forum_Isinstall")=1 Then
	If Rs("Forum_version")="7.0.0" Then
		If dvbbs.master and session("flag")<>"" Then
			If request("isnew")="1" Or request("action")<>""  Then
				dvbbs.execute("update Dv_Setup set Forum_ChallengePassword='raynetwork',Forum_isinstall=0,Forum_ChanSetting='1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1'")
				Dvbbs.Name="setup"
				Dvbbs.ReloadSetup
			Else
				Dvbbs.AddErrCode(36)
			End If
		Else
			Dvbbs.AddErrCode(36)
		End If
	Else
		Dvbbs.AddErrCode(37)
	End If
Else
	If request("action") <> "" Or Request("isnew")<>"" Then
		dvbbs.execute("update dv_Setup set Forum_ChallengePassword='raynetwork',Forum_isinstall=0,Forum_ChanSetting='1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1'")
		Dvbbs.Name="setup"
		Dvbbs.ReloadSetup
	Else
		Dvbbs.AddErrCode(36)
	End If
End If
Dvbbs.Showerr()

Select Case request("action")
	Case "apply"
		If request("isnew")="0" Then
			dvbbs.stats="��д����"
			Dvbbs.Head_var 0,0,"��װ��̳","install.asp"
			reg_2()
		Else
			dvbbs.stats="ȷ������"
			Dvbbs.Head_var 0,0,"��װ��̳","install.asp"
			reg_2a()
		End If
	Case "redir1"
		dvbbs.stats="�ύע��"
		Dvbbs.Head_var 0,0,"��װ��̳","install.asp"
		Call redir1()
	Case "redir2"
		dvbbs.stats="�ύע��"
		Dvbbs.Head_var 0,0,"��װ��̳","install.asp"
		call redir2()
	Case "redir3"
		dvbbs.stats="�ύע��"
		Dvbbs.Head_var 0,0,"��װ��̳","install.asp"
		call redir3()
	Case "redir4"
		dvbbs.stats="�ύע��"
		Dvbbs.Head_var 0,0,"��װ��̳","install.asp"
		call redir4()
	Case Else
		dvbbs.stats="ע��Э��"
		Dvbbs.Head_var 0,0,"��װ��̳","install.asp"
		reg_1
End Select

Dvbbs.ActiveOnline()
Dvbbs.Footer


Function reg_1()
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function submitclick()
{
	if(dvform.isnew[0].checked==false&&dvform.isnew[1].checked==false)
	{
		alert('ע��ע�����\n1������ϸ�Ķ�ע��Э�飬������ѡ����վ��ע�ᡱ����ע�ᣬ����ȷ�����ϡ���\n2�����������վ����û��ע���������ע��������ѡ����վ��ע�ᡱ\n3��������Ѿ�ע���������̳վ����ֻ�Ǹı��˵�ַ�������°�װ��������ע��������ѡ����ע�ᣬ����ȷ�����ϡ� ')
	}
	else
	{
		dvform.submit()
	}
}
//-->
</SCRIPT>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1><form name="dvform"  action="install.asp?action=apply" method=post>
<tr><th align=center colSpan=2>�������������</td></tr>
<tr><td class=tablebody1 align=left colSpan=2>
<b>������װ������̳ǰ�����Ķ������̳Э��</b>
<BR><BR>
������̳��������̳ϵ�����Ϊ���е���̳վ���ṩ�˸��ַḻ�Ļ�������ͬʱ�ɴ˴�������ط������棬����������ź�վ������<BR><BR>
Խ����й��ƶ�ȫ��ͨ�û�ʹ�����Ƽ���������Ų�Ʒ������õĻر���Խ�ࡣ����������Ƽ����û����͵Ķ��ų���2000Ԫ�����㲢���Գ�Ϊ���µĳ�����̳������30�������������<BR><BR>
<B>�������</B>   <BR><BR>
����ʹ��"������̳��������̳ϵ�����"����ʼ�������Ϊ25����<BR>
��������·��͵Ķ��Ŵﵽ2000Ԫ������ڵ��°��ճ�����̳���н��㣬����30%�ĸ߱������档  <BR><BR>
<B>ʹ������</B>   <BR><BR>
����̳����ʱ����������дվ��ע��������ɳ�Ϊ"������̳��������̳ϵ�����"���û������û�������Ϣ�����д���������ᵼ�»�������ȷͶ�ݣ����߲���ҵ���޷�ʹ�á���  <BR><BR>
��Ա��¼������̳�Ĺ������ġ��ڹ������ģ����Խ�����̳�����ѯ���û����ϡ������޸ĵȲ�����  <BR><BR>
ȫ��ͨ�û���ʹ����ҳ���ϵ�������ŷ���ʱ��ÿ�ɹ�����һ�ζ��ţ����Ϳ��Դӷ��ͽ���л�ø߱������档  <BR><BR>
<B>�������</B><BR><BR>
������̳��������̳����Ʒ���������Ȼ��Ϊ�Ʒ����ڡ� <BR>
����ʱ��������Ļ�Ա�����������֧�������޶100Ԫ����100Ԫ�������ǻ���ÿ����20������ͨ���ʾֻ��ķ�ʽ֧�������� <BR><BR>
����ʱ��������Ļ�Ա�������δ�ﵽ100Ԫ�����Զ��ۼƵ��¸��£�ֱ������ۼƴﵽ���֧�������޶�Ϊֹ <BR><BR>
ÿ��֧�������޷ⶥ���ޣ��ʾֻ����Ҫ���ȿ۳����ʼ����˰��  <BR><BR>
����������㽫��ÿ�����ƶ���Ӫ�̺˶��������Ϊ׼��  <BR>
�緢��������Ϊ��ֹͣ���Ѳ�ȡ�����Ա�ʸ�ͬʱ������һ��׷���������ε�Ȩ����  <BR><BR>
���Ծ����ء�ȫ���˴�ί�����ά����������ȫ�ľ��������л����񹲺͹���������ɷ��棬��ֹ�κξ��ھ���ɫ��򷴶���վʹ��"������̳��������̳ϵ�����"��һ�����֣�����������з����ϵ���۷��������棬���Ҹû�Ա���е��ɴ˲�����һ�к����  <BR><BR>
��ʹ��"������̳��������̳ϵ�����"��Ϊע������û��� ����ʾ���Ѿ��Ķ������������������<BR><BR>
</td></tr>
<TR align=middle>
<Th colSpan=2 height=24>��ѡ��ע������</TD>
</TR>
<tr>
<td  width=40% align=center  class=tablebody1  height=24> <input type="radio" name="isnew" value="0"><b>��վ��ע��</b></td>
<td width=60% align=center  class=tablebody1  height=24> <input type="radio" name="isnew" value="1" ><b>��ע�ᣬ����ȷ������</b></td>
</TR>
<TR align=middle>
<td align=center class=tablebody2 colSpan=2><input type="button" value="��ͬ��" Onclick="submitclick();"></td></tr>
</form></table>
<%
End Function

Function reg_3()
	Response.Write "<script>dvbbs_install_reg_3();</script>"
End Function

Function reg_2()
%>
<FORM name=theForm action=install.asp?action=redir1 method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TR align=middle>
<Th colSpan=2 height=24>��վ��������д</TD>
</TR>
<TR>
<TR>
<Td colSpan=2 class=tablebody1>˵�����ڴ˰�װ�����У���Ĭ����ͬ�������̳���Ż��������������Ͻ��ύ��������������ȷ�ϣ�����дǰ��ȷ������д����������Ƿ���ȷ���⽫Ӱ�쵽������ڷ����о�����������֣�������ɴ˲��������ܼ�����װ��̳��</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*�û���</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="username"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*��ʵ����</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="realname"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*���֤��</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="identityNo"></TD>
</TR>
<TR>
<TD width=40%  class=tablebody1><B>*�Ա�</B>��<BR>��ѡ�������Ա�</font></TD>
<TD width=60%  class=tablebody1> <INPUT type=radio CHECKED value="F" name=sex>
�� &nbsp;&nbsp;&nbsp;&nbsp;
<INPUT type=radio value="M" name=sex>
Ů &nbsp;&nbsp;&nbsp;&nbsp;
<INPUT type=radio value="N" name=sex>
����</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*�ʱ�</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="postcode"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*��ַ</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="address"></TD>
</TR>
<INPUT type=hidden name="receiver"><TR>
<TD width=40% class=tablebody1><B>*�ʼ���ַ</B>��<br>Ϊ���ܳɹ���װ��̳���������д��ʵ��Ч��Email��ַ</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="email"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>�绰</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="telephone"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>�ֻ�</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="mobile"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>����ȫ��</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="bankname"></TD>
</TR>
<TD width=40% class=tablebody1><B>�����ʺ�</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="bankid"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*��̳����</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="forumname"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*��̳��ַ</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=hidden value="<%=Dvbbs.Get_ScriptNameUrl%>" name="forumUrl">
<%=Dvbbs.Get_ScriptNameUrl%></TD>
</TR>
<TR align=middle>
<Th colSpan=2 height=24>��̳������д</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*��̳�ṩ��</B>��</TD>
<TD width=60%  class=tablebody1>
�����ȷ�</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*��̳�汾</B>��</TD>
<TD width=60%  class=tablebody1>
Dvbbs 7.0.0</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>ICP��</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="icp"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>��վ���ݽ���</B>��</TD>
<TD width=60%  class=tablebody1>
<textarea class=smallarea cols=100 name=web_intro rows=7 wrap=VIRTUAL>
</textarea>
</TD>
</TR>
</tbody>
</table>
</td></tr></tbody></table>
<table cellpadding=0 cellspacing=0 border=0 width='+Dvbbs.Forum_Body[12]+' align=center>
<tr>
<td width=50% height=24> </td>
<td width=50% ><input type=submit value="ע ��" name=Submit>����<input type=reset value="�� ��" name=Submit2></td>
</tr></table>
</form>
<%
End Function

Function reg_2a()
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0

  window.open(theURL,winName,features);

}
//-->
</SCRIPT>
<FORM name=theForm action=install.asp?action=redir2 method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TR align=middle>
<Th colSpan=2 height=24>ԭע��վ����֤������д</TD>
</TR>
<TR>
<TR>
<Td colSpan=2 class=tablebody1>˵����������Ѿ�ע������Ż���������ֱ���ڴ���д������֤��Ϣ�ύ����������������֤����֤ͨ������Լ�����װ��̳��</B></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>�û���</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="username"> ��ע���ʱ��ʹ�õ��û���</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>����</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=password size=30 name="password"> <a  style="CURSOR:hand" onclick="MM_openBrWindow('http://bbs.ray5198.com/forgetpass.html','getpass','width=300,height=100')">��������</a></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>��̳����</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="forumname"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>��̳��ַ</B>��</TD>
<TD width=60%  class=tablebody1>
<INPUT type=hidden value="<%=Dvbbs.Get_ScriptNameUrl%>" name="forumUrl">
<%=Dvbbs.Get_ScriptNameUrl%></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>��̳�ṩ��</B>��</TD>
<TD width=60%  class=tablebody1>
�����ȷ�</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>��̳�汾</B>��</TD>
<TD width=60%  class=tablebody1>
Dvbbs 7.0.0
</TD>
</TR>
</tbody>
</table>
</td></tr></tbody></table>
<table cellpadding=0 cellspacing=0 border=0 width='+Dvbbs.Forum_Body[12]+' align=center>
<tr>
<td width=50% height=24> </td>
<td width=50% ><input type=submit value="ע ��" name=Submit>����<input type=reset value="�� ��" name=Submit2></td>
</tr></table>
</form>
<%
End Function

Function redir1()
	If request("username")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�����������û�����"
		Exit Function
	End If
	If request("realname")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>������������ʵ������"
		Exit Function
	End If
	If request("identityNo")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�������������֤�š�"
		Exit Function
	End If
	If request("sex")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��ѡ�������Ա�"
		Exit Function
	End If
	If request("postcode")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�����������������롣"
		Exit Function
	End If
	If request("address")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>���������ĵ�ַ��"
		Exit Function
	End If
	If request("email")="" Then 
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�����������ʼ���ַ��"
		Exit Function
	End If
	If request("forumname")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>������������̳���ơ�"
		Exit Function
	End If
	session("Forum_Master_Reg_Temp_1")=checkreal(request("username")) & "|||" & checkreal(request("realname")) & "|||" & checkreal(request("identityNo")) & "|||" & checkreal(request("sex")) & "|||" & checkreal(request("postcode")) & "|||" & checkreal(request("address")) & "|||" & checkreal(request("receiver")) & "|||" & checkreal(request("email")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||" & checkreal(request("telephone")) & "|||" & checkreal(request("mobile")) & "|||�����ȷ�|||7.0.0"

	Get_ChallengeWord

%>
�����ύ���ݣ����Ժ󡭡�
<form name="redir" action="http://bbs.ray5198.com/registerAdminAndForum.jsp" method="post">
<INPUT type=hidden name="username" value="<%=checkreal(request("username"))%>">
<INPUT type=hidden name="realname" value="<%=checkreal(request("realname"))%>">
<INPUT type=hidden name="identityNo" value="<%=checkreal(request("identityNo"))%>">
<INPUT type=hidden name="sex" value="<%=checkreal(request("sex"))%>">
<INPUT type=hidden name="postcode" value="<%=checkreal(request("postcode"))%>">
<INPUT type=hidden name="address" value="<%=checkreal(request("address"))%>">
<INPUT type=hidden name="receiver" value="<%=checkreal(request("realname"))%>">
<INPUT type=hidden name="email" value="<%=checkreal(request("email"))%>">
<INPUT type=hidden name="forumName" value="<%=checkreal(request("forumname"))%>">
<INPUT type=hidden name="forumUrl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<INPUT type=hidden name="telephone" value="<%=checkreal(request("telephone"))%>">
<INPUT type=hidden name="mobile" value="<%=checkreal(request("mobile"))%>">
<INPUT type=hidden name="bankname" value="<%=checkreal(request("bankname"))%>">
<INPUT type=hidden name="bankid" value="<%=checkreal(request("bankid"))%>">
<INPUT type=hidden name="forumProvider" value="�����ȷ�">
<INPUT type=hidden name="version" value="Dvbbs 7.0.0">
<INPUT type=hidden name="icpNo" value="<%=checkreal(request("icpNO"))%>">
<INPUT type=hidden name="web_intro" value="<%=checkreal(request("web_intro"))%>">
<input type=hidden value="<%=Session("challengeWord")%>" name="challengeWord">
<input type=hidden value="install.asp?action=redir3" name="dirPage">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
End Function

Function redir2()

	If request("username")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�����������û�����"
		Exit Function
	End If
	If request("password")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�������������롣"
		Exit Function
	End If
	If request("forumname")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>������������̳���ơ�"
		Exit Function
	End If
	session("Forum_Master_Reg_Temp_2")=checkreal(request("username")) & "|||" & checkreal(request("password")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||�����ȷ�|||7.0.0"

	Get_ChallengeWord

	session("challengeWord_key")=md5(Session("challengeWord") & ":" & Dvbbs.CacheData(21,0),32)
%>
�����ύ���ݣ����Ժ󡭡�
<form name="redir" action="http://bbs.ray5198.com/popRegAdmin.jsp" method="post">
<INPUT type=hidden name="username" value="<%=checkreal(request("username"))%>">
<INPUT type=hidden name="password" value="<%=checkreal(request("password"))%>">
<INPUT type=hidden name="forumName" value="<%=checkreal(request("forumname"))%>">
<INPUT type=hidden name="forumUrl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<INPUT type=hidden name="forumProvider" value="�����ȷ�">
<INPUT type=hidden name="version" value="Dvbbs 6.1.0">
<input type=hidden value="<%=Session("challengeWord")%>" name="challengeWord">
<input type=hidden value="install.asp?action=redir4" name="dirPage">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
End Function

Function redir3()

	Dim ErrorCode,ErrorMsg
	Dim reForumID,reNewPkey,rechallengeWord,retokerWord
	Dim challengeWord_key,rechallengeWord_key
	Dim Forum_Master_Reg_Temp_1,OldForumID
	Dim vipboardsetting,vipboardslist,vipisupdate,i
	vipisupdate=false
	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))
	reForumID=trim(Dvbbs.CheckStr(request("ForumID")))
	reNewPkey=trim(Dvbbs.CheckStr(request("NewPkey")))
	rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
	retokerWord=trim(request("tokenWord"))
	'rechallengeWord_key=md5_32(rechallengeWord & ":" & Dvbbs.CacheData(21,0))

	Select Case ErrorCode
	Case 100
		challengeWord_key=session("challengeWord_key")							
		If challengeWord_key=retokerWord Then
			Forum_Master_Reg_Temp_1=split(Dvbbs.CheckStr(session("Forum_Master_Reg_Temp_1")),"|||")
			Dvbbs.Execute("update dv_Setup set Forum_challengePassWord='"&reNewPkey&"',Forum_ChanName='"&Forum_Master_Reg_Temp_1(0)&"',Forum_IsInstall=1,Forum_Version='"&Forum_Master_Reg_Temp_1(13)&"'")
			Set Rs=Dvbbs.Execute("select top 1 * from Dv_ChallengeInfo")
			OldForumID=rs("D_ForumID")
			Rs.close
			Set Rs=Nothing
			Dvbbs.Execute("update Dv_ChallengeInfo set D_ForumID='"&reForumID&"',D_UserName='"&Forum_Master_Reg_Temp_1(0)&"',D_Password='"&reNewPkey&"',D_RealName='"&Forum_Master_Reg_Temp_1(1)&"',D_identityNo='"&Forum_Master_Reg_Temp_1(2)&"',D_sex='"&Forum_Master_Reg_Temp_1(3)&"',D_postcode='"&Forum_Master_Reg_Temp_1(4)&"',D_address='"&Forum_Master_Reg_Temp_1(5)&"',D_receiver='"&Forum_Master_Reg_Temp_1(6)&"',D_email='"&Forum_Master_Reg_Temp_1(7)&"',D_forumname='"&Forum_Master_Reg_Temp_1(8)&"',D_forumurl='"&Forum_Master_Reg_Temp_1(9)&"',D_telephone='"&Forum_Master_Reg_Temp_1(10)&"',D_mobile='"&Forum_Master_Reg_Temp_1(11)&"',D_forumProvider='"&Forum_Master_Reg_Temp_1(12)&"',D_version='"&Forum_Master_Reg_Temp_1(13)&"',D_challengePassWord='"&OldForumID&"'")
			Set Rs=Dvbbs.Execute("Select BoardID,Board_Setting From Dv_Board")
			Do While Not Rs.Eof
				vipboardsetting=split(rs("Board_Setting"),",")
				For i=0 to UBound(vipboardsetting)
					If i=0 Then
						vipboardslist=vipboardsetting(i)
					Else
						'��֤��̳��ԭ
						If i=2 and vipboardsetting(2)=1 and vipboardsetting(46)>0 Then
							vipboardslist=vipboardslist & ",0"
							vipisupdate=true
						Else
							vipboardslist=vipboardslist & "," & vipboardsetting(i)
						End If
					End If
				Next
				if vipisupdate then Dvbbs.Execute("update dv_board set board_setting='"&vipboardslist&"' where boardid="&rs(0))
				vipisupdate=false
			Rs.Movenext
			Loop
			rs.close
			Set rs=nothing
			Dvbbs.Name="setup"
			Dvbbs.ReloadSetup
		Else
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�Ƿ��Ĳ�����"
			Exit Function
		End If
	Case 102
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��ע����û����Ͷ��Ż����������������ϵ��û����ظ���"&ErrorMsg&"��"
		Exit Function
	Case 201
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>����д����Ϣ�ڶ��Ż����������������ϵ�¼��֤ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case Else
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�Ƿ����ύ���̣�"&ErrorMsg&"��"
		Exit Function
	End Select
	Emp_ChallengeWord
	session("Forum_Master_Reg_Temp_1")=""
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>ע��ɹ���������̳���Ż�������ע��ɹ�</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li><B>��ɾ����̳��install.asp�ļ������������̳</B>��</li><li>��̳Ĭ�Ϲ���Ա�ʺ��ǣ��û���admin������admin888��ǰ��̨һ��</li><li><a href="index.asp">����������̳</a></li></ul>
</td></tr>
</table>
<%
End Function

Function redir4()

	Dim ErrorCode,ErrorMsg
	Dim reForumID,reNewPkey,rechallengeWord,retokerWord
	Dim challengeWord_key,rechallengeWord_key
	Dim Forum_Master_Reg_Temp_2,OldForumID,i
	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))
	reForumID=trim(Dvbbs.CheckStr(request("ForumID")))
	reNewPkey=trim(Dvbbs.CheckStr(request("NewPkey")))
	rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
	retokerWord=trim(request("tokenWord"))
	'rechallengeWord_key=md5_32(rechallengeWord & ":" & Dvbbs.CacheData(21,0))
	Dim vipboardsetting,vipboardslist,vipisupdate
	vipisupdate=false

	Select Case ErrorCode
	Case 100
		challengeWord_key=session("challengeWord_key")
		If challengeWord_key=retokerWord Then
			Forum_Master_Reg_Temp_2=split(Dvbbs.CheckStr(session("Forum_Master_Reg_Temp_2")),"|||")
			Dvbbs.Execute("update dv_Setup set Forum_challengePassWord='"&reNewPkey&"',Forum_ChanName='"&Forum_Master_Reg_Temp_2(0)&"',Forum_IsInstall=1,Forum_Version='"&Forum_Master_Reg_Temp_2(5)&"'")
			Set Rs=Dvbbs.Execute("select top 1 * from Dv_ChallengeInfo")
			OldForumID=rs("D_ForumID")
			Rs.close
			Set Rs=Nothing
			Dvbbs.Execute("update Dv_ChallengeInfo set D_ForumID='"&reForumID&"',D_UserName='"&Forum_Master_Reg_Temp_2(0)&"',D_Password='"&reNewPkey&"',D_forumname='"&Forum_Master_Reg_Temp_2(2)&"',D_forumurl='"&Forum_Master_Reg_Temp_2(3)&"',D_forumProvider='"&Forum_Master_Reg_Temp_2(4)&"',D_version='"&Forum_Master_Reg_Temp_2(5)&"',D_challengePassWord='"&OldForumID&"'")
			Set Rs=Dvbbs.Execute("Select BoardID,Board_Setting From dv_Board")
			Do While Not Rs.Eof
				vipboardsetting=split(rs("Board_Setting"),",")
				For i=0 to UBound(vipboardsetting)
					If i=0 Then
						vipboardslist=vipboardsetting(i)
					Else
						'��֤��̳��ԭ
						If i=2 and vipboardsetting(2)=1 and vipboardsetting(46)>0 Then
							vipboardslist=vipboardslist & ",0"
							vipisupdate=true
						Else
							vipboardslist=vipboardslist & "," & vipboardsetting(i)
						End If
					End If
				Next
				if vipisupdate then Dvbbs.Execute("update dv_board set board_setting='"&vipboardslist&"' where boardid="&rs(0))
				vipisupdate=false
			Rs.Movenext
			Loop
			rs.close
			Set rs=nothing
			Dvbbs.Name="setup"
			Dvbbs.ReloadSetup
		Else
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�Ƿ��Ĳ�����"
			Exit Function
		End If
	Case 101
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>���ڶ��Ż���������������ע��ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case 102
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��ע����û����Ͷ��Ż����������������ϵ��û����ظ���"&ErrorMsg&"��"
		Exit Function
	Case 201
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>����д����Ϣ�ڶ��Ż����������������ϵ�¼��֤ʧ�ܣ�"&ErrorMsg&"��"
		Exit Function
	Case Else
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�Ƿ����ύ���̣�"&ErrorMsg&"��"
		Exit Function
	End Select
	Emp_ChallengeWord
	session("Forum_Master_Reg_Temp_2")=""
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>�ɹ���Ϣ��������̳���Ż���������Ϣ���³ɹ�</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li><a href="index.asp">����������̳</a></li></ul>
</td></tr>
</table>
<%
End Function

function checkreal(v)
Dim w
if not isnull(v) then
	w=replace(v,"|||","����")
	checkreal=w
end if
end function
%>