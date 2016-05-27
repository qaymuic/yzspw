<!--#include file="Conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.LoadTemplates("")
Dvbbs.Stats="安装论坛"
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
			dvbbs.stats="填写资料"
			Dvbbs.Head_var 0,0,"安装论坛","install.asp"
			reg_2()
		Else
			dvbbs.stats="确认资料"
			Dvbbs.Head_var 0,0,"安装论坛","install.asp"
			reg_2a()
		End If
	Case "redir1"
		dvbbs.stats="提交注册"
		Dvbbs.Head_var 0,0,"安装论坛","install.asp"
		Call redir1()
	Case "redir2"
		dvbbs.stats="提交注册"
		Dvbbs.Head_var 0,0,"安装论坛","install.asp"
		call redir2()
	Case "redir3"
		dvbbs.stats="提交注册"
		Dvbbs.Head_var 0,0,"安装论坛","install.asp"
		call redir3()
	Case "redir4"
		dvbbs.stats="提交注册"
		Dvbbs.Head_var 0,0,"安装论坛","install.asp"
		call redir4()
	Case Else
		dvbbs.stats="注册协议"
		Dvbbs.Head_var 0,0,"安装论坛","install.asp"
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
		alert('注册注意事项：\n1、请仔细阅读注册协议，您必须选择“新站长注册”或“已注册，重新确认资料”。\n2、如果您的网站从来没有注册过，请在注册类型中选择“新站长注册”\n3、如果您已经注册过阳光论坛站长，只是改变了地址或者重新安装，可以在注册类型中选择“已注册，重新确认资料” ')
	}
	else
	{
		dvform.submit()
	}
}
//-->
</SCRIPT>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1><form name="dvform"  action="install.asp?action=apply" method=post>
<tr><th align=center colSpan=2>服务条款和声明</td></tr>
<tr><td class=tablebody1 align=left colSpan=2>
<b>继续安装动网论坛前请先阅读相关论坛协议</b>
<BR><BR>
动网论坛和阳光论坛系列软件为所有的论坛站长提供了各种丰富的互动服务，同时由此带来的相关服务收益，将由阳光加信和站长分享。<BR><BR>
越多的中国移动全球通用户使用您推荐的阳光短信产品，您获得的回报便越多。如果本月您推荐的用户发送的短信超过2000元，您便并可以成为本月的超级论坛，享受30％的收益比例。<BR><BR>
<B>收益比例</B>   <BR><BR>
申请使用"动网论坛和阳光论坛系列软件"，初始收益比例为25％。<BR>
如果您当月发送的短信达到2000元，便可在当月按照超级论坛进行结算，享受30%的高比例收益。  <BR><BR>
<B>使用流程</B>   <BR><BR>
在论坛建立时，请认真填写站长注册表单，即可成为"动网论坛和阳光论坛系列软件"的用户。（用户个人信息务必填写清楚，否则会导致汇款单不能正确投递，或者部分业务无法使用。）  <BR><BR>
会员登录到本论坛的管理中心。在管理中心，可以进行论坛收益查询，用户资料、密码修改等操作。  <BR><BR>
全球通用户在使用您页面上的阳光短信服务时，每成功发送一次短信，您就可以从发送金额中获得高比例收益。  <BR><BR>
<B>收益结算</B><BR><BR>
动网论坛和阳光论坛收益计费周期以自然月为计费周期。 <BR>
结算时，如果您的会员帐面余额超过最低支付收益限额：100元（含100元），我们会在每个月20日左右通过邮局汇款的方式支付给您。 <BR><BR>
结算时，如果您的会员帐面余额未达到100元，则自动累计到下个月，直至余额累计达到最低支付收益限额为止 <BR><BR>
每月支付收益无封顶上限，邮局汇款需要事先扣除邮资及相关税金。  <BR><BR>
最终收益结算将以每月与移动运营商核对相关数据为准。  <BR>
如发现作弊行为将停止付费并取消其会员资格，同时保留进一步追究法律责任的权利。  <BR><BR>
请自觉遵守《全国人大常委会关于维护互联网安全的决定》及中华人民共和国其他各项法律法规，禁止任何境内境外色情或反动网站使用"动网论坛和阳光论坛系列软件"，一经发现，立即解除所有服务关系，扣发所有收益，并且该会员将承担由此产生的一切后果。  <BR><BR>
您使用"动网论坛和阳光论坛系列软件"成为注册软件用户， 即表示您已经阅读并接受如上所有条款。<BR><BR>
</td></tr>
<TR align=middle>
<Th colSpan=2 height=24>请选择注册类型</TD>
</TR>
<tr>
<td  width=40% align=center  class=tablebody1  height=24> <input type="radio" name="isnew" value="0"><b>新站长注册</b></td>
<td width=60% align=center  class=tablebody1  height=24> <input type="radio" name="isnew" value="1" ><b>已注册，重新确认资料</b></td>
</TR>
<TR align=middle>
<td align=center class=tablebody2 colSpan=2><input type="button" value="我同意" Onclick="submitclick();"></td></tr>
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
<Th colSpan=2 height=24>新站长资料填写</TD>
</TR>
<TR>
<TR>
<Td colSpan=2 class=tablebody1>说明：在此安装过程中，将默认您同意加入论坛短信互动服务，所有资料将提交至主服务器进行确认，在填写前请确认您填写的相关资料是否正确，这将影响到您今后在服务中具体利益的体现，必须完成此步操作才能继续安装论坛。</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*用户名</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="username"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*真实姓名</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="realname"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*身份证号</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="identityNo"></TD>
</TR>
<TR>
<TD width=40%  class=tablebody1><B>*性别</B>：<BR>请选择您的性别</font></TD>
<TD width=60%  class=tablebody1> <INPUT type=radio CHECKED value="F" name=sex>
男 &nbsp;&nbsp;&nbsp;&nbsp;
<INPUT type=radio value="M" name=sex>
女 &nbsp;&nbsp;&nbsp;&nbsp;
<INPUT type=radio value="N" name=sex>
保密</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*邮编</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="postcode"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*地址</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="address"></TD>
</TR>
<INPUT type=hidden name="receiver"><TR>
<TD width=40% class=tablebody1><B>*邮件地址</B>：<br>为了能成功安装论坛，请务必填写真实有效的Email地址</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="email"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>电话</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="telephone"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>手机</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="mobile"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>银行全称</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="bankname"></TD>
</TR>
<TD width=40% class=tablebody1><B>银行帐号</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="bankid"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*论坛名称</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="forumname"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*论坛地址</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=hidden value="<%=Dvbbs.Get_ScriptNameUrl%>" name="forumUrl">
<%=Dvbbs.Get_ScriptNameUrl%></TD>
</TR>
<TR align=middle>
<Th colSpan=2 height=24>论坛资料填写</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*论坛提供者</B>：</TD>
<TD width=60%  class=tablebody1>
动网先锋</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>*论坛版本</B>：</TD>
<TD width=60%  class=tablebody1>
Dvbbs 7.0.0</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>ICP号</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="icp"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>网站内容介绍</B>：</TD>
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
<td width=50% ><input type=submit value="注 册" name=Submit><input type=reset value="清 除" name=Submit2></td>
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
<Th colSpan=2 height=24>原注册站长验证资料填写</TD>
</TR>
<TR>
<TR>
<Td colSpan=2 class=tablebody1>说明：如果您已经注册过短信互动服务，请直接在此填写您的认证信息提交至主服务器进行验证，验证通过后可以继续安装论坛。</B></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>用户名</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="username"> 您注册的时候使用的用户名</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>密码</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=password size=30 name="password"> <a  style="CURSOR:hand" onclick="MM_openBrWindow('http://bbs.ray5198.com/forgetpass.html','getpass','width=300,height=100')">忘记密码</a></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>论坛名称</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=text size=30 name="forumname"></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>论坛地址</B>：</TD>
<TD width=60%  class=tablebody1>
<INPUT type=hidden value="<%=Dvbbs.Get_ScriptNameUrl%>" name="forumUrl">
<%=Dvbbs.Get_ScriptNameUrl%></TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>论坛提供者</B>：</TD>
<TD width=60%  class=tablebody1>
动网先锋</TD>
</TR>
<TR>
<TD width=40% class=tablebody1><B>论坛版本</B>：</TD>
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
<td width=50% ><input type=submit value="注 册" name=Submit><input type=reset value="清 除" name=Submit2></td>
</tr></table>
</form>
<%
End Function

Function redir1()
	If request("username")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的用户名。"
		Exit Function
	End If
	If request("realname")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的真实姓名。"
		Exit Function
	End If
	If request("identityNo")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的身份证号。"
		Exit Function
	End If
	If request("sex")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请选择您的性别。"
		Exit Function
	End If
	If request("postcode")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的邮政编码。"
		Exit Function
	End If
	If request("address")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的地址。"
		Exit Function
	End If
	If request("email")="" Then 
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的邮件地址。"
		Exit Function
	End If
	If request("forumname")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的论坛名称。"
		Exit Function
	End If
	session("Forum_Master_Reg_Temp_1")=checkreal(request("username")) & "|||" & checkreal(request("realname")) & "|||" & checkreal(request("identityNo")) & "|||" & checkreal(request("sex")) & "|||" & checkreal(request("postcode")) & "|||" & checkreal(request("address")) & "|||" & checkreal(request("receiver")) & "|||" & checkreal(request("email")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||" & checkreal(request("telephone")) & "|||" & checkreal(request("mobile")) & "|||动网先锋|||7.0.0"

	Get_ChallengeWord

%>
正在提交数据，请稍后……
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
<INPUT type=hidden name="forumProvider" value="动网先锋">
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
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的用户名。"
		Exit Function
	End If
	If request("password")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的密码。"
		Exit Function
	End If
	If request("forumname")="" Then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的论坛名称。"
		Exit Function
	End If
	session("Forum_Master_Reg_Temp_2")=checkreal(request("username")) & "|||" & checkreal(request("password")) & "|||" & checkreal(request("forumname")) & "|||" & checkreal(request("forumurl")) & "|||动网先锋|||7.0.0"

	Get_ChallengeWord

	session("challengeWord_key")=md5(Session("challengeWord") & ":" & Dvbbs.CacheData(21,0),32)
%>
正在提交数据，请稍后……
<form name="redir" action="http://bbs.ray5198.com/popRegAdmin.jsp" method="post">
<INPUT type=hidden name="username" value="<%=checkreal(request("username"))%>">
<INPUT type=hidden name="password" value="<%=checkreal(request("password"))%>">
<INPUT type=hidden name="forumName" value="<%=checkreal(request("forumname"))%>">
<INPUT type=hidden name="forumUrl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<INPUT type=hidden name="forumProvider" value="动网先锋">
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
						'认证论坛还原
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
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>非法的参数。"
			Exit Function
		End If
	Case 102
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您注册的用户名和短信互动服务主服务器上的用户名重复；"&ErrorMsg&"。"
		Exit Function
	Case 201
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您填写的信息在短信互动服务主服务器上登录验证失败；"&ErrorMsg&"。"
		Exit Function
	Case Else
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>非法的提交过程；"&ErrorMsg&"。"
		Exit Function
	End Select
	Emp_ChallengeWord
	session("Forum_Master_Reg_Temp_1")=""
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>注册成功：您在论坛短信互动服务注册成功</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li><B>请删除论坛中install.asp文件后进入您的论坛</B>。</li><li>论坛默认管理员帐号是：用户名admin，密码admin888，前后台一样</li><li><a href="index.asp">进入您的论坛</a></li></ul>
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
						'认证论坛还原
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
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>非法的参数。"
			Exit Function
		End If
	Case 101
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您在短信互动服务主服务器注册失败，"&ErrorMsg&"。"
		Exit Function
	Case 102
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您注册的用户名和短信互动服务主服务器上的用户名重复，"&ErrorMsg&"。"
		Exit Function
	Case 201
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您填写的信息在短信互动服务主服务器上登录验证失败，"&ErrorMsg&"。"
		Exit Function
	Case Else
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>非法的提交过程，"&ErrorMsg&"。"
		Exit Function
	End Select
	Emp_ChallengeWord
	session("Forum_Master_Reg_Temp_2")=""
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>成功信息：您在论坛短信互动服务信息更新成功</th>
</tr>
<tr><td class=tablebody1><br>
<ul><li><a href="index.asp">进入您的论坛</a></li></ul>
</td></tr>
</table>
<%
End Function

function checkreal(v)
Dim w
if not isnull(v) then
	w=replace(v,"|||","§§§")
	checkreal=w
end if
end function
%>