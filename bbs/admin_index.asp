<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/md5.asp" -->
<!-- #include file="inc/myadmin.asp" -->
<!-- #include file="inc/dvadchar.asp" -->
<%
dim username
dim password
dim ip
dvbbs.stats="论坛管理控制面板"
select case request("action")
case "admin_left"
	call admin_left()
case "admin_login"
	call admin_login()
case "admin_main"
	call admin_main()
case "admin_head"
	call admin_head()
case Else
	call main()
end Select

sub main()
if not dvbbs.master or session("flag")="" then
	call admin_login()
else
%>
<html>
<head>
<title><%=dvbbs.Forum_info(0)%>--控制面板</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<frameset id="frame" cols="180,*" frameborder="NO" border="0" framespacing="0" rows="*"> 
  <frame name="leftFrame" scrolling="AUTO" noresize src="admin_index.asp?action=admin_left" marginwidth="0" marginheight="0">
<%if not dvbbs.master or session("flag")="" then%>
  <frame name="main" src="admin_index.asp?action=admin_login" scrolling="AUTO" NORESIZE frameborder="0" marginwidth="10" marginheight="10" border="no">
<%else%>
  <frame name="main" src="admin_index.asp?action=admin_main" scrolling="AUTO" NORESIZE frameborder="0" marginwidth="10" marginheight="10" border="no">
<%end if%>
</frameset>
</frameset>
<noframes>

</body></noframes>	
</html>
<%
end if
end sub

sub admin_left()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<title><%=dvbbs.Forum_info(0)%>--管理页面</title>
<style type=text/css>
body  { background:#799AE1; margin:0px; font:normal 12px 宋体; 
SCROLLBAR-FACE-COLOR: #799AE1; SCROLLBAR-HIGHLIGHT-COLOR: #799AE1; 
SCROLLBAR-SHADOW-COLOR: #799AE1; SCROLLBAR-DARKSHADOW-COLOR: #799AE1; 
SCROLLBAR-3DLIGHT-COLOR: #799AE1; SCROLLBAR-ARROW-COLOR: #FFFFFF;
SCROLLBAR-TRACK-COLOR: #AABFEC;
}
table  { border:0px; }
td  { font:normal 12px 宋体; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px 宋体; color:#000000; text-decoration:none; }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
</style>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
<%
REM 管理栏目设置
dim menu(8,10),trs,k
i=0
k=0
set rs=Dvbbs.Execute("select * from dv_help where h_type=1 and h_parentid=0 and not h_stype=1 order by h_id")
do while not rs.eof
	menu(i,k)=rs("h_title")
	'response.write "menu("&i&","&k&")="""&rs("h_title")&""""
	'response.write chr(10)
	k=k+1
	set trs=Dvbbs.Execute("select * from dv_help where h_type=1 and h_parentid="&rs(0)&" and not h_stype=1 order by h_id")
	do while not trs.eof
		menu(i,k)="<a href=admin_help.asp?action=view&id="&trs(0)&" target=main><img src=images/manage/bullet.gif border=0 alt=点击查看该项目的帮助></a>" & trs("h_title")
		'response.write "menu("&i&","&k&")="""&trs("h_title")&""""
		'response.write chr(10)
		k=k+1
	trs.movenext
	loop
	trs.close
	set trs=nothing
	i=i+1
	k=0
rs.movenext
loop
rs.close
set rs=nothing
'response.end
%>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=100% cellpadding=0 cellspacing=0 border=0 align=left>
    <tr><td valign=top>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=42 valign=bottom>
	  <img src="images/manage/title.gif" width=158 height=38>
    </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/manage/title_bg_quit.gif  >
	  <span><a href="admin_index.asp" target=_top><b>管理首页</b></a> | <a href=admin_logout.asp target=_top><b>退出</b></a></span>
    </td>
  </tr>
</table>
&nbsp;
<%
	dim j,i
	dim tmpmenu
	dim menuname
	dim menurl
	Dim TempStr,Menu_1,Menu_2
	TempStr = template.html(0)
	Menu_1 = Split(TempStr,"||")
for i=0 to ubound(Menu_1)
%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
<%
	Menu_2 = Split(Menu_1(i),"@@")
	For j = 0 To Ubound(Menu_2)
	If j=0 Then
%>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/manage/admin_left_<%=i+1%>.gif" id=menuTitle1 onclick="showsubmenu(<%=i%>)">
	  <span><%=Menu_2(0)%></span>
	</td>
  </tr>
  <tr>
    <td style="display" id='submenu<%=i%>'><div class=sec_menu style="width:158"><table cellpadding=0 cellspacing=0 align=center width=150><TBODY>
<%
	Else
	if j=1 then response.write "<tr><td height=5></td></tr>"
%>
<tr><td height=20><img alt src="images/manage/bullet.gif" border="0" width="15" height="20"><%=Menu_2(j)%></td></tr>
<%
	End If
	next
%><TBODY></table></div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=135>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%next%>
&nbsp;
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/manage/admin_left_9.gif" id=menuTitle1>
	  <span>动网论坛信息</span>
	</td>
  </tr>
  <tr>
    <td>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=135>
<tr><td height=20>&nbsp;<br><a href="http://www.aspsky.net/" target=_blank>版权所有：<BR>动网先锋(<font face=Verdana, Arial, Helvetica, sans-serif>AspSky<font color=#CC0000>.Net</font></font>)</a><BR>
<a href="http://www.dvbbs.net/" target=_blank>支持论坛：<BR>动网论坛(<font face=Verdana, Arial, Helvetica, sans-serif>Dvbbs<font color=#CC0000>.Net</font></font>)</a><BR><BR>
</td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
&nbsp;
<%
end sub

Sub admin_login()
	dvbbs.head()
	dvbbs.stats="论坛管理登录"
	if dvbbs.userid=0 then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您不是系统管理员！"
	end if
	If request("reaction")="chklogin" then
		Call chklogin()
	Else
		Call admin_login_main()
	End If
End Sub

sub admin_login_main()
Dim version
If IsSqlDataBase = 1 Then version="SQL 版" Else version="ACCESS 版"
On Error Resume Next
Dim Ados,GetCode
Set Ados=Server.CreateObject("Adodb.Stream")
If Err Then
	GetCode=9999
End If
%>
<html>
<head>
<meta NAME=GENERATOR Content="Microsoft FrontPage 4.0" CHARSET=GB2312>
<meta name=keywords content='动网先锋,动网论坛,dvbbs'>
<title><%=dvbbs.Forum_info(0)%>--<%=dvbbs.stats%></title>
<%=template.html(1)%>
</head>
<body leftmargin=0 bottommargin=0 rightmargin=0 topmargin=0 marginheight=0 marginwidth=0>
<p>&nbsp;</p>
<p>&nbsp;</p>
<form action="admin_index.asp?action=admin_login&reaction=chklogin" method=post>
<table cellpadding="1" cellspacing="0" border="0" align=center style="border: outset 3px;width:0;">
<tr><td>
<table style="width:500" border=0 cellspacing=0 cellpadding=3 align=center class=tablefoot>
    <tr><th valign=middle colspan=2 height=25><%=dvbbs.Forum_info(0)%>管理登录</th></tr>
</table>
<table style="width:500" border=0 cellspacing=0 cellpadding=3 align=center class=tablefoot>
    <tr>
    <td valign=middle colspan=2 align=center class=forumRowHighlight style="background-image: url(images/manage/loginbg.jpg);" height="75">
	<table border="0" width="100%" height="100%">
    <tr><td width="61%" height="100%" rowspan="3"></td>
	<td width="39%" height="0"></td></tr>
    <tr><td height="" valign=top class=tdfoot style=""><BR><a href="index.asp"><b><%=dvbbs.Forum_info(0)%></b></a><br>版本：Dvbbs v7.0.0 <%=version%></td></tr>
    <tr><td height=""></td></tr>
	</table>
	</td></tr>
</table>
<table style="width:500" border=0 cellspacing=0 cellpadding=3 align=center>
	<tr>
    <td valign=middle colspan=2 align=center class=forumRowHighlight height=4></td>
	</tr>
    <tr>
    <td valign=middle class=forumRow width="30%" align=right><b>用户名：</b></td>
    <td valign=middle class=forumRow><INPUT name=username type=text></td></tr>
    <tr>
    <td valign=middle class=forumRow align=right><b>密　码：</b></font></td>
    <td valign=middle class=forumRow><INPUT name=password type=password></td></tr>
    <tr>
    <td valign=middle class=forumRow align=right><b>附加码：</b></td>
    <td valign=middle class=forumRow><INPUT name=verifycode type=text value="<%If GetCode=9999 Then Response.Write "9999"%>">&nbsp;请在附加码框输入 <%=getcode1()%></td></tr>
	<tr>
    <td valign=middle colspan=2 align=center class=forumRowHighlight><input class=button type=submit name=submit value="登 录"></td>
	</tr>
</table>
</td></tr></table>
</form>

</body>
</html>
<%

end sub

sub chklogin()
	username=trim(replace(request("username"),"'",""))
	password=md5(trim(replace(request("password"),"'","")),16)
	'Response.Write session("getcode")
	'Response.Write "<br>"
	'Response.Write request("verifycode")
	'response.end
	if request("verifycode")="" then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请返回输入确认码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
		Exit Sub
	elseif session("getcode")="9999" then
		session("getcode")=""
	elseif session("getcode")="" then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请不要重复提交，如需重新登录请返回登录页面。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
		Exit Sub
	elseif cstr(session("getcode"))<>cstr(trim(request("verifycode"))) then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您输入的确认码和系统产生的不一致，请重新输入。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
		Exit Sub
	end if
	session("getcode")=""
	if username="" or password="" then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>请输入您的用户名或密码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
		Exit Sub
	end if
	ip=Dvbbs.UserTrueIP
	set rs=Dvbbs.Execute("select * from "&admintable&" where username='"&username&"' and adduser='"&dvbbs.membername&"'")
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您输入的用户名和密码不正确或者您不是系统管理员。请<a href=admin_login.asp>重新输入</a>您的密码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
		exit sub
	else
		if trim(rs("password"))<>password then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>您输入的用户名和密码不正确或者您不是系统管理员。请<a href=admin_login.asp>重新输入</a>您的密码。<b>返回后请刷新登录页面后重新输入正确的信息。</b>"
		exit sub
		else
		session("flag")=rs("flag")
		session.timeout=45
		Dvbbs.Execute("update "&admintable&" set LastLogin="&SqlNowString&",LastLoginIP='"&ip&"' where username='"&username&"'")
		rs.close
		set rs=nothing
		response.redirect "admin_index.asp"
		end if
	end if
end sub

sub admin_main()
%>
<title><%=dvbbs.Forum_info(0)%>--管理页面</title>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
<%
if not dvbbs.master or session("flag")="" then
	Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。"
Else
	Dim theInstalledObjects(20)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    
    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"	'Jamil 4.2
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"
	theInstalledObjects(18) = "JMail.Message"	'Jamil 4.3
	theInstalledObjects(19) = "Persits.Upload"
	theInstalledObjects(20) = "SoftArtisans.FileUp"
	Head()
%>
<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>论坛信息统计</th><tr>
<tr><td class="bodytitle" height=23 colspan=2>
<%
dim isaudituser
set rs=Dvbbs.Execute("select count(*) from [dv_user] where usergroupid=5")
isaudituser=rs(0)
if isnull(isaudituser) then isaudituser=0
Dim BoardListNum
set rs=dvbbs.execute("select count(*) from dv_board")
BoardListNum=rs(0)
If isnull(BoardListNum) then BoardListNum=0
set rs=Dvbbs.Execute("select * from dv_setup")
if not rs.eof then
%>
系统信息：论坛帖子数 <B><%=rs("Forum_PostNum")%></B> 主题数 <B><%=rs("Forum_topicnum")%></B> 用户数 <B><%=rs("Forum_usernum")%></B> 待审核用户数 <B><%=isaudituser%></B> 版面总数 <B><%=BoardListNum%></B>
<%
end if
rs.close
set rs=nothing
%>
</td></tr>
<tr><td  class="forumRowHighlight" height=23 colspan=2>
本论坛由动网先锋（aspsky.net）授权给 <%=dvbbs.Forum_info(0)%> 使用，当前使用版本为 动网论坛
<%
If IsSqlDatabase=1 Then
	Response.Write "SQL数据库"
Else
	Response.Write "Access数据库"
End If
Response.Write " Dvbbs " & Dvbbs.Forum_Version
%>
</td></tr>
<tr>
<td width="50%"  class="forumRow" height=23>服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="forumRow">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>站点物理路径：<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
<td width="50%" class="forumRow">数据库地址：</td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>FSO文本读写：<%If Not IsObjInstalled(theInstalledObjects(9)) Then%><font color="<%=dvbbs.mainsetting(1)%>"><b>×</b></font><%else%><b>√</b><%end if%></td>
<td width="50%" class="forumRow">数据库使用：<%If Not IsObjInstalled(theInstalledObjects(10)) Then%><font color="<%=dvbbs.mainsetting(1)%>"><b>×</b></font><%else%><b>√</b><%end if%></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>
<%If IsObjInstalled(theInstalledObjects(18)) Then%>Jmail4.3邮箱组件支持：<%else%>Jmail4.2组件支持：<%end if%>
<%If IsObjInstalled(theInstalledObjects(18)) or IsObjInstalled(theInstalledObjects(13)) Then%>
<b>√</b>
<%else%>
<font color="<%=dvbbs.mainsetting(1)%>"><b>×</b></font>
<%end if%>
</td>
<td width="50%" class="forumRow">CDONTS邮箱组件支持：<%If Not IsObjInstalled(theInstalledObjects(14)) Then%><font color="<%=dvbbs.mainsetting(1)%>"><b>×</b></font><%else%><b>√</b><%end if%></td>
</tr>
<tr><td class="forumRow" height=23 colspan=2>
<%
dim trs
set trs=Dvbbs.Execute("select * from Dv_ChallengeInfo")
set rs=Dvbbs.Execute("select * from dv_setup")
%>
<%if Dvbbs.Forum_ChanSetting(0)=1 and rs("Forum_isinstall")=1 then%>
您已经安装了论坛的短信互动功能，通过短信互动功能，您可以享受到各种不同的网站收益，具体请看下面关于短信互动的说明<BR><BR>
您当前注册的短信互动站点资料是，用户名：<%=trs("d_username")%>，网站名：<%=trs("d_forumname")%>，论坛地址：<%if trs("d_forumurl")="" then%><%=Dvbbs.Get_ScriptNameUrl()%><%else%><%=trs("d_forumurl")%><%end if%>，如果这些资料和您当前所使用的论坛不符（如论坛地址或用户名），您将不能得到相关的短信收益，<a href="install.asp?isnew=1"><font color=blue>您可以点击此处进行资料更新或者重新注册站长资料</font></a>。
<%else%>
您还没有安装了论坛的短信互动功能，通过短信互动功能，您可以享受到各种不同的网站收益，<a href="admin_index.asp?action=admin_main&taction=yes"><font color=red>具体请看关于短信互动的说明</font></a>，在没有安装短信互动功能的情况下，您是不能通过论坛的短信服务获取收益的，<a href="install.asp?isnew=1"><font color=blue>点击此处加入论坛短信互动功能</font></a>（请确保您没有删除原版论坛中的install.asp文件）
<%if request("taction")="yes" then%>
<BR><BR>
<B>论坛站长须知</B>
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
<%end if%>
<%end if%>
<%
rs.close
set rs=nothing
trs.close
set trs=nothing
%>
</td></tr>
</table>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center style="line-height:14pt">
<tr><th class="tableHeaderText" colspan=2 height=25>论坛管理小贴士</th><tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>用户组权限</B>
</td><td class="forumRow" height=23 width="*">动网论坛将注册用户分成不同的用户组，每个用户组可以拥有不同的论坛操作权限，并且在动网论坛7.0版本之后，用户等级结合到了用户组中，假如用户等级没有自定义权限，那么这个等级的权限就使用他所属的用户组权限，反之则拥有这个等级自己的权限。<font color=red>每个等级或者用户组所设定的权限都是是针对整个论坛的</font></td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>分版面权限</B>
</td><td class="forumRow" height=23 width="*">
每个用户组或有自定义权限设置的等级，都可以设置其在论坛中各个版面拥有不同的权限，比如说您可以设置注册用户或者新手上路在版面A不能发贴可以浏览等等权限设置，极大的扩充了论坛权限的设置，<font color=blue>从理论上来说可以分出很多个不同功能类型的论坛</font>
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>用户权限设定</B>
</td><td class="forumRow" height=23 width="*">
每个用户都可以设置其在论坛中各个版面拥有不同的权限或者特殊的权限，比如说您可以设置用户A在版面A中拥有所有管理权限。<U>对于上述三种权限需要注意的是其优先顺序为：用户权限设置(<font color=gray>自定义</font>)<font color=blue> <B>></B> </font>分版面权限设定(<font color=gray>自定义</font>)<font color=blue> <B>></B> </font>用户组权限设定(<font color=gray>默认</font>)</U>
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>对风格模板的管理</B>
</td><td class="forumRow" height=23 width="*">
其中包含对论坛所有模板的管理，模板中论坛的基本CSS设置，论坛主风格的更改，论坛分页面风格的更改，图片的设置，语言包的设置，新建模板页面和模板，模板中新建不同的语言、图片、风格等模板元素等等功能，并且拥有模板的导入导出功能，从真正意义上实现了论坛风格的在线编辑和切换
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>一句话贴士</B>
</td><td class="forumRow" height=23 width="*">
① 对于不同功能模块的页面，要仔细看页面中的说明，以免误操作
<BR>
② 用户组及其扩展的权限设置，对论坛的各种设置有极大扩充性，要充分明白其优先和有效顺序
<BR>
③ 添加论坛大分类的时候，别忘了回头看看该版面高级设置是否正确
<BR>
④ 有问题请到动网论坛官方站点提问，有很多热心的朋友会帮忙，<a href="admin_help.asp">查看更多贴士请点击</a>
</td></tr>
</table>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>论坛管理快捷方式</th><tr>
<FORM METHOD=POST ACTION="admin_user.asp?action=userSearch&userSearch=9&usernamechk=yes"><tr>
<td width="20%"  class="forumRow" height=23>快速查找用户</td>
<td width="80%" class="forumRow">
<input type="text" name="username" size="30"> <input type="submit" value="立刻查找">
<input type="hidden" name="userclass" value="0">
<input type="hidden" name="searchMax" value=100>
</td></FORM>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>快捷功能链接</td>
<td width="80%" class="forumRow"><a href=admin_board.asp?action=add>添加论坛类别</a> | <a href=admin_board.asp>管理论坛版面</a> | <a href="ReloadForumCache.asp">更新服务器缓存</a></td>
</tr>
<tr><form action="admin_update.asp?action=updat" method=post>
<td width="20%" class="forumRow" height=23>快速更新数据</td>
<td width="80%" class="forumRow">
<input type="submit" name="Submit" value="更新论坛数据">&nbsp;
<input type="submit" name="Submit" value="更新论坛总数据">
</td></form>
</tr>
</table>
<%if Dvbbs.Forum_ChanSetting(0)=1 then%>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>关于论坛互动功能</th><tr>
<tr>
<td width="100%" class="forumRow" height=23>
<B>论坛站长须知</B>
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
</td>
</tr>
</table>
<%end if%>
<script language='javascript'> function jumpto(url) { if (url != '') { window.open(url); } } </script>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>动网先锋论坛系统[动网论坛]</th></tr>
<tr>
<td width="20%" class="forumRow" height=23>产品开发</td>
<td width="80%" class="forumRow">
<a href="http://www.aspsky.cn/about.asp" target=_blank>海口动网先锋网络科技有限公司&nbsp;&nbsp;<font color=blue>中国国家版权局著作权登记号2004SR00001</font>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>产品负责</td>
<td width="80%" class="forumRow">
网站事业部 动网论坛项目组&nbsp;&nbsp;<a href="http://www.dvbbs.net/dw.html" target=_blank><font color=blue>企业典型案例</font></a>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>联系方法</td>
<td width="80%" class="forumRow">
网站事业部：0898-68557467 Email eway@aspsky.net<BR>
<a href="http://www.aspsky.cn/" target=_blank><font color=blue>主机事业部</font></a>：0898-68592224 68592294 Email info@aspsky.net<BR>
技　术　部：0898-68592224-13<BR>
传　　　真：0898-68556467<BR>
联系　我们：<a href="http://www.aspsky.cn/email.asp" target=_blank>http://www.aspsky.cn</a><BR>
关于　我们：<a href="http://www.aspsky.cn/about.asp" target=_blank>http://www.aspsky.net</a>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>插件开发</td>
<td width="80%" class="forumRow">
动网论坛插件组织（Dvbbs Plus Organization）
</td>
</tr>
</table>

<%
footer
end if
end sub

Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function

sub admin_head()
%>
<title><%=dvbbs.Forum_info(0)%>--管理页面</title>
<style type="text/css">
a:link { color:#000000;text-decoration:none}
a:hover {color:#666666;}
a:visited {color:#000000;text-decoration:none}

td {FONT-SIZE: 9pt; FILTER: dropshadow(color=#FFFFFF,offx=1,offy=1); COLOR: #000000; FONT-FAMILY: "宋体"}
img {filter:Alpha(opacity:100); chroma(color=#FFFFFF)}
</style>

<script>
function preloadImg(src)
{
	var img=new Image();
	img.src=src
}
preloadImg("images/manage/pic/admin_top_open.gif");

var displayBar=true;
function switchBar(obj)
{
	if (displayBar)
	{
		parent.frame.cols="0,*";
		displayBar=false;
		obj.src="images/manage/admin_top_open.gif";
		obj.title="打开左边管理菜单";
	}else{
		parent.frame.cols="180,*";
		displayBar=true;
		obj.src="images/manage/admin_top_close.gif";
		obj.title="关闭左边管理菜单";
	}
}
</script>
<body background="images/manage/admin_top_bg.gif" leftmargin="0" topmargin="0">
<table width="100%" height="100%" border=0 cellpadding=0 cellspacing=0>
<tr valign=middle>
	<td width=50>
	<img onclick="switchBar(this)" src="images/manage/admin_top_close.gif" title="关闭左边管理菜单" style="cursor:hand">
	</td>
	<td width=150>
		动网论坛系统设置面板
	</td>
	<td width=40>
		<img src="images/manage/admin_top_icon_1.gif">
	</td>
	<td width=100>
		<a href="admin_admin.asp" target=main>修改管理员资料</a>
	</td>
	<%'if Dvbbs.Forum_ChanSetting(0)=1 then%>
	<%
	set rs=Dvbbs.Execute("select top 1 * from dv_challengeinfo")
	%>
	<td width=40>
		<img src="images/manage/admin_top_icon_5.gif">
	</td>
	<td width=100>
		<a href="http://bbs.ray5198.com/login_new.jsp?username=<%=rs("D_username")%>&fourmid=<%=rs("D_ForumID")%>&css=Get_CSS.asp?skinid=1&url=<%=Dvbbs.Get_ScriptNameUrl%>" target=main><font color=blue>站长收益查询</font></a>
	</td>
	<%'end if%>
	<td width=120>
		<a href="http://bbs.dvbbs.net" target=_blank><font color=red>动网论坛官方讨论区</font></a>
	</td>
	<td width=*>
		<a href="index.asp" target=_top>返回论坛首页</a>
	</td>
	<td>
	&nbsp;
	</td>
</tr>
</table>
<%
end Sub
Function getcode1()
	Dim test
	On Error Resume Next
	Set test=Server.CreateObject("Adodb.Stream")
	Set test=Nothing
	If Err Then
		Dim zNum
		Randomize timer
		zNum = cint(8999*Rnd+1000)
		Session("GetCode") = zNum
		getcode1= Session("GetCode")		
	Else
		getcode1= "<img src=""getcode.asp"">"		
	End If
End Function
%>