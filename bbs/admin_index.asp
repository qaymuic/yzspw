<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/md5.asp" -->
<!-- #include file="inc/myadmin.asp" -->
<!-- #include file="inc/dvadchar.asp" -->
<%
dim username
dim password
dim ip
dvbbs.stats="��̳����������"
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
<title><%=dvbbs.Forum_info(0)%>--�������</title>
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
<title><%=dvbbs.Forum_info(0)%>--����ҳ��</title>
<style type=text/css>
body  { background:#799AE1; margin:0px; font:normal 12px ����; 
SCROLLBAR-FACE-COLOR: #799AE1; SCROLLBAR-HIGHLIGHT-COLOR: #799AE1; 
SCROLLBAR-SHADOW-COLOR: #799AE1; SCROLLBAR-DARKSHADOW-COLOR: #799AE1; 
SCROLLBAR-3DLIGHT-COLOR: #799AE1; SCROLLBAR-ARROW-COLOR: #FFFFFF;
SCROLLBAR-TRACK-COLOR: #AABFEC;
}
table  { border:0px; }
td  { font:normal 12px ����; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px ����; color:#000000; text-decoration:none; }
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
REM ������Ŀ����
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
		menu(i,k)="<a href=admin_help.asp?action=view&id="&trs(0)&" target=main><img src=images/manage/bullet.gif border=0 alt=����鿴����Ŀ�İ���></a>" & trs("h_title")
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
	  <span><a href="admin_index.asp" target=_top><b>������ҳ</b></a> | <a href=admin_logout.asp target=_top><b>�˳�</b></a></span>
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
	  <span>������̳��Ϣ</span>
	</td>
  </tr>
  <tr>
    <td>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=135>
<tr><td height=20>&nbsp;<br><a href="http://www.aspsky.net/" target=_blank>��Ȩ���У�<BR>�������ȷ�(<font face=Verdana, Arial, Helvetica, sans-serif>AspSky<font color=#CC0000>.Net</font></font>)</a><BR>
<a href="http://www.dvbbs.net/" target=_blank>֧����̳��<BR>��������̳(<font face=Verdana, Arial, Helvetica, sans-serif>Dvbbs<font color=#CC0000>.Net</font></font>)</a><BR><BR>
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
	dvbbs.stats="��̳�����¼"
	if dvbbs.userid=0 then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>������ϵͳ����Ա��"
	end if
	If request("reaction")="chklogin" then
		Call chklogin()
	Else
		Call admin_login_main()
	End If
End Sub

sub admin_login_main()
Dim version
If IsSqlDataBase = 1 Then version="SQL ��" Else version="ACCESS ��"
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
<meta name=keywords content='�����ȷ�,������̳,dvbbs'>
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
    <tr><th valign=middle colspan=2 height=25><%=dvbbs.Forum_info(0)%>�����¼</th></tr>
</table>
<table style="width:500" border=0 cellspacing=0 cellpadding=3 align=center class=tablefoot>
    <tr>
    <td valign=middle colspan=2 align=center class=forumRowHighlight style="background-image: url(images/manage/loginbg.jpg);" height="75">
	<table border="0" width="100%" height="100%">
    <tr><td width="61%" height="100%" rowspan="3"></td>
	<td width="39%" height="0"></td></tr>
    <tr><td height="" valign=top class=tdfoot style=""><BR><a href="index.asp"><b><%=dvbbs.Forum_info(0)%></b></a><br>�汾��Dvbbs v7.0.0 <%=version%></td></tr>
    <tr><td height=""></td></tr>
	</table>
	</td></tr>
</table>
<table style="width:500" border=0 cellspacing=0 cellpadding=3 align=center>
	<tr>
    <td valign=middle colspan=2 align=center class=forumRowHighlight height=4></td>
	</tr>
    <tr>
    <td valign=middle class=forumRow width="30%" align=right><b>�û�����</b></td>
    <td valign=middle class=forumRow><INPUT name=username type=text></td></tr>
    <tr>
    <td valign=middle class=forumRow align=right><b>�ܡ��룺</b></font></td>
    <td valign=middle class=forumRow><INPUT name=password type=password></td></tr>
    <tr>
    <td valign=middle class=forumRow align=right><b>�����룺</b></td>
    <td valign=middle class=forumRow><INPUT name=verifycode type=text value="<%If GetCode=9999 Then Response.Write "9999"%>">&nbsp;���ڸ���������� <%=getcode1()%></td></tr>
	<tr>
    <td valign=middle colspan=2 align=center class=forumRowHighlight><input class=button type=submit name=submit value="�� ¼"></td>
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
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�뷵������ȷ���롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
		Exit Sub
	elseif session("getcode")="9999" then
		session("getcode")=""
	elseif session("getcode")="" then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�벻Ҫ�ظ��ύ���������µ�¼�뷵�ص�¼ҳ�档<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
		Exit Sub
	elseif cstr(session("getcode"))<>cstr(trim(request("verifycode"))) then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
		Exit Sub
	end if
	session("getcode")=""
	if username="" or password="" then
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>�����������û��������롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
		Exit Sub
	end if
	ip=Dvbbs.UserTrueIP
	set rs=Dvbbs.Execute("select * from "&admintable&" where username='"&username&"' and adduser='"&dvbbs.membername&"'")
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��������û��������벻��ȷ����������ϵͳ����Ա����<a href=admin_login.asp>��������</a>�������롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
		exit sub
	else
		if trim(rs("password"))<>password then
			Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��������û��������벻��ȷ����������ϵͳ����Ա����<a href=admin_login.asp>��������</a>�������롣<b>���غ���ˢ�µ�¼ҳ�������������ȷ����Ϣ��</b>"
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
<title><%=dvbbs.Forum_info(0)%>--����ҳ��</title>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" >
<%
if not dvbbs.master or session("flag")="" then
	Response.Redirect "showerr.asp?action=OtherErr&ErrCodes=<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣"
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
<tr><th class="tableHeaderText" colspan=2 height=25>��̳��Ϣͳ��</th><tr>
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
ϵͳ��Ϣ����̳������ <B><%=rs("Forum_PostNum")%></B> ������ <B><%=rs("Forum_topicnum")%></B> �û��� <B><%=rs("Forum_usernum")%></B> ������û��� <B><%=isaudituser%></B> �������� <B><%=BoardListNum%></B>
<%
end if
rs.close
set rs=nothing
%>
</td></tr>
<tr><td  class="forumRowHighlight" height=23 colspan=2>
����̳�ɶ����ȷ棨aspsky.net����Ȩ�� <%=dvbbs.Forum_info(0)%> ʹ�ã���ǰʹ�ð汾Ϊ ������̳
<%
If IsSqlDatabase=1 Then
	Response.Write "SQL���ݿ�"
Else
	Response.Write "Access���ݿ�"
End If
Response.Write " Dvbbs " & Dvbbs.Forum_Version
%>
</td></tr>
<tr>
<td width="50%"  class="forumRow" height=23>���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="forumRow">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>վ������·����<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
<td width="50%" class="forumRow">���ݿ��ַ��</td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>FSO�ı���д��<%If Not IsObjInstalled(theInstalledObjects(9)) Then%><font color="<%=dvbbs.mainsetting(1)%>"><b>��</b></font><%else%><b>��</b><%end if%></td>
<td width="50%" class="forumRow">���ݿ�ʹ�ã�<%If Not IsObjInstalled(theInstalledObjects(10)) Then%><font color="<%=dvbbs.mainsetting(1)%>"><b>��</b></font><%else%><b>��</b><%end if%></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>
<%If IsObjInstalled(theInstalledObjects(18)) Then%>Jmail4.3�������֧�֣�<%else%>Jmail4.2���֧�֣�<%end if%>
<%If IsObjInstalled(theInstalledObjects(18)) or IsObjInstalled(theInstalledObjects(13)) Then%>
<b>��</b>
<%else%>
<font color="<%=dvbbs.mainsetting(1)%>"><b>��</b></font>
<%end if%>
</td>
<td width="50%" class="forumRow">CDONTS�������֧�֣�<%If Not IsObjInstalled(theInstalledObjects(14)) Then%><font color="<%=dvbbs.mainsetting(1)%>"><b>��</b></font><%else%><b>��</b><%end if%></td>
</tr>
<tr><td class="forumRow" height=23 colspan=2>
<%
dim trs
set trs=Dvbbs.Execute("select * from Dv_ChallengeInfo")
set rs=Dvbbs.Execute("select * from dv_setup")
%>
<%if Dvbbs.Forum_ChanSetting(0)=1 and rs("Forum_isinstall")=1 then%>
���Ѿ���װ����̳�Ķ��Ż������ܣ�ͨ�����Ż������ܣ����������ܵ����ֲ�ͬ����վ���棬�����뿴������ڶ��Ż�����˵��<BR><BR>
����ǰע��Ķ��Ż���վ�������ǣ��û�����<%=trs("d_username")%>����վ����<%=trs("d_forumname")%>����̳��ַ��<%if trs("d_forumurl")="" then%><%=Dvbbs.Get_ScriptNameUrl()%><%else%><%=trs("d_forumurl")%><%end if%>�������Щ���Ϻ�����ǰ��ʹ�õ���̳����������̳��ַ���û��������������ܵõ���صĶ������棬<a href="install.asp?isnew=1"><font color=blue>�����Ե���˴��������ϸ��»�������ע��վ������</font></a>��
<%else%>
����û�а�װ����̳�Ķ��Ż������ܣ�ͨ�����Ż������ܣ����������ܵ����ֲ�ͬ����վ���棬<a href="admin_index.asp?action=admin_main&taction=yes"><font color=red>�����뿴���ڶ��Ż�����˵��</font></a>����û�а�װ���Ż������ܵ�����£����ǲ���ͨ����̳�Ķ��ŷ����ȡ����ģ�<a href="install.asp?isnew=1"><font color=blue>����˴�������̳���Ż�������</font></a>����ȷ����û��ɾ��ԭ����̳�е�install.asp�ļ���
<%if request("taction")="yes" then%>
<BR><BR>
<B>��̳վ����֪</B>
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
<tr><th class="tableHeaderText" colspan=2 height=25>��̳����С��ʿ</th><tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>�û���Ȩ��</B>
</td><td class="forumRow" height=23 width="*">������̳��ע���û��ֳɲ�ͬ���û��飬ÿ���û������ӵ�в�ͬ����̳����Ȩ�ޣ������ڶ�����̳7.0�汾֮���û��ȼ���ϵ����û����У������û��ȼ�û���Զ���Ȩ�ޣ���ô����ȼ���Ȩ�޾�ʹ�����������û���Ȩ�ޣ���֮��ӵ������ȼ��Լ���Ȩ�ޡ�<font color=red>ÿ���ȼ������û������趨��Ȩ�޶��������������̳��</font></td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>�ְ���Ȩ��</B>
</td><td class="forumRow" height=23 width="*">
ÿ���û�������Զ���Ȩ�����õĵȼ�������������������̳�и�������ӵ�в�ͬ��Ȩ�ޣ�����˵����������ע���û�����������·�ڰ���A���ܷ�����������ȵ�Ȩ�����ã��������������̳Ȩ�޵����ã�<font color=blue>����������˵���Էֳ��ܶ����ͬ�������͵���̳</font>
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>�û�Ȩ���趨</B>
</td><td class="forumRow" height=23 width="*">
ÿ���û�����������������̳�и�������ӵ�в�ͬ��Ȩ�޻��������Ȩ�ޣ�����˵�����������û�A�ڰ���A��ӵ�����й���Ȩ�ޡ�<U>������������Ȩ����Ҫע�����������˳��Ϊ���û�Ȩ������(<font color=gray>�Զ���</font>)<font color=blue> <B>></B> </font>�ְ���Ȩ���趨(<font color=gray>�Զ���</font>)<font color=blue> <B>></B> </font>�û���Ȩ���趨(<font color=gray>Ĭ��</font>)</U>
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>�Է��ģ��Ĺ���</B>
</td><td class="forumRow" height=23 width="*">
���а�������̳����ģ��Ĺ���ģ������̳�Ļ���CSS���ã���̳�����ĸ��ģ���̳��ҳ����ĸ��ģ�ͼƬ�����ã����԰������ã��½�ģ��ҳ���ģ�壬ģ�����½���ͬ�����ԡ�ͼƬ������ģ��Ԫ�صȵȹ��ܣ�����ӵ��ģ��ĵ��뵼�����ܣ�������������ʵ������̳�������߱༭���л�
</td></tr>
<tr><td class="forumRow" height=23 width="80" valign=top>
<B>һ�仰��ʿ</B>
</td><td class="forumRow" height=23 width="*">
�� ���ڲ�ͬ����ģ���ҳ�棬Ҫ��ϸ��ҳ���е�˵�������������
<BR>
�� �û��鼰����չ��Ȩ�����ã�����̳�ĸ��������м��������ԣ�Ҫ������������Ⱥ���Ч˳��
<BR>
�� �����̳������ʱ�򣬱����˻�ͷ�����ð���߼������Ƿ���ȷ
<BR>
�� �������뵽������̳�ٷ�վ�����ʣ��кܶ����ĵ����ѻ��æ��<a href="admin_help.asp">�鿴������ʿ����</a>
</td></tr>
</table>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>��̳�����ݷ�ʽ</th><tr>
<FORM METHOD=POST ACTION="admin_user.asp?action=userSearch&userSearch=9&usernamechk=yes"><tr>
<td width="20%"  class="forumRow" height=23>���ٲ����û�</td>
<td width="80%" class="forumRow">
<input type="text" name="username" size="30"> <input type="submit" value="���̲���">
<input type="hidden" name="userclass" value="0">
<input type="hidden" name="searchMax" value=100>
</td></FORM>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>��ݹ�������</td>
<td width="80%" class="forumRow"><a href=admin_board.asp?action=add>�����̳���</a> | <a href=admin_board.asp>������̳����</a> | <a href="ReloadForumCache.asp">���·���������</a></td>
</tr>
<tr><form action="admin_update.asp?action=updat" method=post>
<td width="20%" class="forumRow" height=23>���ٸ�������</td>
<td width="80%" class="forumRow">
<input type="submit" name="Submit" value="������̳����">&nbsp;
<input type="submit" name="Submit" value="������̳������">
</td></form>
</tr>
</table>
<%if Dvbbs.Forum_ChanSetting(0)=1 then%>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>������̳��������</th><tr>
<tr>
<td width="100%" class="forumRow" height=23>
<B>��̳վ����֪</B>
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
</td>
</tr>
</table>
<%end if%>
<script language='javascript'> function jumpto(url) { if (url != '') { window.open(url); } } </script>
<p></p>

<table cellpadding="3" cellspacing="1" border="0" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>�����ȷ���̳ϵͳ[������̳]</th></tr>
<tr>
<td width="20%" class="forumRow" height=23>��Ʒ����</td>
<td width="80%" class="forumRow">
<a href="http://www.aspsky.cn/about.asp" target=_blank>���ڶ����ȷ�����Ƽ����޹�˾&nbsp;&nbsp;<font color=blue>�й����Ұ�Ȩ������Ȩ�ǼǺ�2004SR00001</font>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>��Ʒ����</td>
<td width="80%" class="forumRow">
��վ��ҵ�� ������̳��Ŀ��&nbsp;&nbsp;<a href="http://www.dvbbs.net/dw.html" target=_blank><font color=blue>��ҵ���Ͱ���</font></a>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>��ϵ����</td>
<td width="80%" class="forumRow">
��վ��ҵ����0898-68557467 Email eway@aspsky.net<BR>
<a href="http://www.aspsky.cn/" target=_blank><font color=blue>������ҵ��</font></a>��0898-68592224 68592294 Email info@aspsky.net<BR>
������������0898-68592224-13<BR>
���������棺0898-68556467<BR>
��ϵ�����ǣ�<a href="http://www.aspsky.cn/email.asp" target=_blank>http://www.aspsky.cn</a><BR>
���ڡ����ǣ�<a href="http://www.aspsky.cn/about.asp" target=_blank>http://www.aspsky.net</a>
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>�������</td>
<td width="80%" class="forumRow">
������̳�����֯��Dvbbs Plus Organization��
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
<title><%=dvbbs.Forum_info(0)%>--����ҳ��</title>
<style type="text/css">
a:link { color:#000000;text-decoration:none}
a:hover {color:#666666;}
a:visited {color:#000000;text-decoration:none}

td {FONT-SIZE: 9pt; FILTER: dropshadow(color=#FFFFFF,offx=1,offy=1); COLOR: #000000; FONT-FAMILY: "����"}
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
		obj.title="����߹���˵�";
	}else{
		parent.frame.cols="180,*";
		displayBar=true;
		obj.src="images/manage/admin_top_close.gif";
		obj.title="�ر���߹���˵�";
	}
}
</script>
<body background="images/manage/admin_top_bg.gif" leftmargin="0" topmargin="0">
<table width="100%" height="100%" border=0 cellpadding=0 cellspacing=0>
<tr valign=middle>
	<td width=50>
	<img onclick="switchBar(this)" src="images/manage/admin_top_close.gif" title="�ر���߹���˵�" style="cursor:hand">
	</td>
	<td width=150>
		������̳ϵͳ�������
	</td>
	<td width=40>
		<img src="images/manage/admin_top_icon_1.gif">
	</td>
	<td width=100>
		<a href="admin_admin.asp" target=main>�޸Ĺ���Ա����</a>
	</td>
	<%'if Dvbbs.Forum_ChanSetting(0)=1 then%>
	<%
	set rs=Dvbbs.Execute("select top 1 * from dv_challengeinfo")
	%>
	<td width=40>
		<img src="images/manage/admin_top_icon_5.gif">
	</td>
	<td width=100>
		<a href="http://bbs.ray5198.com/login_new.jsp?username=<%=rs("D_username")%>&fourmid=<%=rs("D_ForumID")%>&css=Get_CSS.asp?skinid=1&url=<%=Dvbbs.Get_ScriptNameUrl%>" target=main><font color=blue>վ�������ѯ</font></a>
	</td>
	<%'end if%>
	<td width=120>
		<a href="http://bbs.dvbbs.net" target=_blank><font color=red>������̳�ٷ�������</font></a>
	</td>
	<td width=*>
		<a href="index.asp" target=_top>������̳��ҳ</a>
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