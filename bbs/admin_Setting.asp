<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%	
	Head()
	dim admin_flag,rs_c
	admin_flag=",1,"
	If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then 
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		Call dvbbs_error()
	Else
		if request("action")="save" then
		call saveconst()
		elseif request("action")="restore" then
		call restore()
		else
		call consted()
		end if
		Footer()
	end if

Sub consted()
Dim  sel
%>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form method="POST" action="admin_setting.asp?action=save" name="theform">
<tr> 
<th width="100%" colspan=3 class="tableHeaderText" height=25>论坛基本设置（目前只提供一种设置)
</th></tr>
<tr> 
<td width="100%" class=Forumrow colspan=3 height=23>
<a href="#setting3">[基本信息]</a>&nbsp;<a href="#setting21">[论坛系统数据设置]</a>&nbsp;<a href="#setting6">[悄悄话选项]</a>&nbsp;<a href="#setting7">[论坛首页选项]</a>&nbsp;<a href="#setting8">[用户与注册选项]</a>&nbsp;<a href="#setting10">[系统设置]</a>&nbsp;<a href="#setting12">[在线和用户来源]</a>&nbsp;<a href="admin_challenge.asp">[<font color=blue>论坛短信设置</font>]</a>
</td>
</tr>
<tr> 
<td width="100%" class=Forumrow colspan=3 height=23>
<a href="#setting13">[邮件选项]</a>&nbsp;<a href="#setting14">[上传设置]</a>　&nbsp;<a href="#setting15">[用户选项(签名、头衔、排行等)]</a>　<a href="#setting16">[帖子选项]</a>&nbsp;<a href="#setting17">[防刷新机制]</a>&nbsp;<a href="#setting18">[论坛分页设置]</a>&nbsp;<a href="#setting19">[门派设置]</a>
</td>
</tr>
<tr> 
<td width="100%" class=Forumrow colspan=3 height=23>
<a href="#setting20">[搜索选项]</a>&nbsp;<a href="#settingxu">[虚拟形象选项]</a>
</td>
</tr>
<tr> 
<td width="93%" class=bodytitle colspan=2 height=23>
如果您的论坛的设置搞乱了，可以使用<a href="?action=restore"><B>还原论坛默认设置</B></a>
</td>
<input type="hidden" id="forum_return" value="<b>还原论坛默认设置:</b><br><li>如果您把论坛设置搞乱了，可以点击还原论坛默认设置进行还原操作。<br><li>使用此操作将使您原来的设置无效而还原到论坛的默认设置，请确认您做了论坛备份或者记得还原后该做哪些针对您论坛所需要的设置">
<td class=bodytitle><a href=# onclick="helpscript(forum_return);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>论坛默认使用风格</U></td>
<td width="43%" class=Forumrow>
<%
	Dim forum_sid,iforum_setting,stopreadme,forum_pack
	Dim Style_Option,css_Option,Forum_cid,TempOption
	set rs=dvbbs.execute("select forum_sid,forum_setting,forum_pack,forum_cid from dv_setup")
	Forum_sid=rs(0)
	Forum_pack=Split(rs(2),"|||")
	Iforum_setting=split(rs(1),"|||")
	Forum_cid=rs(3)
	Rs.close:Set Rs=Nothing
	stopreadme=iforum_setting(5)

	set rs_c= server.CreateObject ("adodb.recordset")
	sql = "select id,StyleName,Forum_CSS from dv_style"
	rs_c.open sql,conn,1,1
	if rs_c.eof and rs_c.bof then
		response.write "请先添加风格"
	else
		sql=rs_c.GetRows(-1)
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">"
		Response.Write chr(10)
		Response.Write "var StyleId="&forum_sid&";"
		Response.Write "var Cssid="&Forum_cid&";"
		Response.Write "var css_Option=new Array();"
		Response.Write chr(10)
		For i=0 To Ubound(SQL,2)
			Style_Option=Style_Option+"<option value="
			Style_Option=Style_Option&SQL(0,i)
			If forum_sid=SQL(0,i) Then Style_Option=Style_Option+" selected "
			Style_Option=Style_Option+" >"+SQL(1,i)+"</option>"
			TempOption=Split(SQL(2,i),"@@@")
			Response.Write "css_Option["&SQL(0,i)&"]='"&TempOption(0)&"';"
			Response.Write chr(10)
		Next
		Response.Write "</SCRIPT>"
	End If
	rs_c.close:Set rs_c=Nothing
%>
模板：<select name=sid onChange="Changeoption(this.value)">
<%
Response.Write Style_Option
%>
</select>
 风格：<select name=cid onChange="">
<option value="" >选择风格皮肤</option>
</select>
<SCRIPT LANGUAGE="JavaScript">
<!--
function Changeoption(sid)
{
var NewOption=css_Option[sid].split("|||");
var j=eval('document.theform.cid.length;');
	for (i=0;i<j;i++){
		eval('document.theform.cid.options[j-i]=null;')
	}
	for (i=0;i<NewOption.length-1;i++){
		tempoption=new Option(NewOption[i],i);
		eval('document.theform.cid.options[i]=tempoption;');
		if (Cssid==i&&sid==StyleId){
		eval('document.theform.cid.options[i].selected=true;');
		}
	}
}
var forum_sid=eval('document.theform.sid.value;');
Changeoption(forum_sid);
//-->
</SCRIPT>
</td>
<input type="hidden" id="forum_skin" value="<b>论坛默认使用风格:</b><br><li>在这里您可以选择您论坛的默认使用风格。<br><li>如果想改变论坛风格请到论坛风格模板管理中进行相关设置">
<td class=Forumrow><a href=# onclick="helpscript(forum_skin);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=forumRowHighlight><U>论坛当前状态</U><BR>维护期间可设置关闭论坛</td>
<td class=forumRowHighlight> 
<input type=radio name="forum_setting(21)" value=0 <%if Dvbbs.forum_setting(21)="0" then%>checked<%end if%>>打开&nbsp;
<input type=radio name="forum_setting(21)" value=1 <%if Dvbbs.forum_setting(21)="1" then%>checked<%end if%>>关闭&nbsp;
</td>
<input type="hidden" id="forum_open" value="<b>论坛当前状态:</b><br><li>如果您需要做更改程序、更新数据或者转移站点等需要暂时关闭论坛的操作，可在此处选择关闭论坛。<br><li>关闭论坛后，可直接使用论坛地址＋login.asp登录论坛，然后使用论坛地址＋admin_index.asp登录后台管理进行打开论坛的操作">
<td class=forumRowHighlight><a href=# onclick="helpscript(forum_open);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=Forumrow><U>维护说明</U><BR>在论坛关闭情况下显示，支持html语法</td>
<td class=Forumrow> 
<textarea name="StopReadme" cols="50" rows="3" ID="TDStopReadme"><%=Stopreadme%></textarea><br><a href="javascript:admin_Size(-3,'TDStopReadme')"><img src="images/manage/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'TDStopReadme')"><img src="images/manage/plus.gif" unselectable="on" border='0'></a>
</td>
<input type="hidden" id="forum_opens" value="<b>论坛维护说明:</b><br><li>如果您在论坛当前状态中关闭了论坛，请在此输入维护说明，他将显示在论坛的前台给会员浏览，告知论坛关闭的原因，在这里可以使用HTML语法。">
<td class=forumRow><a href=# onclick="helpscript(forum_opens);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=forumRowHighlight>
<U>论坛定时设置：</U></td>
<td class=forumRowHighlight> 
<input type=radio name="forum_setting(69)" value="0" <%If Dvbbs.forum_setting(69)="0" Then %>checked <%End If%>>关 闭</option>
<input type=radio name="forum_setting(69)" value="1" <%If Dvbbs.forum_setting(69)="1" Then %>checked <%End If%>>定时关闭</option>
<input type=radio name="forum_setting(69)" value="2" <%If Dvbbs.forum_setting(69)="2" Then %>checked <%End If%>>定时只读</option>
</td>
<input type="hidden" id="forum_isopentime" value="<b>定时设置选择:</b><br><li>在这里您可以设置是否起用定时的各种功能，如果开启了本功能，请设置好下面选项中的论坛设置时间。<br><li>如果在非开放时间内需要更改本设置，可直接使用论坛地址＋login.asp登录论坛，然后使用论坛地址＋admin_index.asp登录后台管理进行打开论坛的操作">
<td class=forumRowHighlight><a href=# onclick="helpscript(forum_isopentime);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=Forumrow>
<U>定时设置</U><BR>请根据需要选择开或关</td>
<td class=Forumrow> 
<%
Dvbbs.forum_setting(70)=split(Dvbbs.forum_setting(70),"|")
If UBound(Dvbbs.forum_setting(70))<2 Then 
	Dvbbs.forum_setting(70)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
	Dvbbs.forum_setting(70)=split(Dvbbs.forum_setting(70),"|")
End If
For i= 0 to UBound(Dvbbs.forum_setting(70))
If i<10 Then Response.Write "&nbsp;"
%>
  <%=i%>点：<input type="checkbox" name="forum_setting(70)<%=i%>" value="1" <%If Dvbbs.forum_setting(70)(i)="1" Then %>checked<%End If%>>开
 <%
 If (i+1) mod 4 = 0 Then Response.Write "<br>"
 Next
 %>
</td>
<input type="hidden" id="forum_opentime" value="<b>论坛开放时间:</b><br><li>设置本选项请确认您打开了定时开放论坛功能。<br><li>本设置以小时为单位，请务必按规定正确填写<br><li>如果在非开放时间内需要更改本设置，可直接使用论坛地址＋login.asp登录论坛，然后使用论坛地址＋admin_index.asp登录后台管理进行打开论坛的操作">
<td class=forumRow><a href=# onclick="helpscript(forum_opentime);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=3 class="tableHeaderText" height=25>论坛自动更新通知设置
</th></tr>
<tr> 
<td class=forumRow width="50%">
<U>是否起用动网自动更新通知系统</U></td>
<td class=forumRow width="43%"> 
<input type=radio name="forum_pack(0)" value=0 <%if cint(forum_pack(0))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_pack(0)" value=1 <%if cint(forum_pack(0))=1 then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="forum_pack1" value="<b>是否起用动网自动更新通知系统:</b><br><li>开启后管理后台顶部会提示动网的最新程序、补丁、通知等。">
<td class=forumRow><a href=# onclick="helpscript(forum_pack1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumrowHighlight>
<U>开启通知系统用户名与密码</U><BR>用户名与密码用符号“|||”分开</td>
<td class=ForumrowHighlight>
<%
If UBound(forum_pack)<2 Then ReDim forum_pack(3)
%>
<input type=text size=21 name="forum_pack(1)" value="<%=forum_pack(1)%>|||<%=forum_pack(2)%>">
</td>
<input type="hidden" id="forum_pack2" value="<b>开启通知系统用户名与密码:</b><br><li>如要开启通知系统，请您先到动网官方论坛注册一个用户名并在动网通知系统里取得密码，并填写于此栏即可开启。">
<td class=forumRowHighlight><a href=# onclick="helpscript(forum_pack2);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting3"></a><b>论坛基本信息</b>[<a href="#top">顶部</a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛名称</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_info(0)" size="35" value="<%=Dvbbs.Forum_info(0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>论坛的访问地址</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_info(1)" size="35" value="<%=Dvbbs.Forum_info(1)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>论坛的创建日期(格式：YYYY-M-D)</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="forum_setting(74)" size="35" value="<%=Dvbbs.forum_setting(74)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRow> <U>论坛首页文件名</U></td>
<td width="50%" class=forumRow>  
<input type="text" name="Forum_info(11)" size="35" value="<%=Dvbbs.Forum_info(11)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>网站主页名称</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_info(2)" size="35" value="<%=Dvbbs.Forum_info(2)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRow> <U>网站主页访问地址</U></td>
<td width="50%" class=forumRow>  
<input type="text" name="Forum_info(3)" size="35" value="<%=Dvbbs.Forum_info(3)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>论坛管理员Email</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_info(5)" size="35" value="<%=Dvbbs.Forum_info(5)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRow> <U>联系我们的链接（不填写为Mailto管理员）</U></td>
<td width="50%" class=forumRow>  
<input type="text" name="Forum_info(7)" size="35" value="<%=Dvbbs.Forum_info(7)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>论坛首页Logo图片地址</U><BR>显示在论坛顶部左上角，可用相对路径或者绝对路径</td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_info(6)" size="35" value="<%=Dvbbs.Forum_info(6)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRow> <U>站点关键字</U><BR>将被搜索引擎用来搜索您网站的关键内容<BR>每个关键字用“|”号分隔</td>
<td width="50%" class=forumRow>  
<input type="text" name="Forum_info(8)" size="35" value="<%=Dvbbs.Forum_info(8)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>站点描述</U><BR>将被搜索引擎用来说明您网站的主要内容<BR><font color=red>介绍中请不要带英文的逗号</font></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_info(10)" size="35" value="<%=Dvbbs.Forum_info(10)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRow> <U>论坛版权信息</U></td>
<td width="50%" class=forumRow valign=top>  
<textarea name="Copyright" cols="50" rows="5" id=TdCopyright><%=Dvbbs.Forum_Copyright%></textarea>
<a href="javascript:admin_Size(-5,'TdCopyright')"><img src="images/manage/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(5,'TdCopyright')"><img src="images/manage/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting21"></a><b>论坛系统数据设置</b>[<a href="#top">顶部</a>]--(以下信息不建议用户修改)</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛会员总数</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_UserNum" size="25" value="<%=Dvbbs.CacheData(10,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>论坛主题总数</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_TopicNum" size="25" value="<%=Dvbbs.CacheData(7,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛帖子总数</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_PostNum" size="25" value="<%=Dvbbs.CacheData(8,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>论坛最高日发贴</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_MaxPostNum" size="25" value="<%=Dvbbs.CacheData(12,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛最高日发贴发生时间</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_MaxPostDate" size="25" value="<%=Dvbbs.CacheData(13,0)%>">(格式：YYYY-M-D H:M:S)
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>历史最高同时在线纪录人数</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_Maxonline" size="25" value="<%=Dvbbs.Maxonline%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>历史最高同时在线纪录发生时间</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_MaxonlineDate" size="25" value="<%=Dvbbs.CacheData(6,0)%>">(格式：YYYY-M-D H:M:S)
</td>
</tr>
</table><BR>

<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting6"></a><b>悄悄话选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>新短消息弹出窗口</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(10)" value=0 <%if Dvbbs.forum_setting(10)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(10)" value=1 <%if Dvbbs.forum_setting(10)="1" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发论坛短消息是否采用验证码</U><BR>开启此项可以防止恶意短消息</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(80)" value=0 <%if Dvbbs.forum_setting(80)="0" Then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(80)" value=1 <%if Dvbbs.forum_setting(80)="1" Then%>checked<%end if%>>是&nbsp;
</td>
</tr>
</table><BR>

<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><a name="setting7"></a><b>论坛首页选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr>
<td width="50%" class=Forumrow>
<U>首页显示论坛深度</U>
<input type="hidden" id="forum_depth" value="<b>首页显示论坛深度帮助:</b><br><li>0代表一级，1代表2级，以此类推；<li>设置过大的论坛深度将影响论坛整体性能，请根据自己论坛情况做设置，建议设置为1。">
</td>
<td width="43%" class=Forumrow> 
<input type=text size=10 name="forum_setting(5)" value="<%=Dvbbs.forum_setting(5)%>"> 级
</td>
<td class=Forumrow><a href=# onclick="helpscript(forum_depth);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=forumRowHighlight> <U>是否显示过生日会员</U>
<input type="hidden" id="forum_userbirthday" value="<b>首页显示过生日会员帮助:</b><br><li>凡当天有会员过生日则显示于论坛首页；<li>开启本功能较消耗资源。">
</td>
<td class=forumRowHighlight>  
<input type=radio name="forum_setting(29)" value=0 <%if Dvbbs.forum_setting(29)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(29)" value=1 <%if Dvbbs.forum_setting(29)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=forumRowHighlight><a href=# onclick="helpscript(forum_userbirthday);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
</table><BR>

<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting8"></a><b>用户与注册选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否允许新用户注册</U><BR>关闭后论坛将不能注册</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(37)" value=0 <%if Dvbbs.forum_setting(37)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(37)" value=1 <%if Dvbbs.forum_setting(37)="1" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>注册是否采用验证码</U><BR>开启此项可以防止恶意注册</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(78)" value=0 <%if Dvbbs.forum_setting(78)="0" Then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(78)" value=1 <%if Dvbbs.forum_setting(78)="1" Then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>登录是否采用验证码</U><BR>开启此项可以防止恶意登录猜解密码</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(79)" value=0 <%if Dvbbs.forum_setting(79)="0" Then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(79)" value=1 <%if Dvbbs.forum_setting(79)="1" Then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>会员取回密码是否采用验证码</U><BR>开启此项可以防止恶意登录猜解密码</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(81)" value=0 <%if Dvbbs.forum_setting(81)="0" Then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(81)" value=1 <%if Dvbbs.forum_setting(81)="1" Then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>会员取回密码次数限制</U><BR>0则表示无限制，若取回问答错误超过此限制，则停止至24小时后才能再次使用取回密码功能。</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(84)" size="3" value="<%=Dvbbs.forum_setting(84)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>最短用户名长度</U><BR>填写数字，不能小于1大于50</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(40)" size="3" value="<%=Dvbbs.forum_setting(40)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>最长用户名长度</U><BR>填写数字，不能小于1大于50</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(41)" size="3" value="<%=Dvbbs.forum_setting(41)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>同一IP注册间隔时间</U><BR>如不想限制可填写0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(22)" size="3" value="<%=Dvbbs.forum_setting(22)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>Email通知密码</U><BR>确认您的站点支持发送mail，所包含密码为系统随机生成</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(23)" value=0 <%if Dvbbs.forum_setting(23)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(23)" value=1 <%if Dvbbs.forum_setting(23)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>一个Email只能注册一个帐号</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(24)" value=0 <%if Dvbbs.forum_setting(24)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(24)" value=1 <%if Dvbbs.forum_setting(24)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>注册需要管理员认证</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(25)" value=0 <%if Dvbbs.forum_setting(25)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(25)" value=1 <%if Dvbbs.forum_setting(25)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发送注册信息邮件</U><BR>请确认您打开了邮件功能</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(47)" value=0 <%if Dvbbs.forum_setting(47)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(47)" value=1 <%if Dvbbs.forum_setting(47)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>开启短信欢迎新注册用户</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(46)" value=0 <%if Dvbbs.forum_setting(46)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(46)" value=1 <%if Dvbbs.forum_setting(46)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>

</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting10"></a><b>系统设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>论坛所在时区</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_info(9)" size="35" value="<%=Dvbbs.Forum_info(9)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>服务器时差</U></td>
<td width="50%" class=Forumrow>  
<select name="forum_setting(0)">
<%for i=-23 to 23%>
<option value="<%=i%>" <%if i=CInt(Dvbbs.forum_setting(0)) then%>selected<%end if%>><%=i%>
<%next%>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>脚本超时时间</U><BR>默认为300，一般不做更改</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(1)" size="3" value="<%=Dvbbs.forum_setting(1)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否显示页面执行时间</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(30)" value=0 <%If Dvbbs.forum_setting(30)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(30)" value=1 <%if Dvbbs.forum_setting(30)="1" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>禁止的邮件地址</U><BR>在下面指定的邮件地址将被禁止注册，每个邮件地址用“|”符号分隔<BR>本功能支持模糊搜索，如设置了eway禁止，将禁止eway@aspsky.net或者eway@dvbbs.net类似这样的注册</td>
<td width="50%" class=Forumrow> 
<input type="text" name="forum_setting(52)" size="50" value="<%=Dvbbs.forum_setting(52)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>论坛脚本过滤扩展设置</U><BR>此设置为开启HTML解释的时候对脚本代码的识别设置，<br>您可以根据需要添加自定的过滤<br>格式是：过滤字| 如：abc|efg| 这样就添加了abc和efg的过滤</td>
<td width="50%" class=Forumrow> 
<Input type="text" name="forum_setting(77)" size="50" value="<%=Dvbbs.forum_setting(77)%>"><br> 没有添加可以填0,如果添加了最后一个字符必须是"|"
</td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting12"></a><b>在线和用户来源</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线显示用户IP</U><BR>关闭后如果所属用户组、论坛权限、用户权限中设置了用户可浏览则可见</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(28)" value=0 <%if Dvbbs.forum_setting(28)="0" then%>checked<%end if%>>保密&nbsp;
<input type=radio name="forum_setting(28)" value=1 <%if Dvbbs.forum_setting(28)="1" then%>checked<%end if%>>公开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线显示用户来源</U><BR>关闭后如果所属用户组、论坛权限、用户权限中设置了用户可浏览则可见<BR>开启本功能较消耗资源</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(36)" value=0 <%if Dvbbs.forum_setting(36)="0" then%>checked<%end if%>>保密&nbsp;
<input type=radio name="forum_setting(36)" value=1 <%if Dvbbs.forum_setting(36)="1" then%>checked<%end if%>>公开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线资料列表显示用户当前位置</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(33)" value=0 <%if Dvbbs.forum_setting(33)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(33)" value=1 <%if Dvbbs.forum_setting(33)="1" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线资料列表显示用户登录和活动时间</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(34)" value=0 <%if Dvbbs.forum_setting(34)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(34)" value=1 <%if Dvbbs.forum_setting(34)="1" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线资料列表显示用户浏览器和操作系统</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(35)" value=0 <%If Dvbbs.forum_setting(35)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(35)" value=1 <%if Dvbbs.forum_setting(35)="1" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线名单显示客人在线</U><BR>为节省资源建议关闭</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(15)" value=0 <%if Dvbbs.forum_setting(15)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(15)" value=1 <%if Dvbbs.forum_setting(15)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>在线名单显示用户在线</U><BR>为节省资源建议关闭</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(14)" value=0 <%if Dvbbs.forum_setting(14)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(14)" value=1 <%if Dvbbs.forum_setting(14)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>删除不活动用户时间</U><BR>可设置删除多少分钟内不活动用户<BR>单位：分钟，请输入数字</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(8)" size="3" value="<%=Dvbbs.forum_setting(8)%>">&nbsp;分钟
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>总论坛允许同时在线数</U><BR>如不想限制，可设置为0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(26)" size="6" value="<%=Dvbbs.forum_setting(26)%>">&nbsp;人
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>展开用户在线列表每页显示用户数</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(58)" size="6" value="<%=Dvbbs.forum_setting(58)%>">&nbsp;人
</td>
</tr>

</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><a name="setting13"></a><b>邮件选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>发送邮件组件</U>
<input type="hidden" id="forum_emailplus" value="<b>发送邮件组件帮助:</b><br><li>选择组件时请确认服务器是否支持。">
<BR>如果您的服务器不支持下列组件，请选择不支持</td>
<td width="43%" class=Forumrow>  
<select name="forum_setting(2)" onChange="chkselect(options[selectedIndex].value,'know1');">
<option value="0" <%if Dvbbs.forum_setting(2)=0 then%>selected<%end if%>>不支持 
<option value="1" <%if Dvbbs.forum_setting(2)=1 then%>selected<%end if%>>JMAIL 
<option value="2" <%if Dvbbs.forum_setting(2)=2 then%>selected<%end if%>>CDONTS 
<option value="3" <%if Dvbbs.forum_setting(2)=3 then%>selected<%end if%>>ASPEMAIL 
</select><div id=know1></div></td>
<td class=Forumrow><a href=# onclick="helpscript(forum_emailplus);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumRowHighlight> <U>SMTP Server地址</U>
<input type="hidden" id="forum_smtp" value="<b>SMTP Server地址帮助:</b><br><li>当选择了邮件组件时此项建议填写，例如：smtp.21cn.com；<li>此邮件服务器地址的填写是根据管理员邮箱而定，例如管理员邮箱为abc@163.net，则此栏可填：smtp.163.net。">
<BR>只有在论坛使用设置中打开了发送邮件功能，该填写内容方有效</td>
<td class=ForumRowHighlight>  
<input type="text" name="Forum_info(4)" size="35" value="<%=Dvbbs.Forum_info(4)%>">
</td>
<td class=forumRowHighlight><a href=# onclick="helpscript(forum_smtp);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=Forumrow> <U>邮件登录用户名</U><BR>只有在论坛使用设置中打开了发送邮件功能，该填写内容方有效</td>
<td colspan=2 class=Forumrow>
<input type="text" name="Forum_info(12)" size="35" value="<%=Dvbbs.Forum_info(12)%>">
</td></tr>
<tr> 
<td class=ForumRowHighlight> <U>邮件登录密码</U></td>
<td colspan=2 class=ForumRowHighlight>  
<input type="password" name="Forum_info(13)" size="35" value="<%=Dvbbs.Forum_info(13)%>">
</td>
</tr>
</table>
<a name="setting14"></a>
<BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><b>上传设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>头像上传</U></td>
<td width="43%" class=Forumrow>
<SELECT name="forum_setting(7)" >
<OPTION value=0 <%if Dvbbs.forum_setting(7)=0 then%>selected<%end if%>>完全关闭&nbsp;
<OPTION value=1 <%if Dvbbs.forum_setting(7)=1 then%>selected<%end if%>>完全打开&nbsp;
<OPTION value=2 <%if Dvbbs.forum_setting(7)=2 then%>selected<%end if%>>只允许会员上传&nbsp;
</SELECT>
</td>
<input type="hidden" id="Forum_FaceUpload" value="<b>头像上传帮助:</b><br><li>当开启此功能，用户可以把图像文件上传到服务器作为头像。<li>在上传管理中有对上传头像进行管理。<LI>完全关闭：即注册和修改资料都不允许上传头像。<LI>完全打开：即注册和修改资料都允许上传头像。<LI>只允许会员上传：即会员修改个人资料时允许上传头像。">
<td class=Forumrow><a href=# onclick="helpscript(Forum_FaceUpload);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumRowHighlight><U>允许的最大头像文件大小</U></td>
<td class=ForumRowHighlight> 
<input type="text" name="forum_setting(56)" size="6" value="<%=Dvbbs.forum_setting(56)%>">&nbsp;K
</td>
<input type="hidden" id="Forum_FaceUploadSize" value="<b>头像文件大小帮助:</b><br><li>限制上传头像文件的大小。<li>用户头像除上传限制外，请查看“用户选项”相关设置。">
<td class=ForumRowHighlight><a href=# onclick="helpscript(Forum_FaceUploadSize);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td class=Forumrow ><U>选取上传组件:</U></td>
<td class=Forumrow >
<select name="forum_setting(43)" onChange="chkselect(options[selectedIndex].value,'know2');">
<option value="999" <%if Dvbbs.forum_setting(43)=999 then%>selected<%end if%>>关闭
<option value="0" <%if Dvbbs.forum_setting(43)=0 then%>selected<%end if%>>无组件上传类
<option value="1" <%if Dvbbs.forum_setting(43)=1 then%>selected<%end if%>>Lyfupload组件
<option value="2" <%if Dvbbs.forum_setting(43)=2 then%>selected<%end if%>>Aspupload3.0组件 
<option value="3" <%if Dvbbs.forum_setting(43)=3 then%>selected<%end if%>>SA-FileUp 4.0组件
<option value="4" <%if Dvbbs.forum_setting(43)=4 then%>selected<%end if%>>DvFile-Up V1.0组件
</option></select><div id="know2"></div>
</td>
<td class=Forumrow >
<input type="hidden" id="forum_upload" value="<b>选取上传组件帮助:</b><br><li>当选取时，论坛系统会自动为您检测服务器是否支持该组件；<li>若提示不支持，请选择关闭。">
<a href=# onclick="helpscript(forum_upload);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumRowHighlight><U>选取生成预览图片组件:</U></td>
<td class=ForumRowHighlight> 
<select name="forum_setting(45)" onChange="chkselect(options[selectedIndex].value,'know3');">
<option value="999" <%if Dvbbs.forum_setting(45)=999 then%>selected<%end if%>>关闭
<option value="0" <%if Dvbbs.forum_setting(45)=0 then%>selected<%end if%>>CreatePreviewImage组件
<option value="1" <%if Dvbbs.forum_setting(45)=1 then%>selected<%end if%>>AspJpeg组件
<option value="2" <%if Dvbbs.forum_setting(45)=2 then%>selected<%end if%>>SA-ImgWriter组件
<option value="3" <%if Dvbbs.forum_setting(45)=3 then%>selected<%end if%>>SJCatSoft V2.6组件
</select><div id="know3"></div>
</td>
<td class=forumRowHighlight>
<input type="hidden" id="forum_CreatImg" value="<b>选取生成预览图片组件帮助:</b><br><li>当选取时，论坛系统会自动为您检测服务器是否支持该组件；<li>若提示不支持，请选择关闭。">
<a href=# onclick="helpscript(forum_CreatImg);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumRow><U>上传图片添加水印文字（可为空）:</U></td>
<td class=ForumRow> 
<INPUT TYPE="text" NAME="forum_setting(73)" size=40 value="<%=Dvbbs.Forum_Setting(73)%>">
</td>
<td class=ForumRow>
<input type="hidden" id="forum_CreatText" value="<b>上传图片添加水印文字帮助:</b><br><li>若不需要水印文字效果，请设置为空；<li>水印文字字数不宜超过15个字符,不支持任何WEB编码标记；<li>目前支持的相关图片组件有：AspJpeg组件，SA-ImgWriter V1.21组件，SJCatSoft V2.6组件。">
<a href=# onclick="helpscript(forum_CreatText);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumRow><U>生成预览图片大小设置(宽度|高度):</U></td>
<td class=ForumRow><INPUT TYPE="text" NAME="forum_Setting(72)" size=10 value="<%=Dvbbs.Forum_Setting(72)%>"> 像素</td>
<td class=ForumRow>
<input type="hidden" id="forum_CreatImgSize" value="<b>生成预览图片大小设置帮助:</b><br><li>当选取了生成预览图片组件，并且服务器上装有相应组件，此功能才能生效；<li>生成图像大小设置的格式为：宽度|高度，宽度与高度之间用“|”分隔；">
<a href=# onclick="helpscript(forum_CreatImgSize);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<%
If IsObjInstalled("Scripting.FileSystemObject") Then 
%>
<tr> 
<td class=ForumRow><U>是否采用文件、图片防盗链</U></td>
<td class=ForumRow>
<input type=radio name="Forum_Setting(75)" value=0 <%if Dvbbs.Forum_Setting(75)=0 Then %>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(75)" value=1 <%if Dvbbs.Forum_Setting(75)=1 Then %>checked<%end if%>>打开&nbsp;

</td>
<td class=ForumRow>
</td>
</tr>
<tr> 
<td class=ForumRow><U>上传目录设定</U></td>
<td class=ForumRow>
<%
If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
%>
<input type=text name="Forum_Setting(76)" value=<%=Dvbbs.Forum_Setting(76)%>>如果修改了此项，请用FTP手工创建目录和移动原有上传文件。
</td>
<td class=ForumRow>
</td>
</tr>

<%
End If 
%>
</table>
<BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting15"></a><b>用户选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>允许个人签名</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(42)" value=0 <%if Dvbbs.forum_setting(42)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(42)" value=1 <%if Dvbbs.forum_setting(42)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>允许用户使用头像</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(53)" value=0 <%if Dvbbs.forum_setting(53)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(53)" value=1 <%if Dvbbs.forum_setting(53)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>最大头像尺寸</U><BR>定义内容为头像的最大高度和宽度</td>
<td class=Forumrow width="50%"> 
<input type="text" name="forum_setting(57)" size="6" value="<%=Dvbbs.forum_setting(57)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>默认头像宽度</U><BR>定义内容为论坛头像的默认宽度</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(38)" size="6" value="<%=Dvbbs.forum_setting(38)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>默认头像高度</U><BR>定义内容为论坛头像的默认宽度</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(39)" size="6" value="<%=Dvbbs.forum_setting(39)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>使用自定义头像的最少发帖数</U></td>
<td class=Forumrow width="50%"> 
<input type="text" name="forum_setting(54)" size="6" value="<%=Dvbbs.forum_setting(54)%>">&nbsp;篇
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>允许从其他站点链接头像</U><BR>就是是否可以直接使用http..这样的url来直接显示头像</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(55)" value=0 <%if Dvbbs.forum_setting(55)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(55)" value=1 <%if Dvbbs.forum_setting(55)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户签名是否开启UBB代码</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(65)" value=0 <%if Dvbbs.forum_setting(65)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(65)" value=1 <%if Dvbbs.forum_setting(65)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户签名是否开启HTML代码</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(66)" value=0 <%if Dvbbs.forum_setting(66)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(66)" value=1 <%if Dvbbs.forum_setting(66)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户是否开启贴图标签</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(67)" value=0 <%if Dvbbs.forum_setting(67)=0 then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(67)" value=1 <%if Dvbbs.forum_setting(67)=1 then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户是否开启Flash标签</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(71)" value=0 <%if Dvbbs.forum_setting(71)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(71)" value=1 <%if Dvbbs.forum_setting(71)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户头衔</U><BR>是否允许用户自定义头衔</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(6)" value=0 <%if Dvbbs.forum_setting(6)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(6)" value=1 <%if Dvbbs.forum_setting(6)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>用户头衔最大长度</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(59)" size="6" value="<%=Dvbbs.forum_setting(59)%>">&nbsp;byte
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>自定义头衔最少发帖数量限制</U><BR>不做限制请设置为0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(60)" size="6" value="<%=Dvbbs.forum_setting(60)%>">&nbsp;篇
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>自定义头衔注册天数限制</U><BR>不做限制请设置为0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(61)" size="6" value="<%=Dvbbs.forum_setting(61)%>">&nbsp;天
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>自定义头衔上面两个条件加在一起限制</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(62)" value=0 <%if Dvbbs.forum_setting(62)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(62)" value=1 <%if Dvbbs.forum_setting(62)="1" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>自定义头衔中要屏蔽的词语</U><BR>每个限制字符用“|”符号隔开</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(63)" size="50" value="<%=Dvbbs.forum_setting(63)%>">
</td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting17"></a><b>防刷新机制</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>防刷新机制</U><BR>如选择打开请填写下面的限制刷新时间<BR>对版主和管理员无效</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(19)" value=0 <%if Dvbbs.forum_setting(19)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="forum_setting(19)" value=1 <%if Dvbbs.forum_setting(19)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>浏览刷新时间间隔</U><BR>填写该项目请确认您打开了防刷新机制<BR>仅对帖子列表和显示帖子页面起作用</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(20)" size="3" value="<%=Dvbbs.forum_setting(20)%>">&nbsp;秒
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>防刷新功能有效的页面</U><BR>请确认您打开了防刷新功能<BR>您指定的页面将有防刷新作用，用户在限定的时间内不能重复打开该页面，具有一定减少资源消耗的作用<BR>每个页面名请用“|”符号隔开</td>
<td width="50%" class=Forumrow> 
<input type="text" name="forum_setting(64)" size="50" value="<%=Dvbbs.forum_setting(64)%>">
</td>
</tr>

</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><a name="setting20"></a><b>搜索选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>每次搜索时间间隔</U></td>
<td class=Forumrow width="43%"> 
<input type="text" name="Forum_Setting(3)" size="6" value="<%=Dvbbs.Forum_Setting(3)%>">&nbsp;秒
</td>
<input type="hidden" id="s_1" value="<b>每次搜索时间间隔</b><br><li>设置合理的每次搜索时间间隔，可以避免用户反复进行相同搜索而消耗大量论坛资源">
<td class=forumRow><a href=# onclick="helpscript(s_1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumrowHighLight><U>搜索字串最小和最大长度</U><BR>最小和最大字符请用符号“|”分隔，单位为字节<BR>最小字符不宜设置过小，最大字符不宜设置过大，建议用默认值</td>
<td class=ForumrowHighLight > 
<input type="text" name="Forum_Setting(4)" size="8" value="<%=Dvbbs.Forum_Setting(4)%>">
</td>
<input type="hidden" id="s_2" value="<b>搜索字串最小和最大长度</b><br><li>最小和最大字符请用符号“|”分隔，单位为字节<br><li>最小字符不宜设置过小，最大字符不宜设置过大，设置过小或者过大都将消耗大量论坛资源">
<td class=ForumrowHighLight><a href=# onclick="helpscript(s_2);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=Forumrow ><U>搜索可以不受字串长度限制的词</U><BR>每个字符请用符号“|”分隔</td>
<td class=Forumrow> 
<input type="text" name="Forum_Setting(9)" size="50" value="<%=Dvbbs.Forum_Setting(9)%>">&nbsp;
</td>
<input type="hidden" id="s_3" value="<b>搜索可以不受字串长度限制的词</b><br><li>每个字符请用符号“|”分隔<br><li>合理的填写不受字串长度限制的词，可以使一些常用且简单的单词搜索到结果，但您同时必须考虑搜索字串长度的长短是和消耗的资源成正比的">
<td class=Forumrow><a href=# onclick="helpscript(s_3);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumrowHighLight><U>搜索返回最多的结果数</U><BR>建议不要设置过大</td>
<td class=ForumrowHighLight> 
<input type="text" name="Forum_Setting(12)" size="6" value="<%=Dvbbs.Forum_Setting(12)%>">&nbsp;个
</td>
<input type="hidden" id="s_4" value="<b>搜索返回最多的结果数</b><br><li>单位为数字<br><li>返回搜索的结果数和消耗的资源成正比，请合理设置">
<td class=ForumrowHighLight><a href=# onclick="helpscript(s_4);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=Forumrow>
<U>搜索热门帖子条件中对应的搜索天数和浏览次数标准</U><BR>搜索天数和浏览次数请用符号“|”分隔，单位为数字<BR>搜索天数不宜设置过大，建议用默认值</td>
<td class=Forumrow> 
<input type="text" name="Forum_Setting(13)" size="8" value="<%=Dvbbs.Forum_Setting(13)%>">
</td>
<input type="hidden" id="s_5" value="<b>搜索热门帖子条件中对应的搜索天数和浏览次数标准</b><br><li>搜索天数和浏览次数请用符号“|”分隔，单位为数字<br><li>作为热门主题的搜索天数和浏览次数标准和论坛资源消耗成正比，请合理设置">
<td class=Forumrow><a href=# onclick="helpscript(s_5);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumrowHighLight> <U>是否开启全文搜索</U><BR>ACCESS数据库不建议开启<BR>SQL数据库做了全文索引可以开启</td>
<td class=ForumrowHighLight>  
<input type=radio name="Forum_Setting(16)" value=0 <%If Dvbbs.Forum_Setting(16)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(16)" value=1 <%If Dvbbs.Forum_Setting(16)="1" then%>checked<%end if%>>打开&nbsp;
</td>
<input type="hidden" id="s_6" value="<b>是否开启全文搜索</b><br><li>ACCESS数据库在数据容量较大情况下开启搜索将消耗大量资源，SQL数据库开启数据库全文搜索后可使用本选项<br><li>设置SQL数据库的全文搜索请看微软相关帮助文档">
<td class=ForumrowHighLight><a href=# onclick="helpscript(s_6);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=Forumrow> <U>用户列表允许用户名搜索</U></td>
<td class=Forumrow>  
<input type=radio name="Forum_Setting(17)" value=0 <%if Dvbbs.Forum_Setting(17)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(17)" value=1 <%if Dvbbs.Forum_Setting(17)="1" then%>checked<%end if%>>打开&nbsp;
</td>
<input type="hidden" id="s_7" value="<b>用户列表允许用户名搜索</b><br><li>开启本项目，在用户列表中可以对用户名做简单搜索<br><li>出于用户数据安全上的考虑，您也可以关闭该选项">
<td class=Forumrow><a href=# onclick="helpscript(s_7);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumrowHighLight> <U>用户列表允许列出管理团队</U></td>
<td class=ForumrowHighLight>  
<input type=radio name="Forum_Setting(18)" value=0 <%if Dvbbs.Forum_Setting(18)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(18)" value=1 <%if Dvbbs.Forum_Setting(18)="1" then%>checked<%end if%>>打开&nbsp;
</td>
<input type="hidden" id="s_8" value="<b>用户列表允许列出管理团队</b><br><li>开启本项目，在用户列表中可以列出论坛中的管理团队资料，即版主或其以上等级的用户<br><li>出于用户数据安全上的考虑，您也可以关闭该选项">
<td class=ForumrowHighLight><a href=# onclick="helpscript(s_8);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=Forumrow> <U>用户列表允许列出所有用户</U></td>
<td class=Forumrow>  
<input type=radio name="Forum_Setting(27)" value=0 <%if Dvbbs.Forum_Setting(27)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(27)" value=1 <%if Dvbbs.Forum_Setting(27)="1" then%>checked<%end if%>>打开&nbsp;
</td>
<input type="hidden" id="s_9" value="<b>用户列表允许列出所有用户</b><br><li>开启本项目，在用户列表中可以列出论坛中的所有的用户资料<br><li>出于用户数据安全上的考虑，您也可以关闭该选项">
<td class=Forumrow><a href=# onclick="helpscript(s_9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=ForumrowHighLight> <U>用户列表允许列出TOP排行用户</U></td>
<td class=ForumrowHighLight>  
<input type=radio name="Forum_Setting(31)" value=0 <%if Dvbbs.Forum_Setting(31)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_Setting(31)" value=1 <%if Dvbbs.Forum_Setting(31)="1" then%>checked<%end if%>>打开&nbsp;
</td>
<input type="hidden" id="s_10" value="<b>用户列表允许列出TOP排行用户</b><br><li>开启本项目，在用户列表中可以列出论坛按照发贴和积分数等用户排行<br><li>出于用户数据安全上的考虑，您也可以关闭该选项">
<td class=ForumrowHighLight><a href=# onclick="helpscript(s_10);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td class=Forumrow><U>用户列表TOP个数</U></td>
<td class=Forumrow> 
<input type="text" name="forum_setting(68)" size="6" value="<%=Dvbbs.forum_setting(68)%>">&nbsp;个
</td>
<input type="hidden" id="s_11" value="<b>用户列表TOP个数</b><br><li>在开启了TOP排行的情况下，将根据这里所设置的数字读取所规定数目的用户数据<br><li>出于用户数据安全上的考虑和出于论坛资源消耗方面的考虑，您也可以减少该选项的设置数目">
<td class=Forumrow><a href=# onclick="helpscript(s_11);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting18"></a><b>论坛分页设置</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td class=Forumrow  width="50%"> <U>每页显示最多纪录</U><BR>用于论坛所有和分页有关的项目（帖子列表和浏览帖子除外）</td>
<td class=Forumrow  width="50%">  
<input type="text" name="forum_setting(11)" size="3" value="<%=Dvbbs.forum_setting(11)%>">&nbsp;条
</td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting16"></a><b>帖子选项</b>[<a href="#top">顶部</a>]</td>
</tr>
<tr> 
<td class=Forumrow  width="50%"> <U>作为热门话题的最低人气值</U><BR>标准为主题回复数</td>
<td class=Forumrow  width="50%">  
<input type="text" name="forum_setting(44)" size="3" value="<%=Dvbbs.forum_setting(44)%>">&nbsp;条
</td>
</tr>
<tr> 
<td class=Forumrow> <U>编辑过的帖子显示"由xxx于yyy编辑"的信息</U></td>
<td class=Forumrow>  
<input type=radio name="forum_setting(48)" value=0 <%if Dvbbs.forum_setting(48)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(48)" value=1 <%if Dvbbs.forum_setting(48)="1" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow> <U>管理员编辑后显示"由XXX编辑"的信息</U></td>
<td class=Forumrow>  
<input type=radio name="forum_setting(49)" value=0 <%if Dvbbs.forum_setting(49)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(49)" value=1 <%if Dvbbs.forum_setting(49)="1" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow> <U>等待"由XXX编辑"信息显示的时间</U><BR>允许用户编辑自己的帖子而不在帖子底部显示"由XXX编辑"信息的时限(以分钟为单位)</td>
<td class=Forumrow>  
<input type="text" name="forum_setting(50)" size="3" value="<%=Dvbbs.forum_setting(50)%>">&nbsp;分钟
</td>
</tr>
<tr> 
<td class=Forumrow> <U>编辑帖子时限</U><BR>编辑处理帖子的时间限制(以分钟为单位, 1天是1440分钟) 超过这个时间限制, 只有管理员和版主才能编辑和删除帖子. 如果不想使用这项功能, 请设置为0</td>
<td class=Forumrow>  
<input type="text" name="forum_setting(51)" size="3" value="<%=Dvbbs.forum_setting(51)%>">&nbsp;分钟
</td>
</tr>
</table>
<BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="settingxu"></a><b>虚拟形象选项</b>[<a href="#top">顶部</a>]
</tr>
<tr> 
<td class=Forumrow  width="50%"> <U>是否启用虚拟形象</U></td>
<td class=Forumrow>  
<input type=radio name="forum_setting(82)" value=1 <%if Dvbbs.forum_setting(82)="1" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(82)" value=0 <%if Dvbbs.forum_setting(82)="0" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
</table>
<BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting19"></a><b>门派设置</b>[<a href="#top">顶部</a>]
</tr>
<tr> 
<td class=Forumrow  width="50%"> <U>是否开启论坛门派</U></td>
<td class=Forumrow>  
<input type=radio name="forum_setting(32)" value=0 <%if Dvbbs.forum_setting(32)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="forum_setting(32)" value=1 <%if Dvbbs.forum_setting(32)="1" then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> &nbsp;</td>
<td width="50%" class=Forumrow>  
<div align="center"> 
<input type="submit" name="Submit" value="提 交">
</div>
</td>
</tr>
</table>
</form>
<div id="Issubport0" style="display:none">请选择EMAIL组件！</div>
<div id="Issubport999" style="display:none"></div>
<%
Dim InstalledObjects(12)
InstalledObjects(1) = "JMail.Message"				'JMail 4.3
InstalledObjects(2) = "CDONTS.NewMail"				'CDONTS
InstalledObjects(3) = "Persits.MailSender"			'ASPEMAIL
'-----------------------
InstalledObjects(4) = "Scripting.FileSystemObject"	'Fso
InstalledObjects(5) = "LyfUpload.UploadFile"		'LyfUpload
InstalledObjects(6) = "Persits.Upload"				'Aspupload3.0
InstalledObjects(7) = "SoftArtisans.FileUp"			'SA-FileUp 4.0
InstalledObjects(8) = "DvFile.Upload"				'DvFile-Up V1.0
'-----------------------
InstalledObjects(9) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
InstalledObjects(10)	= "Persits.Jpeg"				'AspJpeg
InstalledObjects(11) = "SoftArtisans.ImageGen"		'SoftArtisans ImgWriter V1.21
InstalledObjects(12) = "sjCatSoft.Thumbnail"		'sjCatSoft.Thumbnail V2.6

For i=1 to 12
	Response.Write "<div id=""Issubport"&i&""" style=""display:none"">"
	If IsObjInstalled(InstalledObjects(i)) Then Response.Write "<font color=red><b>√</b>服务器支持!</font>" Else Response.Write "<b>×</b>服务器不支持!" 
	Response.Write "</div>"
Next
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function chkselect(s,divid)
{
var divname='Issubport';
var chkreport;
	s=Number(s)
	if (divid=="know1")
	{
	divname=divname+s;
	}
	if (divid=="know2")
	{
	s+=4;
	if (s==1003){s=999;}
	divname=divname+s;
	}
	if (divid=="know3")
	{
	s+=9;
	if (s==1008){s=999;}
	divname=divname+s;
	}
document.getElementById(divid).innerHTML=divname;
chkreport=document.getElementById(divname).innerHTML;
document.getElementById(divid).innerHTML=chkreport;
}
//-->
</SCRIPT>
<%
end sub

sub saveconst()
Dim Forum_copyright,Forum_info,forum_setting,iforum_setting,isetting
Dim Forum_Maxonline,Forum_TopicNum,Forum_PostNum
Dim Forum_UserNum,Forum_MaxPostNum,Forum_MaxPostDate,Forum_MaxonlineDate
Dim Forum_pack

If not IsDate(Request.Form("Forum_Setting(74)")) Then 
	Errmsg=ErrMsg + "<li>论坛创建日期必须是一个有效日期。"
	Dvbbs_error()
	Exit Sub
End If

If not IsDate(Request.Form("Forum_MaxPostDate")) Then 
	Errmsg=ErrMsg + "<li>论坛最高日发贴发生时间日期必须是一个有效日期。"
	Dvbbs_error()
	Exit Sub
Else
	Forum_MaxPostDate=Request.Form("Forum_MaxPostDate")
End If

If not IsDate(Request.Form("Forum_MaxonlineDate")) Then 
	Errmsg=ErrMsg + "<li>历史最高同时在线纪录发生时间日期必须是一个有效日期。"
	Dvbbs_error()
	Exit Sub
Else
	Forum_MaxonlineDate=Request.Form("Forum_MaxonlineDate")
End If

Forum_Maxonline	= Request.Form("Forum_Maxonline")
Forum_TopicNum	= Request.Form("Forum_TopicNum")
Forum_PostNum	= Request.Form("Forum_PostNum")
Forum_UserNum	= Request.Form("Forum_UserNum")
Forum_MaxPostNum= Request.Form("Forum_MaxPostNum")
Forum_pack	= Request.Form("Forum_pack(0)")&"|||"&Trim(Request.Form("Forum_pack(1)"))

If Not ISNumeric(Forum_Maxonline&Forum_TopicNum&Forum_PostNum&Forum_UserNum&Forum_MaxPostNum) Then 
	Errmsg=ErrMsg + "<li>非法的参数，论坛系统数据出错，提交中止。"
	Dvbbs_error()
	Exit Sub
End If

If not isnumeric(request.Form("cid")) or not isnumeric(request.Form("Sid")) Then
	Errmsg=ErrMsg + "<li>请选择模板与风格！"
	Dvbbs_error()
	Exit Sub
End IF

Dim setingdata,j
If Forum_Maxonline="" Then Forum_Maxonline=0
If Forum_TopicNum="" Then Forum_TopicNum=0
If Forum_PostNum="" Then Forum_PostNum=0
If Forum_UserNum="" Then Forum_UserNum=0
If Forum_MaxPostNum="" Then Forum_MaxPostNum=0
For i = 0 To 100
	If Trim(request.Form("Forum_Setting("&i&")"))=""  Or i=70 Then
		'Response.Write "Forum_Setting("&i&")<br>"
		isetting=0
		If i=70 Then
			isetting=""
			For j=0 to  23
				If isetting="" Then
					If Request.form("Forum_Setting(70)"&j)="1" Then
						isetting="1"
					Else
						isetting="0"
					End If
				Else
					If Request.form("Forum_Setting(70)"&j)="1" Then
						isetting=isetting&"|1"
					Else
						isetting=isetting&"|0"
					End If
				End If
			Next
		End If
	Else
		isetting=Replace(Trim(request.Form("Forum_Setting("&i&")")),",","")
	End If
	If i = 0 Then
		forum_setting = isetting
	Else
		forum_setting = forum_setting & "," & isetting
	End If
Next

For i = 0 To 13
	If Trim(Request.Form("Forum_info("&i&")")) = "" And i <> 4 And i <> 12 And i<>13 Then
		'Response.Write "Forum_info("&i&")<br>"
		isetting=0
	Else
		isetting=Replace(Trim(request.Form("Forum_info("&i&")")),",","")
	End If
	If i = 0 Then
		Forum_info = isetting
	Else
		Forum_info = Forum_info & "," & isetting
	End If
Next
'response.write Forum_info
'response.write "<br>"
'Response.Write Dvbbs.Forum_Setting
'Response.End
Forum_copyright=request("copyright")

'forum_info|||forum_setting|||forum_user|||copyright|||splitword|||stopreadme
Set rs=Dvbbs.execute("select forum_setting from dv_setup")
iforum_setting=split(rs(0),"|||")
forum_setting=forum_info & "|||" & forum_setting & "|||" & iforum_setting(2) & "|||" & Forum_copyright & "|||" & iforum_setting(4) & "|||" & request.Form("StopReadme")
forum_setting=Replace(forum_setting,"'","''")

'Response.Write forum_setting
'response.end

sql="update Dv_setup set Forum_Setting='"&forum_setting&"',forum_sid="&request.Form("Sid")&",Forum_cid="&request.Form("cid")
sql=sql+",Forum_Maxonline="&Forum_Maxonline&",Forum_TopicNum="&Forum_TopicNum&",Forum_PostNum="&Forum_PostNum &",Forum_UserNum="&Forum_UserNum&",Forum_MaxPostNum="&Forum_MaxPostNum&",Forum_MaxPostDate='"&Forum_MaxPostDate &"',Forum_MaxonlineDate='"&Forum_MaxonlineDate&"',Forum_pack='"&Forum_pack&"'"
dvbbs.execute(sql)
Dvbbs.Name="setup"
dvbbs.ReloadSetup
Dv_suc("设置论坛常规信息成功")
end sub

'恢复默认设置
Sub restore()
	Dim Forum_setting
	forum_setting="动网先锋论坛,http://bbs.dvbbs.net,动网先锋,http://www.aspsky.net/,,eway@aspsky.net,images/logo.gif,http://www.aspsky.cn/email.asp,aspsky|dvbbs|动网|动网论坛|asp|论坛|插件,北京时间,动网论坛是使用量最多、覆盖面最广的免费中文论坛，也是国内知名的技术讨论站点，希望我们辛苦的努力可以为您带来很多方便,index.asp,0,0|||0,300,0,60,2|20,1,1,1,20,dvbbs|sql|aspsky|asp|php|cgi|jsp|htm,0,20,500,20|200,1,1,1,1,1,0,3,0,40,0,0,0,0,1,1,0,1,1,1,1,1,1,0,1,32,32,0,10,1,0,10,999,1,0,1,1,0,0,0,1,0,1,200,120,60,9,15,4,0,0,list.asp,1,0,1,20,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,0,100|100,http://BBs.Dvbbs.Net 动网先锋,2000-3-26,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0|||1000,5,2,7,1,200,12,1,10,10,30,3,2,5,1,10,5,10|||Copyright &copy;2002 - 2005  <a href=""http://www.aspsky.net""><font face=Verdana, Arial, Helvetica, sans-serif><b>Aspsky<font color=#CC0000>.Net</font></b></font></a>|||!,@,#,$,%,^,&,*,(,),{,},[,],|,\,.,/,?,`,~|||论坛暂停使用"
	Conn.Execute("update Dv_setup set Forum_Setting='"&forum_setting&"'")
	Dv_suc("还原论坛常规设置成功")
	Dvbbs.Name="setup"
	dvbbs.ReloadSetup
End Sub

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>