<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
	Dim Board_Setting
	if not Dvbbs.master or instr(","&session("flag")&",",",9,")=0 then
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		if request("action")="save" then
		call saveconst()
		else
		call consted()
		end if
		Footer()
	end if

sub consted()
if not isnumeric(request("editid")) then
	Errmsg=ErrMsg + "<BR><li>错误的版面信息"
	dvbbs_error()
	exit sub
end if
set rs=Dvbbs.Execute("select * from dv_board where boardid="&request("editid"))
Board_Setting=split(rs("board_setting"),",")
%>
<table width="95%" cellspacing="1" cellpadding="1"  align=center class="tableBorder">
<tr><th height="25" colspan="5" align=left>论坛高级设置 → <%=rs("boardtype")%></th></tr>
<tr> 
<td width="100%" class=Forumrow colspan=5 height=25>
说明：<BR>
1、请仔细设置下面的高级选项，Flash标签如果打开，对安全有一定影响，请根据您的具体情况考虑。<BR>
2、您可以将高级设置的某项设置（选择该行设置右边的复选框）保存到所有版面、相同分类下所有版面（不包括分类）、相同分类下所有版面（包括分类）、同分类同级别版面，该项设置请慎重操作。<BR>
3、<font color=red>注意，选择批量更新包括主题将会使用相同设置</font>。
</td>
</tr>
<form method="POST" action="admin_boardsetting.asp?action=save">
<input type=hidden value="<%=request("editid")%>" name="editid">
<tr> 
<td width="100%" class=ForumrowHighlight colspan=5 height=25>
<font color=blue>保存目标</font>：<input type=radio name="savetype" value=0 checked>该版面&nbsp;<input type=radio name="savetype" value=1>所有版面&nbsp;<input type=radio name="savetype" value=2>相同分类下所有版面（不包括分类）&nbsp;<input type=radio name="savetype" value=3>相同分类下所有版面（包括分类）&nbsp;<input type=radio name="savetype" value=4>同分类同级别版面
</td>
</tr>
<tr> 
<td width="100%" class=Forumrow colspan=5 height=25>
<font color=blue>
这里指的分类仅指一级分类，而不是该版面的上级版面</font>，比如您目前设置的是一个五级版面，选择了相同分类下所有版面都更新，那么这里将更新包括该分类的一级、二级、三级、四级所有版面，如果您担心更新范围太大，可以选择更新同分类同级别版面。
</td>
</tr>
<tr><th height="25" colspan="5" align=left> &nbsp;功能设置导航</th></tr>
<tr> 
<td width="100%" class=Forumrow colspan=5 height=25>
[<a href="#setting1">基本属性</a>]
[<a href="#setting2">访问权限</a>]
[<a href="#setting3">前台管理权限</a>]
[<a href="#setting4">发贴相关</a>]
[<a href="#setting5">帖子列表显示</a>]
[<a href="#setting6">帖子内容显示</a>]
[<a href="#setting7">附件限制设置</a>]
[<a href="#setting8">论坛专题设置</a>]
[<a href="#setting9">论坛虚拟形象设置</a>]
</td>
</tr>

<tr><th height="25" colspan="5" id=tabletitlelink align=left> &nbsp;<a name="setting1">基本属性</a>[<a href="#top">顶部</a>]</th></tr>
<tr> 
<td width="50%" colspan=2 class=Forumrow>
<U>外部连接</U><BR>填写本内容后，在论坛列表点击此版面将自动切换到该网址<BR>请填写URL绝对路径</td>
<td colspan=2 class=Forumrow>
<input type=text name="Board_Setting(50)" value="<%=Board_Setting(50)%>" size=50>
</td>
<input type="hidden" id="b0" value="<b>外部连接</b><br><li>填写本内容后，在论坛列表点击此版面将自动切换到该网址<br><li>请填写URL绝对路径">
<td class=Forumrow><a href=# onclick="helpscript(b0);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td width="50%" colspan=2 class=ForumrowHighlight>
<U>分论坛LOGO</U><BR>填写图片的相对或绝对路径，不填写则当前版面LOGO为论坛设置中LOGO</td>
<td colspan=2 class=ForumrowHighlight>
<input type=text name="Board_Setting(51)" value="<%=Board_Setting(51)%>" size=50>
</td>
<input type="hidden" id="ba1" value="<b>分论坛LOGO</b><br><li>填写图片的相对或绝对路径，不填写则当前版面LOGO为论坛设置中LOGO">
<td class=ForumrowHighlight><a href=# onclick="helpscript(ba1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>是否采用版主继承制度</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(40)" value=0 <%if Board_Setting(40)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(40)" value=1 <%if Board_Setting(40)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<input type="hidden" id="b6" value="<b>是否采用版主继承制度</b><br><li>如果采用该制度，则上级论坛版主可管理下级论坛相关信息">
<td class=Forumrow><a href=# onclick="helpscript(b6);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>论坛列表显示下属论坛风格</U><BR></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(39)" value=0 <%if Board_Setting(39)="0" then%>checked<%end if%>>列表&nbsp;
<input type=radio name="Board_Setting(39)" value=1 <%if Board_Setting(39)="1" then%>checked<%end if%>>简洁&nbsp;
</td>
<input type="hidden" id="b7" value="<b>论坛列表显示下属论坛风格</b><br><li>当该论坛有下属论坛的时候生效">
<td class=ForumrowHighlight><a href=# onclick="helpscript(b7);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>论坛列表简洁风格一行版面数</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(41)" value="<%=Board_Setting(41)%>"> 个
</td>
<input type="hidden" id="b8" value="<b>论坛列表简洁风格一行版面数</b><br><li>当论坛列表开启了下属论坛风格为简洁，此选项有效，此选项为设置简洁论坛列表风格一行排列版面数">
<td class=Forumrow><a href=# onclick="helpscript(b8);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>是否公开论坛事件中的操作者</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(36)" value=0 <%if Board_Setting(36)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(36)" value=1 <%if Board_Setting(36)="1" then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="b12" value="<b>是否公开论坛事件中的操作者</b><br><li>论坛中对帖子的删除、固顶、设置精华等操作都是要记录操作者和操作内容的，管理员默认可看到这些操作内容，一般用户如果打开了此选项，他们将能看到操作者">
<td class=Forumrow><a href=# onclick="helpscript(b12);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting2">访问权限相关</a>[<a href="#top">顶部</a>]</th></tr>
<tr> 
<td width="50%" colspan=2 class=Forumrow>
<U>本论坛作为分类论坛不允许发贴</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(43)" value=0 <%if Board_Setting(43)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(43)" value=1 <%if Board_Setting(43)="1" then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="b1" value="<b>本论坛作为分类论坛不允许发贴</b><br><li>如果已经有贴则显示或者您可以转移到别的论坛<br><li>选择了该项后所有会员均不能在本版发贴/回帖等操作">
<td class=Forumrow><a href=# onclick="helpscript(b1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>是否锁定论坛</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(0)" value=0 <%if Board_Setting(0)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(0)" value=1 <%If Board_Setting(0)="1" then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="b2" value="<b>是否锁定论坛</b><br><li>锁定论坛只有管理员和该版面版主可进">
<td class=ForumrowHighlight><a href=# onclick="helpscript(b2);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>是否隐藏论坛</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(1)" value=0 <%If Board_Setting(1)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(1)" value=1 <%if Board_Setting(1)="1" then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="b3" value="<b>是否隐藏论坛</b><br><li>隐藏论坛只有管理员和该版面版主可见和进入<br><li>如果用户组或论坛权限管理或用户权限管理中允许则用户可见和进入<br><li>本限制对一级论坛不生效">
<td class=Forumrow><a href=# onclick="helpscript(b3);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>是否认证论坛</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(2)" value=0 <%if Board_Setting(2)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(2)" value=1 <%if Board_Setting(2)="1" then%>checked<%end if%>>是&nbsp;
</td>
<input type="hidden" id="b4" value="<b>是否认证论坛</b><br><li>认证论坛只有管理员和该版面版主可见和进入<br><li>认证论坛对认证用户的添加和管理在版面管理中有连接<br><li>设置了本选项后只有认证用户可进入">
<td class=ForumrowHighlight><a href=# onclick="helpscript(b4);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>帖子审核制度</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(3)" value=0 <%if Board_Setting(3)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(3)" value=1 <%if Board_Setting(3)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<input type="hidden" id="b5" value="<b>帖子审核制度</b><br><li>版主、管理员和开放权限用户可进行审核帖子<br><li>版主、管理员和开放权限用户可直接发贴<br><li>一般用户需审核后帖子方可见">
<td class=Forumrow><a href=# onclick="helpscript(b5);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>扩展审核制度</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(57)" value=0 <%if Board_Setting(57)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(57)" value=1 <%if Board_Setting(57)="1" then%>checked<%end if%>>开放&nbsp;
<input type="hidden" id="bnew" value="<b>扩展帖子审核制度</b><br><li>版主、管理员和开放权限用户可进行审核帖子<br><li>版主、管理员和开放权限用户可直接发贴<br><li>一般用户如发贴内容如果有被过滤的敏感字需审核后帖子方可见,<br>如果无被过滤的内容，则可免审核发贴。">
</td>
<td class=Forumrow><a href=# onclick="helpscript(bnew);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>敏感字设置</U></td>
<td colspan=2 class=Forumrow>
<input type="text" Name=Board_Setting(58) Value="<%=Board_Setting(58)%>" Size=50><br>可设置多个敏感字中间用"|"分隔如不填写可以填0
<input type="hidden" id="bnewS" value="<b>敏感字设置</b><br><li>可设置多个敏感字中间用 | 分隔">
</td>
<td class=Forumrow><a href=# onclick="helpscript(bnewS);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>允许同时在线数</U><BR>不限制则设置为0</td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(18)" value="<%=Board_Setting(18)%>"> 人
</td>
<input type="hidden" id="b9" value="<b>允许同时在线数</b><br><li>不限制则设置为0，如设置了允许同时在线数，则当论坛在线人数超过此数字的时候未登录用户将不能访问该版面">
<td class=ForumrowHighlight><a href=# onclick="helpscript(b9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>论坛定时设置：</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(21)" value="0" <%If Board_Setting(21)="0" Then %>checked <%End If%>>关 闭</option>
<input type=radio name="Board_Setting(21)" value="1" <%If Board_Setting(21)="1" Then %>checked <%End If%>>定时关闭</option>
<input type=radio name="Board_Setting(21)" value="2" <%If Board_Setting(21)="2" Then %>checked <%End If%>>定时只读</option>
</td>
<input type="hidden" id="b10" value="<b>定时设置选择:</b><br><li>在这里您可以设置是否起用定时的各种功能，如果开启了本功能，请设置好下面选项中的论坛设置时间，论坛该版面将在您规定的时间内有指定的设置">
<td class=Forumrow><a href=# onclick="helpscript(b10);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>定时设置</U><BR>请根据需要选择开或关</td></td>
<td colspan=2 class=ForumrowHighlight>
<%
Board_Setting(22)=split(Board_Setting(22),"|")
If UBound(Board_Setting(22))<2 Then 
	Board_Setting(22)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
	Board_Setting(22)=split(Board_Setting(22),"|")
End If
For i= 0 to UBound(Board_Setting(22))
If i<10 Then Response.Write "&nbsp;"
%>
 <%=i%>点：<input type="checkbox" name="Board_Setting(22)<%=i%>" value="1" <%If Board_Setting(22)(i)="1" Then %>checked<%End If%>>开
   
 <%
 If (i+1) mod 4 = 0 Then Response.Write "<br>"
 Next
 %>
</td>
<input type="hidden" id="b11" value="<b>论坛开放时间</b><br><li>设置了本选项必须同时打开是否起用定时开关论坛设置才有效，设置了此选项，论坛该版面将在您规定的时间内给用户开放">
<td class=ForumrowHighlight><a href=# onclick="helpscript(b11);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<%
Dim VisitConfirm
VisitConfirm=Split(Board_Setting(54),"|")
IF Ubound(VisitConfirm)<>8 Then
	Redim VisitConfirm(8)
	For i=0 To 8
	VisitConfirm(i)=0
	Next
End If
%>
<tr> 
<td colspan=2 class=Forumrow>
<U>用户至少文章数</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(0)%>">
</td>
<input type="hidden" id="VisitConfirm1" value="<b>用户至少文章数</b><br><li>当用户发表的文章达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(VisitConfirm1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>用户至少积分</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(1)%>">
</td>
<input type="hidden" id="VisitConfirm2" value="<b>用户至少积分值</b><br><li>当用户的积分值达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(VisitConfirm2);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>用户至少金钱</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(2)%>">
</td>
<input type="hidden" id="VisitConfirm3" value="<b>用户至少金钱数</b><br><li>当用户的金钱达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(VisitConfirm3);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>用户至少魅力</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(3)%>">
</td>
<input type="hidden" id="VisitConfirm4" value="<b>用户至少魅力</b><br><li>当用户的魅力值达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(VisitConfirm4);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>用户至少威望</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(4)%>">
</td>
<input type="hidden" id="VisitConfirm5" value="<b>用户至少威望</b><br><li>当用户威望达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(VisitConfirm5);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>用户至少精华文章</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(5)%>">
</td>
<input type="hidden" id="VisitConfirm6" value="<b>用户至少精华文章数</b><br><li>当用户发表的精华文章达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(VisitConfirm6);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>用户被删帖子数上限</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(6)%>">
</td>
<input type="hidden" id="VisitConfirm7" value="<b>用户被删帖子数上限</b><br><li>当用户被删帖子数超过此设置时，不能访问该分版！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(VisitConfirm7);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>至少注册时间（单位为分钟）</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(7)%>">
</td>
<input type="hidden" id="VisitConfirm8" value="<b>用户至少注册时间</b><br><li>注册时间是指用户注册多少分钟后可进入论坛。<li>单位为分钟。<li>不限制设置为0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(VisitConfirm8);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>至少上传文件个数</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(8)%>">
</td>
<input type="hidden" id="VisitConfirm9" value="<b>用户至少上传文件个数</b><br><li>当用户至少上传文件个数达到此设置时，才能拥有访问权限！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(VisitConfirm9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting3">前台管理权限</a>[<a href="#top">顶部</a>]</th></tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>主版主可以增删副版主</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(33)" value=0 <%if Board_Setting(33)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(33)" value=1 <%if Board_Setting(33)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>主版主可以修改广告设置</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(34)" value=0 <%if Board_Setting(34)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(34)" value=1 <%if Board_Setting(34)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>所有版主可以修改广告设置</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(35)" value=0 <%if Board_Setting(35)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(35)" value=1 <%if Board_Setting(35)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting4">发贴相关</a>[<a href="#top">顶部</a>]</th>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>发贴是否采用验证码</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(4)" value=1 <%if Board_Setting(4)="1" then%>checked<%end if%>>采用&nbsp;
<input type=radio name="Board_Setting(4)" value=0 <%if Board_Setting(4)="0" then%>checked<%end if%>>不采用&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>主题限制长度</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(45)" value="<%=Board_Setting(45)%>"> Byte
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>发贴后返回</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(17)" value=1 <%if Board_Setting(17)="1" then%>checked<%end if%>>首页&nbsp;
<input type=radio name="Board_Setting(17)" value=2 <%if Board_Setting(17)="2" then%>checked<%end if%>>论坛&nbsp;
<input type=radio name="Board_Setting(17)" value=3 <%if Board_Setting(17)="3" then%>checked<%end if%>>帖子&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>帖子内容最大字节数</U><BR>1024字节等于1K</td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(16)" value="<%=Board_Setting(16)%>"> 字节
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>帖子内容最小字节数</U><BR>1024字节等于1K</td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(52)" value="<%=Board_Setting(52)%>"> 字节
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>投票后是否将投票贴提升到帖子列表顶部</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(53)" value=0 <%if Board_Setting(53)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(53)" value=1 <%if Board_Setting(53)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>上传文件类型</U><BR>每种文件类型用“|”号分开</td>
<td colspan=2 class=Forumrow>
<input type=text size=50 name="Board_Setting(19)" value="<%=Board_Setting(19)%>">
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>是否起用防灌水机制</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(30)" value=0 <%if Board_Setting(30)="0" then%>checked<%end if%>>否&nbsp;
<input type=radio name="Board_Setting(30)" value=1 <%if Board_Setting(30)="1" then%>checked<%end if%>>是&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>每次发贴间隔</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(31)" value="<%=Board_Setting(31)%>"> 秒
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>最多投票项目</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(32)" value="<%=Board_Setting(32)%>"> 个
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting5">帖子列表显示相关</a>[<a href="#top">顶部</a>]</th></tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>帖子列表标题显示字符数</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(25)" value="<%=Board_Setting(25)%>">
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>帖子列表每页记录数</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(26)" value="<%=Board_Setting(26)%>">
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>浏览帖子每页记录数</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(27)" value="<%=Board_Setting(27)%>">
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>帖子列表默认读取数据量</U></td>
<td colspan=2 class=Forumrow>
<select size="1" name="Board_Setting(37)">
<option value="1"<%if Board_Setting(37)="0" then%> selected<%end if%>>全部显示帖子</option>
<option value="2"<%if Board_Setting(37)="5" then%> selected<%end if%>>五天内帖子</option>
<option value="3"<%if Board_Setting(37)="15" then%> selected<%end if%>>半月内帖子</option>
<option value="4"<%if Board_Setting(37)="30" then%> selected<%end if%>>一月内帖子</option>
<option value="5"<%if Board_Setting(37)="60" then%> selected<%end if%>>两月内帖子</option>
<option value="6"<%if Board_Setting(37)="120" then%> selected<%end if%>>四月内帖子</option>
<option value="7"<%if Board_Setting(37)="180" then%> selected<%end if%>>半年内帖子</option>
</select>
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>显示最新帖图片显示方式</U></td>
<td colspan=2 class=ForumrowHighlight>
<select size="1" name="Board_Setting(38)">
<option value="0"<%if Board_Setting(38)="0" then%> selected<%end if%>>最后回复时间</option>
<option value="1"<%if Board_Setting(38)="1" then%> selected<%end if%>>发贴时间</option>
</select>
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>显示最新帖图片标识时间设置</U></td>
<td colspan=2 class=Forumrow>
<select size="1" name="Board_Setting(61)">
<option value="0"<%if Board_Setting(61)="0" then%> selected<%end if%>>0分钟</option>
<option value="10"<%if Board_Setting(61)="10" then%> selected<%end if%>>10分钟</option>
<option value="30"<%if Board_Setting(61)="30" then%> selected<%end if%>>30分钟</option>
<option value="60"<%if Board_Setting(61)="60" then%> selected<%end if%>>1小时</option>
<option value="360"<%If Board_Setting(61)="360" then%> selected<%end if%>>6小时</option>
<option value="720"<%if Board_Setting(61)="720" then%> selected<%end if%>>12小时</option>
<option value="1440"<%if Board_Setting(61)="1440" then%> selected<%end if%>>1天</option>
<option value="2880"<%if Board_Setting(61)="2880" then%> selected<%end if%>>2天</option>
</select>：内更新的帖子
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>显示最新帖图片地址设置（new）:值为0或空时即不显示，填写准确地址；</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=30 name="Board_Setting(60)" value="<%=Board_Setting(60)%>">
<%
If instr(Board_Setting(60),".gif") Then Response.Write "<img src="""&Board_Setting(60)&""" border=0>"
%>
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting6">帖子内容显示相关</a>[<a href="#top">顶部</a>]</th>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>HTML代码解析</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(5)" value=0 <%if Board_Setting(5)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(5)" value=1 <%if Board_Setting(5)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>UBB代码解析</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(6)" value=0 <%if Board_Setting(6)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(6)" value=1 <%if Board_Setting(6)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>贴图标签</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(7)" value=0 <%if Board_Setting(7)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(7)" value=1 <%if Board_Setting(7)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>表情标签</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(8)" value=0 <%if Board_Setting(8)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(8)" value=1 <%if Board_Setting(8)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>Flash标签</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(44)" value=0 <%if Board_Setting(44)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(44)" value=1 <%if Board_Setting(44)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>多媒体标签</U><BR>包括RM,AVI等</td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(9)" value=0 <%if Board_Setting(9)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(9)" value=1 <%if Board_Setting(9)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>是否开放金钱贴</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(10)" value=0 <%if Board_Setting(10)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(10)" value=1 <%if Board_Setting(10)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>是否开放积分贴</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(11)" value=0 <%if Board_Setting(11)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(11)" value=1 <%if Board_Setting(11)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>是否开放魅力贴</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(12)" value=0 <%If Board_Setting(12)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(12)" value=1 <%If Board_Setting(12)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>是否开放威望贴</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(13)" value=0 <%if Board_Setting(13)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(13)" value=1 <%if Board_Setting(13)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>是否开放文章贴</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(14)" value=0 <%if Board_Setting(14)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(14)" value=1 <%if Board_Setting(14)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>是否开放回复可见贴</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(15)" value=0 <%if Board_Setting(15)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(15)" value=1 <%if Board_Setting(15)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>是否开放出售帖子功能</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(23)" value=0 <%if Board_Setting(23)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(23)" value=1 <%if Board_Setting(23)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>是否开放定员帖子功能</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(56)" value=0 <%if Board_Setting(56)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Board_Setting(56)" value=1 <%if Board_Setting(56)="1" then%>checked<%end if%>>开放&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>帖子正文字号</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(28)" value="<%=Board_Setting(28)%>">
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>帖子正文行间距</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(29)" value="<%=Board_Setting(29)%>">
</td>
<td class=Forumrow></td>
</tr>
<%
Dim DownConfirm
DownConfirm=Split(Board_Setting(55),"|")
IF Ubound(DownConfirm)<>8 Then
	Redim DownConfirm(8)
	For i=0 To 8
	DownConfirm(i)=0
	Next
End If
%>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting7">下载附件限制设置</a>[<a href="#top">顶部</a>]</th></tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>用户至少文章数</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(0)%>">
</td>
<input type="hidden" id="Down1" value="<b>用户至少文章数</b><br><li>当用户发表的文章达到此设置时，才能拥有下载权限！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(Down1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>用户至少积分</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(1)%>">
</td>
<input type="hidden" id="Down2" value="<b>用户至少积分值</b><br><li>当用户的积分值达到此设置时，才能拥有下载权限！<li>不限制设置为0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(Down2);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>用户至少金钱</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(2)%>">
</td>
<input type="hidden" id="Down3" value="<b>用户至少金钱数</b><br><li>当用户的金钱达到此设置时，才能拥有下载权限！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(Down3);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>用户至少魅力</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(3)%>">
</td>
<input type="hidden" id="Down4" value="<b>用户至少魅力</b><br><li>当用户的魅力值达到此设置时，才能拥有下载权限！<li>不限制设置为0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(Down4);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>用户至少威望</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(4)%>">
</td>
<input type="hidden" id="Down5" value="<b>用户至少威望</b><br><li>当用户威望达到此设置时，才能拥有下载权限！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(Down5);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>用户至少精华文章</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(5)%>">
</td>
<input type="hidden" id="Down6" value="<b>用户至少精华文章数</b><br><li>当用户发表的精华文章达到此设置时，才能拥有下载权限！<li>不限制设置为0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(Down6);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>用户被删帖子数上限</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(6)%>">
</td>
<input type="hidden" id="Down7" value="<b>用户被删帖子数上限</b><br><li>当用户被删帖子数超过此设置时，不能下载该版附件！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(Down7);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>至少注册时间</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(7)%>">
</td>
<input type="hidden" id="Down8" value="<b>用户至少注册天数</b><br><li>当用户至少注册分钟达到此设置时，才能拥有下载权限！<li>以分钟为单位，不限制为0。">
<td class=ForumrowHighlight><a href=# onclick="helpscript(Down8);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>至少上传文件个数</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(8)%>">
</td>
<input type="hidden" id="Down9" value="<b>用户至少上传文件个数</b><br><li>当用户至少上传文件个数达到此设置时，才能拥有下载权限！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(Down9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting9">论坛虚拟形象设置</a>[<a href="#top">顶部</a>]</th></tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>浏览帖子中虚拟形象</U></td>
<td colspan=2 class=Forumrow>
<input type="radio" name="Board_Setting(59)" value="0"
<%
If Board_Setting(59)="0" Then
%>
 checked 
 <%
 End If
 %>
 >&nbsp;显示全身
<input type="radio" name="Board_Setting(59)" value="1"
<%
If Board_Setting(59)="1" Then
%>
 checked 
 <%
 End If
 %>
 >&nbsp;显示脸
 <input type="radio" name="Board_Setting(59)" value="2"
<%
If Board_Setting(59)="2" Then
%>
 checked 
 <%
 End If
 %>
 >&nbsp;显示半身
 <input type="radio" name="Board_Setting(59)" value="3"
<%
If Board_Setting(59)="3" Then
%>
 checked 
 <%
 End If
 %>
 >&nbsp;不显示（显示头像）
</td>
<input type="hidden" id="xx9" value="<b>用户至少上传文件个数</b><br><li>当用户至少上传文件个数达到此设置时，才能拥有下载权限！<li>不限制设置为0">
<td class=Forumrow><a href=# onclick="helpscript(xx9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting8">论坛专题分类相关设置</a>[<a href="#top">顶部</a>]</th></tr>
<tr><td colspan="5" class=Forumrow>
<li>允许发表专题权限，请到相应用户组发帖权限中设置；
<li>专题栏目可以添加，修改；
<li>注意删除专题同时,会将最后专题的所有文章更新为普通主题。</td></tr>
<%
Dim BoardTopic,BoardTopicImg,ii
BoardTopic=Split(Board_Setting(48),"$$")
BoardTopicImg=Split(Board_Setting(49),"$$")
For ii=0 to Ubound(BoardTopic)-1
%>
<tr>
<td width="15%" class=Forumrow><U>专题名称:</U></td>
<td width="35%" class=Forumrow>
<input type=text size=30 name="BoardTopic" value="<%=Server.Htmlencode(BoardTopic(ii))%>"></td>
<td width="15%" class=Forumrow><U>相应显示图标：</U></td>
<td width="35%" class=Forumrow>
<input type=text size=30 name="BoardTopicImg" value="<%=BoardTopicImg(ii)%>">
<%
If BoardTopicImg(ii)<>"" and Instr(BoardTopicImg(ii),".gif") Then Response.Write "<img src="&BoardTopicImg(ii)&" border=0>"
%>
</td>
<td class=Forumrow></td>
</tr>
<%Next%>
<input type=hidden value="<%=ii%>" name="BoardTopicNum">
<tr>
<td width="15%" class=Forumrow><U>添加专题:</U></td>
<td width="35%" class=Forumrow>
<input type=text size=30 name="BoardTopic" value=""></td>
<td width="15%" class=Forumrow><U>相应显示图标：</U></td>
<td width="35%" class=Forumrow>
<input type=text size=30 name="BoardTopicImg" value=""></td>
<td class=Forumrow></td>
</tr>
<tr>
<td colspan=5 class=ForumRowHighlight>
<div align="center"> 
<input type=hidden value="<%=Board_Setting(20)%>" name="Board_Setting(20)">
<input type=hidden value="<%=Board_Setting(46)%>" name="Board_Setting(46)">
<input type=hidden value="<%=Board_Setting(47)%>" name="Board_Setting(47)">
<input type="submit" name="Submit" value="提 交">
</div>
</td>
</tr>
</form>
</table>
<%
end sub

sub saveconst()
if not isnumeric(request("editid")) then
Errmsg=ErrMsg + "<BR><li>错误的版面参数"
dvbbs_error()
exit sub
else
Dim iboard_setting,isetting
Dim BoardTopic,BoardTopicImg,TempStr,ii,BoardTopicNum
Dim DownConfirm,ViewConfirm
	ii=0
	i=0
	For Each TempStr in Request.Form("Board_Setting(54)")
		i=i+1
		ViewConfirm=ViewConfirm&TempStr
		If i<>Request.Form("Board_Setting(54)").count Then
		ViewConfirm=ViewConfirm&"|"
		End If
	Next
	i=0
	If not ISNumeric(Replace(ViewConfirm,"|","")) or Request.Form("Board_Setting(54)").count<>9 Then
		Errmsg=ErrMsg + "<BR><li>下载附件参数有错误，提交被中止。"
		dvbbs_error()
		exit sub
	End if
	For Each TempStr in Request.Form("Board_Setting(55)")
		i=i+1
		DownConFirm=DownConFirm&TempStr
		If i<>Request.Form("Board_Setting(55)").count Then
		DownConFirm=DownConFirm&"|"
		End If
	Next
	i=0
	If not ISNumeric(Replace(DownConFirm,"|","")) or Request.Form("Board_Setting(55)").count<>9 Then
		Errmsg=ErrMsg + "<BR><li>下载附件参数有错误，提交被中止。"
		dvbbs_error()
		exit sub
	End if
	
	IF Request("BoardTopicNum")<>"" and Isnumeric(Request("BoardTopicNum")) Then
	BoardTopicNum=Request("BoardTopicNum") 
	Else
	BoardTopicNum=0
	End If
	For Each TempStr in Request.form("BoardTopic")
		If TempStr<>"" Then 
			BoardTopic=BoardTopic&TempStr&"$$"
			ii=ii+1
		End If
	Next
	TempStr=""
	For Each TempStr in Request.form("BoardTopicImg")
			BoardTopicImg=BoardTopicImg&TempStr&"$$"
	Next
	TempStr=""
	If ii>99 Then
		Errmsg=ErrMsg + "<BR><li>专题栏目数目在１００以内。"
		dvbbs_error()
		exit sub
	End If
	Dim setingdata,j
	For i = 0 To 70
		If Trim(request.Form("Board_Setting("&i&")"))="" Or i=22 Then
			'Response.Write "Board_Setting("&i&")<br>"
			isetting=0
			If i=22 Then
				isetting=""
				For j=0 to  23
					If isetting="" Then
						If Request.form("Board_Setting(22)"&j)="1" Then
							isetting="1"
						Else
							isetting="0"
						End If
					Else
						If Request.form("Board_Setting(22)"&j)="1" Then
							isetting=isetting&"|1"
						Else
							isetting=isetting&"|0"
						End If
					End If
				Next
			End If
		Else
			isetting=Replace(Trim(request.Form("Board_Setting("&i&")")),",","")
		End If
		If i = 0 Then
			iboard_Setting = isetting
		ElseIf i = 48 Then
			iboard_Setting = iboard_Setting & "," & BoardTopic
		ElseIf i = 49 Then
			iboard_Setting = iboard_Setting & "," & BoardTopicImg
		ElseIf i=54 Then
			iboard_Setting = iboard_Setting & "," & ViewConfirm
		ElseIf i=55 Then 
			iboard_Setting = iboard_Setting & "," & DownConFirm
		Else
			iboard_Setting = iboard_Setting & "," & isetting
		End If
	Next

Dim FoundCKBoard
FoundCKBoard=False
For i=0 to UBOUND(Dvbbs.Forum_Setting)
	If request.Form("CK_Board_Setting("&i&")")<>"" Then
		FoundCKBoard=True
		Exit For
	End If
Next

Dim Forum_Boards,upBoardid,upid,temprs
select case request("savetype")
'当前版面
case "0"
	Dvbbs.Execute("update dv_board set board_setting='"&iboard_setting&"' where boardid="&Request("editid"))
	Dvbbs.ReloadBoardInfo(Request("editid"))
	upBoardid=" and boardid="&Request("editid")
'所有版面
case "1"
	Dvbbs.Execute("update dv_board set board_setting='"&iboard_setting&"'")
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
	upBoardid=""
'相同分类下所有版面（不包括分类）
case "2"
	set rs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("editid"))
	if not rs.eof then
		Dvbbs.Execute("update dv_board set board_setting='"&iboard_setting&"' where (Not ParentID=0) and rootid="&rs(0))
		Set temprs=Dvbbs.Execute("Select boardid from Dv_board where (Not ParentID=0) and rootid="&rs(0))
		if not temprs.eof then
			upid=temprs.GetString(,, "",",","")
		end if
		temprs.close:Set temprs=Nothing
	end if
	rs.close:set rs=nothing
	upBoardid=" and boardid in ("&left(upid,(len(upid)-1))&")"
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
'相同分类下所有版面（包括分类）
case "3"
	set rs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("editid"))
	if not rs.eof then
		Dvbbs.Execute("update dv_board set board_setting='"&iboard_setting&"' where rootid="&rs(0))
		Set temprs=Dvbbs.Execute("select boardid from dv_board where rootid="&rs(0))
		if not temprs.eof then
			upid=temprs.GetString(,, "",",","")
		end if
		temprs.close:Set temprs=Nothing
	end if
	rs.close:set rs=nothing
	upBoardid=" and boardid in ("&left(upid,(len(upid)-1))&")"

	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
'同分类同级别版面
case "4"
	set rs=Dvbbs.Execute("select rootid,ParentStr,ParentID from dv_board where boardid="&request("editid"))
	if not rs.eof then
		Dvbbs.Execute("update dv_board set board_setting='"&iboard_setting&"' where rootid="&rs(0)&" and ParentID="&rs(2)&" and ParentStr='"&rs(1)&"'")
		Set temprs=Dvbbs.Execute("select boardid from dv_board where rootid="&rs(0)&" and ParentID="&rs(2)&" and ParentStr='"&rs(1)&"'")
		if not temprs.eof then
			upid=temprs.GetString(,, "",",","")
		end if
		temprs.close:Set temprs=Nothing
	end if
	rs.close:set rs=nothing
	upBoardid=" and boardid in ("&left(upid,(len(upid)-1))&")"
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
End Select

If BoardTopicNum>ii Then
	Dvbbs.Execute("update Dv_Topic set Mode=0 where Mode >= "&ii+1&" "&upBoardid&" ")
End If

dv_suc("设置成功。<a href=admin_boardsetting.asp?editid="&request("editid")&">返回版面高级设置</a>")
End If
End sub
%>
