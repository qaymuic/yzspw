<%
Function GroupPermission(GroupSetting)
Dim reGroupSetting,Rs,UserHtml,UserHtmlA,UserHtmlB
If GroupSetting="" Then
	Set Rs = Dvbbs.Execute("Select GroupSetting From Dv_UserGroups Where UserGroupID=4")
	reGroupSetting = Split(Rs(0),",")
Else
	reGroupSetting = Split(GroupSetting,",")
End If
If reGroupSetting(58)="0" Then reGroupSetting(58)="§"
UserHtml = Split(reGroupSetting(58),"§")
If Ubound(UserHtml)=1 Then
	UserHtmlA=UserHtml(0)
	UserHtmlB=UserHtml(1)
Else
	UserHtmlA=""
	UserHtmlB=""
End If
%>
<tr> 
<th height="23" colspan="3"  align=left>＝＝浏览相关选项</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>用户名在帖子内容中显示标记<BR>HTML语法，左右标记代码将加于用户名前后两头</td>
<td height="23" width="40%" class=Forumrow>左标记 <input name="GroupSetting(58)A" type=text size=8 value="<%=Server.HtmlEncode(UserHtmlA)%>"> 右标记 <input name="GroupSetting(58)B" type=text size=8 value="<%=Server.HtmlEncode(UserHtmlB)%>"></td>
<input type="hidden" id="g1" value="<b>用户名在帖子内容中显示标记</b><br><li>HTML语法，左右标记代码将加于用户名前后两头<br><li>如您设置了前后分别为《b》和《/b》，则在帖子内容中该组用户或者相关等级用户名显示为<B>粗体</B>">
<td class=forumRow><a href=# onclick="helpscript(g1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>允许用户自选风格</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(57)" type=radio value="1" <%if reGroupSetting(57)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(57)" type=radio value="0" <%if reGroupSetting(57)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g2" value="<b>允许用户自选风格</b><br><li>如果关闭了本选项，论坛中用户将不能自己选择浏览显示的风格（包括用户在个人信息中设定的风格）">
<td class=forumRowHighlight><a href=# onclick="helpscript(g2);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以浏览论坛</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(0)" type=radio value="1" <%if reGroupSetting(0)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(0)" type=radio value="0" <%if reGroupSetting(0)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g3" value="<b>用户名在帖子内容中显示标记</b><br><li>关闭此选项，相关组或等级用户将不能浏览论坛<br><li>使用技巧：您可以设定某个用户组不能使用本设置，而当其身份变化后的用户组可使用本设置，如设置客人不能使用本设置，这样将迫使他登录">
<td class=forumRow><a href=# onclick="helpscript(g3);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以查看会员信息(包括其他会员的资料和会员列表)
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(1)" type=radio value="1" <%if reGroupSetting(1)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(1)" type=radio value="0" <%if reGroupSetting(1)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g4" value="<b>可以查看会员信息</b><br><li>关闭此选项，相关组或等级用户将不能浏览论坛用户资料，包括会员资料和会员列表资料<br><li>使用技巧：您可以设定某个用户组不能使用本设置，而当其身份变化后的用户组可使用本设置，如设置客人不能使用本设置，这样将迫使他登录">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g4);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以查看其他人发布的主题
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(2)" type=radio value="1" <%if reGroupSetting(2)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(2)" type=radio value="0" <%if reGroupSetting(2)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g5" value="<b>可以查看其他人发布的主题</b><br><li>关闭此选项，相关组或等级用户将不能浏览论坛中其他人发布的帖子<br><li>使用技巧：您可以设定某个用户组不能使用本设置，而当其身份变化后的用户组可使用本设置，如设置客人不能使用本设置，这样将迫使他登录">
<td class=Forumrow><a href=# onclick="helpscript(g5);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以浏览精华帖子
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(41)" type=radio value="1" <%if reGroupSetting(41)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(41)" type=radio value="0" <%if reGroupSetting(41)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g6" value="<b>可以浏览精华帖子</b><br><li>关闭此选项，相关组或等级用户将不能浏览论坛中的精华帖子<br><li>使用技巧：您可以设定某个用户组不能使用本设置，而当其身份变化后的用户组可使用本设置，如设置客人不能使用本设置，这样将迫使他登录">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g6);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<th height="23" colspan="3"  align=left>＝＝发帖权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布新主题</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(3)" type=radio value="1" <%if reGroupSetting(3)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(3)" type=radio value="0" <%if reGroupSetting(3)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g9" value="<b>可以发布新主题</b><br><li>打开此选项，相关组或等级用户将可以可以发布新主题。鉴于国家规定，论坛默认的未登录用户组将即使设置此选项也不能发贴">
<td class=Forumrow><a href=# onclick="helpscript(g9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>在审核模式下可直接发贴而不需经过审核</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(64)" type=radio value="1" <%if reGroupSetting(64)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(64)" type=radio value="0" <%if reGroupSetting(64)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g10" value="<b>在审核模式下可直接发贴而不需经过审核</b><br><li>打开此选项，相关组或等级用户将可以可以发布新主题或回复而不经审核<br><li>当论坛版面设置为审核状态时该选项有效">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g10);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>一天最多发贴数目
</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(62)" type=text size=4 value="<%=reGroupSetting(62)%>"></td>
<input type="hidden" id="g11" value="<b>一天最多发贴数目</b><br><li>填写0为不作限制，出于对付灌水或者使用软件发贴的用户，请在此设置合理的数字<br><li>使用技巧：您可以给不同用户组设置不同的数字">
<td class=Forumrow><a href=# onclick="helpscript(g11);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以回复自己的主题
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(4)" type=radio value="1" <%if reGroupSetting(4)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(4)" type=radio value="0" <%if reGroupSetting(4)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g12" value="<b>可以回复自己的主题</b><br><li>打开此选项，相关用户组或等级用户可以回复自己发布的主题">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g12);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以回复其他人的主题
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(5)" type=radio value="1" <%if reGroupSetting(5)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(5)" type=radio value="0" <%if reGroupSetting(5)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g13" value="<b>可以回复其他人的主题</b><br><li>打开此选项，相关用户组或等级用户可以回复其他人的主题">
<td class=Forumrow><a href=# onclick="helpscript(g13);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以在论坛允许评分的时候参与评分(鲜花和鸡蛋)?
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(6)" type=radio value="1" <%if reGroupSetting(6)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(6)" type=radio value="0" <%if reGroupSetting(6)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g14" value="<b>可以在论坛允许评分的时候参与评分</b><br><li>打开此选项，相关用户组或等级用户可以在论坛允许评分的时候参与评分，也就是帖子内容中的鲜花和鸡蛋选项">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g14);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>参与评分所需金钱
</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(47)" type=text size=4 value="<%=reGroupSetting(47)%>"></td>
<input type="hidden" id="g15" value="<b>参与评分所需金钱</b><br><li>相关用户组或等级用户参与评分所需金钱，点击鲜花或鸡蛋后扣除">
<td class=Forumrow><a href=# onclick="helpscript(g15);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以发布新投票</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(8)" type=radio value="1" <%if reGroupSetting(8)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(8)" type=radio value="0" <%if reGroupSetting(8)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g21" value="<b>可以发布新投票</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以发布新投票">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g21);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以参与投票</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(9)" type=radio value="1" <%if reGroupSetting(9)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(9)" type=radio value="0" <%if reGroupSetting(9)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g22" value="<b>可以发布新投票</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以参与投票">
<td class=Forumrow><a href=# onclick="helpscript(g22);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>投票可以使用HTML语法</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(68)" type=radio value="1" <%if reGroupSetting(68)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(68)" type=radio value="0" <%if reGroupSetting(68)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g_u_HTML" value="<b>投票可以使用HTML语法</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以在投票中使用HTML语法">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g_u_HTML);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以发布小字报</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(17)" type=radio value="1"  <%if reGroupSetting(17)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(17)" type=radio value="0" <%if reGroupSetting(17)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g23" value="<b>可以发布小字报</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以发布小字报">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g23);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>发布小字报所需金钱</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(46)" type=text value="<%=reGroupSetting(46)%>" size=4></td>
<input type="hidden" id="g24" value="<b>发布小字报所需金钱</b><br><li>在这里您可以根据需要设置不同用户组或等级用户发布小字报所需金钱">
<td class=Forumrow><a href=# onclick="helpscript(g24);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以发布特殊标题帖子（如标题加红、UBB语法等）</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(51)" type=radio value="1"  <%if reGroupSetting(51)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(51)" type=radio value="0" <%if reGroupSetting(51)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g25" value="<b>可以发布特殊标题帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户可以发布特殊标题帖子，如标题加颜色、HTML语法、UBB语法等，您可针对个别用户组可使用此特殊功能">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g25);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<!--
<tr>
<td height="23" width="60%" class=Forumrow>发表模式选择</td>
<td height="23" width="40%" class=Forumrow>
<select name="GroupSetting(67)" >
<option value="0"  <%if reGroupSetting(67)="0" then%>selected<%end if%>>关闭HTML编辑
<option value="1"  <%if reGroupSetting(67)="1" then%>selected<%end if%>>允许HTML编辑
<option value="2"  <%if reGroupSetting(67)="2" then%>selected<%end if%>>简单模式编辑
<option value="3"  <%if reGroupSetting(67)="3" then%>selected<%end if%>>全功能编辑
</select>
</td>
<input type="hidden" id="g0" value="<b>发表模式选择</b><br><li>发表模式包括：Design编辑模式,Ubb简单模式，HTML可编辑模式；<li>关闭HTML编辑：当版块允许发表高级模式下，用户只保留Design编辑模式和Ubb简单模式；<li>允许HTML编辑：当版块允许发表高级模式下，用户拥有全功能编辑模式；<li>简单模式编辑：当版块允许发表高级模式下，用户只保留Ubb简单模式；<li>全功能编辑：当版块在发表简单模式下，拥有所有发表模式；<li>为避免用户滥用HTML的各种语法，建议只对部分用户关闭HTML编辑；">
<td class=Forumrow><a href=# onclick="helpscript(g0);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
-->
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以发表论坛专题</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(65)" type=radio value="1"  <%if reGroupSetting(65)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(65)" type=radio value="0" <%if reGroupSetting(65)="0" then%>checked<%end if%>></td>
<td class=ForumrowHighLight><a href=# class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>新注册用户多少分钟后才能发言</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(52)" type=text value="<%=reGroupSetting(52)%>" size=4> 分钟</td>
<input type="hidden" id="g26" value="<b>新注册用户多少分钟后才能发言</b><br><li>在这里您可以根据需要设置不同用户组或等级用户新注册需要多少分钟后才能发言，建议合理设置此选项，以避免一些恶意用户乱注册散发非法帖子或广告帖子">
<td class=Forumrow><a href=# onclick="helpscript(g26);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<th height="23" colspan="3"  align=left>＝＝<b>帖子/主题编辑权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以编辑自己的帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(10)" type=radio value="1" <%if reGroupSetting(10)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(10)" type=radio value="0" <%if reGroupSetting(10)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g27" value="<b>可以编辑自己的帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以编辑自己的帖子">
<td class=Forumrow><a href=# onclick="helpscript(g27);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以删除自己的帖子
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(11)" type=radio value="1" <%if reGroupSetting(11)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(11)" type=radio value="0" <%if reGroupSetting(11)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g28" value="<b>可以删除自己的帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以删除自己的帖子，请根据自己的需要合理设置此选项">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g28);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以移动自己的帖子到其他论坛
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(12)" type=radio value="1" <%if reGroupSetting(12)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(12)" type=radio value="0" <%if reGroupSetting(12)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g29" value="<b>可以移动自己的帖子到其他论坛</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以移动自己的帖子到其他论坛，请根据自己的需要合理设置此选项">
<td class=Forumrow><a href=# onclick="helpscript(g29);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以打开/关闭自己发布的主题
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(13)" type=radio value="1" <%if reGroupSetting(13)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(13)" type=radio value="0" <%if reGroupSetting(13)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g30" value="<b>可以打开/关闭自己发布的主题</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以打开/关闭自己发布的主题，请根据自己的需要合理设置此选项">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g30);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<th height="23" colspan="3" align=left>＝＝上传权限设置</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以上传附件
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(7)" type=radio value="1" <%if reGroupSetting(7)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(7)" type=radio value="0" <%if reGroupSetting(7)="0" then%>checked<%end if%>>
&nbsp;发帖可以上传<input name="GroupSetting(7)" type=radio value="2" <%if reGroupSetting(7)="2" then%>checked<%end if%>>&nbsp;回复可以上传<input name="GroupSetting(7)" type=radio value="3" <%if reGroupSetting(7)="3" then%>checked<%end if%>>
</td>
<input type="hidden" id="g16" value="<b>可以上传附件</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以上传附件，选择是则发贴和回贴都可以上传，否则不行。您也可以可以根据需要分别设置发贴或回帖是否可以上传">
<td class=Forumrow><a href=# onclick="helpscript(g16);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>一次批量上传数量(设置为0，即不允许使用此功能;建议不要超过5个)
</td>
<td height="23" width="40%" class=ForumrowHighLight><input name="GroupSetting(66)" type=text size=4 value="<%=reGroupSetting(66)%>"></td>
<input type="hidden" id="GroupSetting66" value="<b>一次批量上传数量</b><br><li>设置为0，即不允许使用此功能;<li>建议不要超过5个，因为上传操作将消耗大量服务器资源">
<td class=ForumrowHighLight><a href=# onclick="helpscript(GroupSetting66);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>一次最多上传文件个数
</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(40)" type=text size=4 value="<%=reGroupSetting(40)%>"></td>
<input type="hidden" id="g17" value="<b>一次最多上传文件个数</b><br><li>在这里您可以根据需要设置不同用户组或等级用户一次最多上传文件个数，建议不要设置过大，因为上传操作将消耗大量服务器资源">
<td class=Forumrow><a href=# onclick="helpscript(g17);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>一天最多上传文件个数
</td>
<td height="23" width="40%" class=ForumrowHighLight><input name="GroupSetting(50)" type=text size=4 value="<%=reGroupSetting(50)%>"></td>
<input type="hidden" id="g18" value="<b>一天最多上传文件个数</b><br><li>在这里您可以根据需要设置不同用户组或等级用户一天最多上传文件个数，建议不要设置过大，因为上传操作将消耗大量服务器资源">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g18);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>上传文件大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(44)" type=text size=4 value="<%=reGroupSetting(44)%>"> KB</td>
<input type="hidden" id="g19" value="<b>上传文件大小限制</b><br><li>在这里您可以根据需要设置不同用户组或等级用户上传文件大小，建议不要设置过大，因为上传操作将消耗大量服务器资源">
<td class=Forumrow><a href=# onclick="helpscript(g19);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以下载附件</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(61)" type=radio value="1" <%if reGroupSetting(61)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(61)" type=radio value="0" <%if reGroupSetting(61)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g20" value="<b>可以下载附件</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以下载附件，比如可以设置未登录用户不许下载">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g20);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<th height="23" colspan="3" align=left>＝＝管理权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(18)" type=radio value="1" <%if reGroupSetting(18)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(18)" type=radio value="0"  <%if reGroupSetting(18)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g38" value="<b>可以删除其它人帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以删除其它人帖子，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=Forumrow><a href=# onclick="helpscript(g38);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以移动其它人帖子
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(19)" type=radio value="1" <%if reGroupSetting(19)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(19)" type=radio value="0"  <%if reGroupSetting(19)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g39" value="<b>可以移动其它人帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以移动其它人帖子，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g39);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以打开/关闭其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(20)" type=radio value="1" <%if reGroupSetting(20)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(20)" type=radio value="0"  <%if reGroupSetting(20)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g40" value="<b>可以打开/关闭其它人帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以打开/关闭其它人帖子，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=Forumrow><a href=# onclick="helpscript(g40);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以固顶/解除固顶帖子
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(21)" type=radio value="1" <%if reGroupSetting(21)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(21)" type=radio value="0"  <%if reGroupSetting(21)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g41" value="<b>可以固顶/解除固顶帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以固顶/解除固顶帖子，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g41);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以进行帖子区域固顶操作
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(54)" type=radio value="1" <%if reGroupSetting(54)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(54)" type=radio value="0"  <%if reGroupSetting(54)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g42" value="<b>可以进行帖子区域固顶操作</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以进行帖子区域固顶操作，请根据自己的需要合理设置此选项，建议对超级版主及其以上用户组设置此权限">
<td class=Forumrow><a href=# onclick="helpscript(g42);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以进行帖子总固顶操作
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(38)" type=radio value="1"  <%if reGroupSetting(38)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(38)" type=radio value="0" <%if reGroupSetting(38)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g43" value="<b>可以进行帖子总固顶操作</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以进行帖子总固顶操作，请根据自己的需要合理设置此选项，建议对超级版主及其以上用户组设置此权限">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g43);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以奖励/惩罚发贴用户
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(22)" type=radio value="1" <%if reGroupSetting(22)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(22)" type=radio value="0"  <%if reGroupSetting(22)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g44" value="<b>可以奖励/惩罚发贴用户</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以奖励/惩罚发贴用户，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=Forumrow><a href=# onclick="helpscript(g44);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以奖励/惩罚用户
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(43)" type=radio value="1" <%if reGroupSetting(43)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(43)" type=radio value="0"  <%if reGroupSetting(43)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g45" value="<b>可以奖励/惩罚用户</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以奖励/惩罚用户，请根据自己的需要合理设置此选项，建议对超级版主及其以上用户组设置此权限">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g45);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以编辑其它人帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(23)" type=radio value="1" <%if reGroupSetting(23)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(23)" type=radio value="0" <%if reGroupSetting(23)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g46" value="<b>可以编辑其它人帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以编辑其它人帖子，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=Forumrow><a href=# onclick="helpscript(g46);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以加入/解除精华帖子
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(24)" type=radio value="1" <%if reGroupSetting(24)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(24)" type=radio value="0"  <%if reGroupSetting(24)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g47" value="<b>可以加入/解除精华帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以加入/解除精华帖子，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g47);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发布公告
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(25)" type=radio value="1" <%if reGroupSetting(25)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(25)" type=radio value="0"  <%if reGroupSetting(25)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g48" value="<b>可以发布公告</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以发布公告，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=Forumrow><a href=# onclick="helpscript(g48);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以管理公告
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(26)" type=radio value="1" <%if reGroupSetting(26)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(26)" type=radio value="0"  <%if reGroupSetting(26)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g49" value="<b>可以管理公告</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以管理公告，请根据自己的需要合理设置此选项，建议对超级版主及其以上用户组设置此权限">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g49);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以管理小字报
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(27)" type=radio value="1" <%if reGroupSetting(27)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(27)" type=radio value="0"  <%if reGroupSetting(27)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g50" value="<b>可以管理小字报</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以管理小字报，请根据自己的需要合理设置此选项，建议对超级版主及其以上用户组设置此权限">
<td class=Forumrow><a href=# onclick="helpscript(g50);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以锁定/屏蔽/解除锁定用户
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(28)" type=radio value="1" <%if reGroupSetting(28)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(28)" type=radio value="0"  <%if reGroupSetting(28)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g51" value="<b>可以锁定/屏蔽/解除锁定用户</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以锁定/屏蔽/解除锁定用户，请根据自己的需要合理设置此选项，建议对超级版主及其以上用户组设置此权限">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g51);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以删除用户1－10天内所发帖子
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(29)" type=radio value="1" <%if reGroupSetting(29)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(29)" type=radio value="0"  <%if reGroupSetting(29)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g52" value="<b>可以删除用户1－10天内所发帖子</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以删除用户1－10天内所发帖子，请根据自己的需要合理设置此选项，建议对超级版主及其以上用户组设置此权限">
<td class=Forumrow><a href=# onclick="helpscript(g52);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以查看来访IP及来源
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(30)" type=radio value="1" <%if reGroupSetting(30)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(30)" type=radio value="0"  <%if reGroupSetting(30)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g53" value="<b>可以查看来访IP及来源</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以查看来访IP及来源，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g53);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以限定IP来访
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(31)" type=radio value="1" <%if reGroupSetting(31)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(31)" type=radio value="0"  <%if reGroupSetting(31)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g54" value="<b>可以限定IP来访</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以限定IP来访，请根据自己的需要合理设置此选项，建议对超级版主及其以上用户组设置此权限">
<td class=Forumrow><a href=# onclick="helpscript(g54);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以管理用户权限
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(42)" type=radio value="1" <%if reGroupSetting(42)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(42)" type=radio value="0"  <%if reGroupSetting(42)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g55" value="<b>可以管理用户权限</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以管理用户权限，请根据自己的需要合理设置此选项，建议对超级版主及其以上用户组设置此权限">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g55);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以批量删除帖子（前台）
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(45)" type=radio value="1" <%if reGroupSetting(45)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(45)" type=radio value="0"  <%if reGroupSetting(45)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g56" value="<b>可以批量删除帖子（前台）</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以批量删除帖子（前台），请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=Forumrow><a href=# onclick="helpscript(g56);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>是否有审核帖子的权限
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(36)" type=radio value="1" <%if reGroupSetting(36)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(36)" type=radio value="0" <%if reGroupSetting(36)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g57" value="<b>是否有审核帖子的权限</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否有审核帖子的权限，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g57);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>是否有进入隐含论坛的权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(37)" type=radio value="1"  <%if reGroupSetting(37)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(37)" type=radio value="0" <%if reGroupSetting(37)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g58" value="<b>是否有进入隐含论坛的权限</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否有进入隐含论坛的权限，请根据自己的需要合理设置此选项">
<td class=Forumrow><a href=# onclick="helpscript(g58);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>有论坛文件管理权限
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(48)" type=radio value="1" <%if reGroupSetting(48)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(48)" type=radio value="0" <%if reGroupSetting(48)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g59" value="<b>有论坛文件管理权限</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否有论坛文件管理权限，请根据自己的需要合理设置此选项，建议对版主及其以上用户组设置此权限，相关管理操作在论坛展区">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g59);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<th height="23" colspan="3" align=left>＝＝短信权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以发送短信
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(32)" type=radio value="1"  <%if reGroupSetting(32)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(32)" type=radio value="0" <%if reGroupSetting(32)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g60" value="<b>可以发送短信</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否有可以发送短信的权限，请根据自己的需要合理设置此选项">
<td class=Forumrow><a href=# onclick="helpscript(g60);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>最多发送用户
</td>
<td height="23" width="40%" class=ForumrowHighLight><input name="GroupSetting(33)" size=5 type=text value="<%=reGroupSetting(33)%>"></td>
<input type="hidden" id="g61" value="<b>最多发送用户</b><br><li>在这里您可以根据需要设置不同用户组或等级用户最多发送用户，请根据自己的需要合理设置此选项，建议不要设置过大以免消耗论坛资源。">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g61);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>短信内容大小限制
</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(34)" size=5 type=text value="<%=reGroupSetting(34)%>"> byte</td>
<input type="hidden" id="g62" value="<b>短信内容大小限制</b><br><li>在这里您可以根据需要设置不同用户组或等级用户短信内容大小限制，请根据自己的需要合理设置此选项，建议不要设置过大以免消耗论坛资源.">
<td class=Forumrow><a href=# onclick="helpscript(g62);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>信箱大小限制
</td>
<td height="23" width="40%" class=ForumrowHighLight><input name="GroupSetting(35)" size=5 type=text value="<%=reGroupSetting(35)%>"> KB</td>
<input type="hidden" id="g63" value="<b>信箱大小限制</b><br><li>在这里您可以根据需要设置不同用户组或等级用户信箱大小限制，请根据自己的需要合理设置此选项，建议不要设置过大以免消耗论坛资源">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g63);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>新注册用户多少分钟后才能发短信</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(53)" type=text value="<%=reGroupSetting(53)%>" size=4> 分钟</td>
<input type="hidden" id="g64" value="<b>新注册用户多少分钟后才能发短信</b><br><li>在这里您可以根据需要设置不同用户组或等级新注册用户多少分钟后才能发短信，出于防止恶意群发或使用软件群发短信的目的，建议合理设置此选项">
<td class=Forumrow><a href=# onclick="helpscript(g64);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>一天最多发短信数目</td>
<td height="23" width="40%" class=ForumrowHighLight><input name="GroupSetting(63)" type=text value="<%=reGroupSetting(63)%>" size=4></td>
<input type="hidden" id="g65" value="<b>一天最多发短信数目</b><br><li>在这里您可以根据需要设置不同用户组或等级用户一天最多发短信数目，出于防止恶意群发或使用软件群发短信的目的，建议合理设置此选项">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g65);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<th height="23" colspan="3" align=left>＝＝其他权限</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以搜索论坛
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(14)" type=radio value="1" <%if reGroupSetting(14)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(14)" type=radio value="0" <%if reGroupSetting(14)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g31" value="<b>可以搜索论坛</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以搜索论坛，请根据自己的需要合理设置此选项，建议对未登录用户关闭此选项">
<td class=Forumrow><a href=# onclick="helpscript(g31);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以使用'发送本页给好友'功能
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(15)" type=radio value="1" <%if reGroupSetting(15)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(15)" type=radio value="0" <%if reGroupSetting(15)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g32" value="<b>可以使用'发送本页给好友'功能</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以使用'发送本页给好友'功能，请根据自己的需要合理设置此选项，建议对未登录用户关闭此选项">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g32);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可以修改个人资料
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(16)" type=radio value="1" <%if reGroupSetting(16)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(16)" type=radio value="0" <%if reGroupSetting(16)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g33" value="<b>可以修改个人资料</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以修改个人资料">
<td class=Forumrow><a href=# onclick="helpscript(g33);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>可以浏览论坛事件
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(39)" type=radio value="1"  <%if reGroupSetting(39)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(39)" type=radio value="0" <%if reGroupSetting(39)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g34" value="<b>可以浏览论坛事件</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可以浏览论坛事件，请根据自己的需要合理设置此选项，建议对未登录用户关闭此选项">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g34);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>可浏览论坛展区的权限
</td>
<td height="23" width="40%" class=Forumrow>是<input name="GroupSetting(49)" type=radio value="1"  <%if reGroupSetting(49)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(49)" type=radio value="0" <%if reGroupSetting(49)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g35" value="<b>可浏览论坛展区的权限</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否可浏览论坛展区的权限，请根据自己的需要合理设置此选项，建议对未登录用户关闭此选项">
<td class=Forumrow><a href=# onclick="helpscript(g35);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>是否可以使用签名
</td>
<td height="23" width="40%" class=ForumrowHighLight>是<input name="GroupSetting(55)" type=radio value="1"  <%if reGroupSetting(55)="1" then%>checked<%end if%>>&nbsp;否<input name="GroupSetting(55)" type=radio value="0" <%if reGroupSetting(55)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g36" value="<b>是否可以使用签名</b><br><li>在这里您可以根据需要设置不同用户组或等级用户是否是否可以使用签名，请根据自己的需要合理设置此选项">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g36);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>签名的最大长度</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(56)" type=text value="<%=reGroupSetting(56)%>" size=4> 字节</td>
<input type="hidden" id="g37" value="<b>签名的最大长度</b><br><li>在这里您可以根据需要设置不同用户组或等级用户签名的最大长度，请根据自己的需要合理设置此选项，为了避免影响帖子内容浏览，不建议对此项设置过大字节数">
<td class=Forumrow><a href=# onclick="helpscript(g37);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="点击查阅管理帮助！"></a></td>
</tr>
<tr> 
<td height="23" width="60%" class=Forumrow>
</td>
<td height="23" width="40%" class=Forumrow colspan=2><input type="submit" name="submit" value="提 交"></td>
</tr>
<%
End Function

Function GetGroupPermission()
	Dim i,TempSetting
	For i = 0 To 90
		If Trim(Request.Form("GroupSetting("&i&")"))="" Then
			TempSetting = 0
		Else
			TempSetting = Replace(Trim(Request.Form("GroupSetting("&i&")")),",","")
		End If
		If i = 0 Then
			GetGroupPermission = TempSetting
		ElseIf i = 58 Then
			GetGroupPermission = GetGroupPermission & "," & Replace(Trim(Request.Form("GroupSetting("&i&")A")),",","") & "§" & Replace(Trim(Request.Form("GroupSetting("&i&")B")),",","")
		Else
			GetGroupPermission = GetGroupPermission & "," & TempSetting
		End If
	Next
	GetGroupPermission = Replace(GetGroupPermission,"'","''")
End Function
%>