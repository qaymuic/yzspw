<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<!-- #include file="inc/GroupPermission.asp" -->
<%
	Head()
	dim admin_flag
	admin_flag=",15,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		dvbbs_error()
	else
		call main()
		Footer()
	end if

	sub main()
if request("action")="save" then
call savegroup()
elseif request("action")="savedit" then
call savedit()
elseif request("action")="del" then
call del()
elseif request("action")="group" then
call gradeinfo()
elseif request("action")="addgroup" then
call addgroup()
elseif request("action")="editgroup" then
call editgroup()
elseif request("action")="delgroup" then
call delgroup()
elseif request("action")="saveorders" then
call saveorders()
else
call usergroup()
end if
end sub

sub usergroup()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="5" >用户组管理&nbsp;&nbsp;<a href="?action=addgroup"><font color=#FFFFFF>[添加用户组]</font></a></th>
</tr>
<tr align=center>
<td height="23" width="30%" class=forumHeaderBackgroundAlternate><B>用户组(等级名称)</B></td>
<td height="23" width="15%" class=forumHeaderBackgroundAlternate><B>用户数量</B></td>
<td height="23" width="20%" class=forumHeaderBackgroundAlternate><B>编辑权限</B></td>
<td height="23" width="15%" class=forumHeaderBackgroundAlternate><B>性质</B></td>
<td height="23" width="20%" class=forumHeaderBackgroundAlternate><B>列出用户</B></td>
</tr>
<%
dim trs
set rs=Dvbbs.Execute("select * from dv_usergroups where issetting=1 And ParentGID=0 order by UserGroupID")
do while not rs.eof
set trs=Dvbbs.Execute("select count(*) from [dv_user] where UserGroupID="&rs("UserGroupID"))
%>
<tr align=center>
<td height="23" width="30%" class="Forumrow"><%=rs("title")%>(<font color=gray><%=rs("usertitle")%></font>)</td>
<td height="23" width="15%" class="Forumrow"><%=trs(0)%></td>
<td height="23" width="20%" class="Forumrow"><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">编辑</a><%if rs("UserGroupID")>8 then%> | <a href="?action=delgroup&groupid=<%=rs("UserGroupID")%>">删除</a><%end if%></td>
<td height="23" width="15%" class="Forumrow">
<%
If Rs("IsSetting")=1 Then
	If Rs("ParentGID")=0 Then
		Response.Write "默认组"
	Else
		Response.Write "自定义等级"
	End If
End If
%></td>
<td height="23" width="20%" class="Forumrow"><a href="admin_user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">列出所有用户</a></td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
<tr><td colspan=5 height=25 class="ForumrowHighlight"><B>说明</B>：<BR>①在这里您可以设置各个用户组在论坛中的默认权限，论坛默认用户组不能删除和编辑用户组名<BR>②可以进行添加用户组操作并设置其权限，可以将其他组用户转移到该组，请到用户管理中进行相关操作，对某个组用户进行管理请点击该组用户列表<BR>③可以删除和编辑新添加的用户组<BR>④<B>如果删除用户组，则该用户组所包含的用户将自动转到注册用户组，同时删除在等级中和该用户组关联的等级，并更新该用户组所包含的用户的等级为注册用户组按照文章计算的等级</B><BR>⑥各个组在<font color=blue>论坛首页以及在线列表中的图例</font>，请直接点击编辑进入修改页面</td></tr>
<tr> 
<th height="23" colspan="5" >用户组排序</th>
<FORM METHOD=POST ACTION="admin_group.asp?action=saveorders">
<tr><td colspan=5 height=25 class="Forumrow">
<table width=100% >
说明：该排序主要用于论坛首页的在线图例显示顺序
<%
set rs=Dvbbs.Execute("select * from dv_usergroups where issetting=1 and isdisp=1 order by orders")
do while not rs.eof
%>
<tr><td width=200><%=rs("title")%></td><td><input type=text size=5 name="orders" value="<%=rs("orders")%>"></td></tr>
<input type=hidden value="<%=rs("usergroupid")%>" name="groupid">
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</table>
</td></tr>
<tr><td colspan=5 height=25 class="Forumrowhighlight">
<input type=submit name=submit value="提 交">
</td></tr>
</FORM>
</tr>
</table>
<%
end sub

sub saveorders()
dim orders
for i=1 to request.form("groupid").count
orders=request.form("orders")(i)
Dvbbs.Execute("update Dv_usergroups set orders="&orders&" where usergroupid="&request.form("groupid")(i))
next
dv_suc("更新用户组排序成功")
ReloadGroupTitle()
end sub

Sub Delgroup()
	If Not Isnumeric(Request("Groupid")) Then
		Response.Write "错误的参数！"
		Exit Sub
	End If
	If Clng(Request("Groupid")) < 9 Then
		Response.Write "系统默认的用户组不允许删除。"
		Exit Sub
	End If
	'更新用户等级数据
	Server.ScriptTimeout = 999999
	Dim UserGrade
	Set Rs = Dvbbs.Execute("SELECT Userid, UserPost FROM [Dv_User] WHERE UserGroupID = " & Request("Groupid"))
	Do While Not Rs.Eof
		Rem 读取注册用户组对应文章数的等级 2004-5-1 Dvbbs.YangZheng
		Set UserGrade = Dvbbs.Execute("SELECT TOP 1 Usertitle, Grouppic, UserGroupID From Dv_Usergroups WHERE ParentGID = 4 AND NOT MinArticle = -1 ORDER BY MinArticle")
		Dvbbs.Execute("UPDATE [Dv_User] SET Userclass = '" & UserGrade(0) & "', Titlepic = '" & UserGrade(1) & "', UserGroupid = 4 WHERE Userid = " & Rs(0))
		Rs.Movenext
	Loop
	Set Rs = Nothing
	'删除用户组
	Dvbbs.Execute("DELETE FROM Dv_Usergroups WHERE UserGroupID = " & Request("Groupid"))
	'删除其关联等级
	Dvbbs.Execute("DELETE FROM Dv_Usergroups WHERE ParentGID = " & Request("Groupid"))
	Response.Write "删除成功！"
	ReloadGroupTitle()
End Sub

sub editgroup()
if not isnumeric(request("groupid")) then
response.write "错误的参数！"
exit sub
end if
if request("groupaction")="yes" then
	dim GroupSetting
	if request.form("title")="" then
	response.write "请输入用户组名称！"
	exit sub
	end if
	GroupSetting=GetGroupPermission
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from dv_usergroups where usergroupid="&request("groupid")
	rs.open sql,conn,1,3
	rs("title")=request.form("title")
	rs("GroupSetting")=GroupSetting
	rs("isdisp")=request("isdisp")
	rs("titlepic")=request("grouppic")
	rs.update
	rs.close
	set rs=nothing
	response.write "修改成功！"
	ReloadGroup(request("groupid"))
else
Dim reGroupSetting
set rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "未找到该用户组！"
exit sub
end if
reGroupSetting=split(rs("GroupSetting"),",")
%>
<FORM METHOD=POST ACTION="?action=editgroup">
<input type=hidden name="groupid" value="<%=request("groupid")%>">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><td colspan=3 height=25 class="Forumrow"><B>说明</B>：<BR>①在这里您可以设置各个用户组在论坛中的默认权限，论坛默认用户组不能删除和编辑用户组名<BR>②可以删除和编辑新添加的用户组</td></tr>
<tr> 
<th height="23" colspan="3">编辑用户组：<%=rs("title")%></th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>用户组名称</td>
<td height="23" width="40%" class=Forumrow colspan=2><input size=35 name="title" type=text value="<%=rs("title")%>"  <%if Cint(request("GroupID"))<9 then%>disabled<%end if%>></td>
</tr>
<%if Cint(request("GroupID"))<9 then%>
<input name="title" type=hidden value="<%=rs("title")%>">
<%end if%>
<tr>
<td height="23" width="60%" class=Forumrow>是否在首页的在线图例中显示</td>
<td height="23" width="40%" class=Forumrow colspan=2>是<input name="isdisp" type=radio value="1" <%if rs("isdisp")=1 then%>checked<%end if%>>&nbsp;否<input name="isdisp" type=radio value="0" <%if rs("isdisp")=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>首页在线图例图片地址</td>
<td height="23" width="40%" class=Forumrow colspan=2><input size=35 name="grouppic" type=text value="<%=rs("titlepic")%>"></td>
</tr>
<%
GroupPermission(rs("GroupSetting"))
%>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
set rs=nothing
end if
end sub

sub addgroup()
if request("groupaction")="yes" then
	dim GroupSetting
	if request.form("title")="" then
	response.write "请输入用户组名称！"
	exit sub
	end if
	GroupSetting=GetGroupPermission
	'response.write len(GroupSetting)
	'response.end
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from dv_usergroups where title='"&request.form("title")&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		rs.addnew
		rs("title")=request.form("title")
		rs("GroupSetting")=GroupSetting
		rs("isdisp")=request("isdisp")
		rs("IsSetting")=1
		rs("minarticle")=-1
		rs("parentgid")=0
		rs("grouppic")="level10.gif"
		rs("usertitle")=request.form("title")
		rs("titlepic")=request("grouppic")
		rs.update
	else
		Errmsg=ErrMsg + "<BR><li>该用户组名称已经存在。"
		dvbbs_error()
		exit sub
	end if
	rs.close
	set rs=nothing
	response.write "添加用户组成功，用户组名称同时为等级名称，新等级采用了默认的图片设置，您可以到等级管理中进行修改！"
	Dim MaxID
	Set Rs=Dvbbs.Execute("Select Top 1 * from Dv_UserGroups Order by UserGroupID Desc")
	MaxID=rs(0)
	Set Rs=Nothing
	ReloadGroup(MaxID)
else
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<FORM METHOD=POST ACTION="?action=addgroup">
<tr><td colspan=3 height=25 class="Forumrow"><B>说明</B>：<BR>①可以进行添加用户组操作并设置其权限，可以将其他组用户转移到该组，请到用户管理中进行相关操作，对某个组用户进行管理请点击该组用户列表<BR>②可以删除和编辑新添加的用户组<BR>③<B>添加用户组后，该用户组名同时为等级名称</B></td></tr>
<tr> 
<th height="23" colspan="3" >添加新的用户组</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>用户组名称</td>
<td height="23" width="40%" class=Forumrow colspan=2><input size=35 name="title" type=text></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>是否在首页的在线图例中显示</td>
<td height="23" width="40%" class=Forumrow colspan=2>是<input name="isdisp" type=radio value="1">&nbsp;否<input name="isdisp" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>首页在线图例图片地址</td>
<td height="23" width="40%" class=Forumrow colspan=2><input size=35 name="grouppic" type=text></td>
</tr>
<%
GroupPermission("")
%>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
end if
end sub

Function ReloadGroup(UserGroupID)
	Dim Rs,SQL
	SQL = "Select GroupSetting From [Dv_UserGroups] where UserGroupID = " & UserGroupID & ""
	Set Rs = Dvbbs.Execute(SQL)
	Dvbbs.Name="GroupSetting_"& UserGroupID
	Dvbbs.Value=Rs(0)
	Set Rs = Nothing
	ReloadGroupTitle
End Function

Function ReloadGroupTitle()
	Dvbbs.DelCahe "GroupTitle"
	Dvbbs.DelCahe "GetGroupTitlePic"
End Function
%>