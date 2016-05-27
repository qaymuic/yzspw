<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<!-- #include file="inc/GroupPermission.asp" -->

<%
	Head()
	dim admin_flag
	admin_flag=",17,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		dvbbs_error()
	Else
		if request("action")="save" then
		call savegrade()
		elseif request("action")="add" then
		call add()
		elseif request("action")="savenew" then
		call savenew()
		elseif request("action")="del" then
		call del()
		elseif request("action")="per" then
		call per()
		else
		call gradeinfo()
		end if
		Footer()
	End If


sub gradeinfo()
%>
<form method="POST" action=admin_grade.asp?action=save>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="5" class=forumrowHighlight>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr>
<td valign=top>
<B>关于用户等级设置的说明，请仔细阅读后做设置</B>：<BR>
相关用户组如果无对应等级名称，则注册用户自动按照文章升级<BR>
相关用户组的等级名称可以和用户组名不一样<BR>
各个等级可设定不自定义权限，权限类型和组权限一样。如果某个等级设定了自定义权限，这样该等级在论坛所有版面就有了自定义的权限，而且该等级将出现在版面权限定义的组菜单中（这样就可设定某个等级在某个版面中的权限），而且相关等级用户的组ID将变成等级ID
</td>
<td width="50%" valign=top>
<B>在等级中设定用户组有什么用？</B><BR>
一般来说，只有注册用户拥有等级，所以在等级所属组中一般都设定对应注册用户组，如果设置成别的组，那么该用户在升级到这个等级的同时也将自动归入所设置的组<BR>
比如你新添加了一个用户组，并且给予了这个用户组某一些权限，那么你可以设置达到一定等级（帖子）的用户自动更新到这个用户组以使用这个用户组的权限。<BR>如果您想某个等级的用户不跟随帖子数而上升等级，那么就把最少发贴设置为<B>-1</B>，一般为特殊用户组需要这样的设置，设置某个级别最少发贴为<B>-1</B>后，该级别的用户将不能根据帖子增加而升级，别的用户也不能自动升级到该级别，只有在用户管理中方能更改其级别
</td>
</tr>
</table>
</td>
</tr>
<tr> 
<th height="23" colspan="5" >用户等级设定</th>
</tr>
<tr> 
<td width="25%" class=forumrowHighlight><B>等级名称</B></td>
<td width="15%" class=forumrowHighlight><B>最少发贴</B></td>
<td width="20%" class=forumrowHighlight><B>图片</B></td>
<td width="25%" class=forumrowHighlight><B>所属组</B></td>
<td width="15%" class=forumrowHighlight><B>操作</B></td>
</tr>
<%
Dim TempArray,DefaultLock
Set Rs=Dvbbs.Execute("select UserGroupID,title from dv_usergroups where issetting=1 And ParentGID=0 order by UserGroupID")
TempArray = Rs.GetRows(-1)
set rs=Dvbbs.Execute("select * from dv_usergroups order by ParentGID,UserGroupID,minarticle desc")
do while not rs.eof
	If Rs("ParentGID")=0 Then 
		DefaultLock="1"
	Else
	DefaultLock=""
	End If
	%>
	<input type=hidden value="<%=rs("UserGroupID")%>" name="usertitleid">
	<tr> 
	<td width="25%" class=Forumrow><input size=15 value="<%=rs("usertitle")%>" name="usertitle" type=text></td>
	<td width="15%" class=Forumrow>
	<%If DefaultLock <>"1" Then %>
		<input size=5 value="<%=rs("MinArticle")%>" name="minarticle" type=text >
	<%Else%>
		<input type=hidden   value="<%=rs("MinArticle")%>" name="minarticle"  >
		<%=rs("MinArticle")%>
	<%End If%>
	</td>
	<td width="20%" class=Forumrow><input size=15 value="<%=rs("grouppic")%>" name="titlepic" type=text></td>
	<td width="25%" class=Forumrow>
	<%If DefaultLock <>"1" Then %>
		<select name="groupid" size=1>
		<%For i=0 To Ubound(TempArray,2)%>
			<%If Rs("ParentGID")=0 Then%>
				<option value="<%=TempArray(0,i)%>" <%If Rs("UserGroupID")=TempArray(0,i) Then Response.Write "selected"%>><%=TempArray(1,i)%></option>
			<%Else%>
			<option value="<%=TempArray(0,i)%>" <%If Rs("ParentGID")=TempArray(0,i) Then Response.Write "selected"%>><%=TempArray(1,i)%></option>
			<%End If%>
		<%Next%>
		</select>
	<%Else
		Response.Write "<input type=hidden   value="""&Rs("UserGroupID")&""" name=""groupid""  >"
		For i=0 To Ubound(TempArray,2)
			If Rs("UserGroupID")=TempArray(0,i) Then
				Response.Write TempArray(1,i)
			End If
		Next		
	End If%>
	</td>
	<td width="15%" class=Forumrow><%If Rs("UserGroupID")>8 Then%><a href="?action=del&id=<%=rs("UserGroupID")%>">删除</a> | <%End If%>
	<%
	If Rs("ParentGID")=0 Then
		Response.Write "<a href=admin_group.asp?action=editgroup&groupid="&Rs("UserGroupID")&">"
	Else
		If Rs("IsSetting")=1 Then
			Response.Write "<a href=admin_grade.asp?action=per&groupid="&Rs("UserGroupID")&"&regroupid="&rs("UserGroupID")&">"
		Else
			Response.Write "<a href=admin_grade.asp?action=per&groupid="&Rs("ParentGID")&"&regroupid="&rs("UserGroupID")&">"
		End If
	End If
	%>权限</a>
	<%If Rs("UserGroupID")>8 And Rs("IsSetting")=1 Then Response.Write " <font color=red>自</font>"%>
	</td>
	</tr>
	<%
	rs.movenext
	loop
rs.close
set rs=nothing
%>
<tr> 
<td width="100%" colspan=5 class=Forumrow> 
<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</table>
</form>
<%
end sub

Sub savegrade()
	Server.ScriptTimeout=99999999
	Dim usertitleid,iuserclass,usertitle,Minarticle,titlepic,groupid
	For i=1 to request.form("usertitleid").count
		usertitleid=replace(request.form("usertitleid")(i),"'","")
		usertitle=replace(request.form("usertitle")(i),"'","")
		minarticle=replace(request.form("minarticle")(i),"'","")
		titlepic=replace(request.form("titlepic")(i),"'","")
		groupid=replace(request.form("groupid")(i),"'","")
		if isnumeric(usertitleid) and isnumeric(iuserclass) and usertitle<>"" and isnumeric(minarticle) and titlepic<>"" and isnumeric(groupID) then
		set rs=Dvbbs.Execute("select * from dv_usergroups where UserGroupID="&usertitleID)
		if rs("usertitle")<>trim(usertitle) or rs("grouppic")<>trim(titlepic) or (rs("parentgid")<>cint(groupid) and rs("parentgid")>0) then
			'如果有自定义权限，则UserGroupID为等级所在的ID，反之则为组ID
			if rs("issetting")=1 then groupid=rs("usergroupid")
			Dvbbs.Execute("update [dv_user] set userclass='"&usertitle&"',titlepic='"&titlepic&"',usergroupid="&groupid&" where userclass='"&rs("usertitle")&"'")
		end if
		if rs("parentgid")=0 then groupid=0
		Dvbbs.Execute("update dv_usergroups set usertitle='"&usertitle&"',minarticle="&minarticle&",grouppic='"&titlepic&"',parentgid="&groupid&" where usergroupid="&usertitleID)
		end if
	next
	response.write "设置成功，请返回。"
	set rs=nothing
End Sub

sub add()
%>
<form method="POST" action=admin_grade.asp?action=savenew>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th colspan="2">添加新的用户等级</th>
</tr>
<tr>
<td width="40%" class=forumrow><B>所属用户组</B></td>
<td width="60%" class=forumrow>
<select size=1 name="usergroupid">
<%
set rs=Dvbbs.Execute("select * from dv_usergroups where parentgid=0 order by usergroupid")
do while not rs.eof
%>
<option value="<%=rs("usergroupid")%>" <%if rs("usergroupid")=4 then%>selected<%end if%>><%=rs("title")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width="40%" class=forumrow><B>等级名称</B></td>
<td width="60%" class=forumrow><input size=30 name="usertitle" type=text></td>
</tr>
<tr>
<td width="40%" class=forumrow><B>最少发贴</B><BR>如果该等级是荣誉称号或者管理身份，这里可以填写-1，表示不跟随帖子增长而升级</td>
<td width="60%" class=forumrow><input size=30 name="minarticle" type=text></td>
</tr>
<tr>
<td width="40%" class=forumrow><B>等级图片</B></td>
<td width="60%" class=forumrow><input size=30 name="titlepic" type=text>&nbsp;这将体现在帖子内容显示左边的用户资料中</td>
</tr>
<tr> 
<td width="100%" colspan=2 class=forumrow> 
<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</table>
</form>
<%
end sub
sub savenew()
if request.form("minarticle")="" then
	Errmsg=ErrMsg + "<BR><li>请输入新的等级需要文章数。"
	dvbbs_error()
	exit sub
elseif not isnumeric(request.form("minarticle")) then
	Errmsg=ErrMsg + "<BR><li>新的等级文章数只能是数字。"
	dvbbs_error()
	exit sub
end if
if request.form("titlepic")="" then
	Errmsg=ErrMsg + "<BR><li>请输入新的等级图片。"
	dvbbs_error()
	exit sub
end if
if request.form("usertitle")="" then
	Errmsg=ErrMsg + "<BR><li>请输入新的等级名称。"
	dvbbs_error()
	exit sub
end if
Dim GroupTitle,GroupSetting,GroupPic
Set rs=dvbbs.execute("select * from dv_usergroups where usergroupid="&request.form("usergroupid"))
GroupTitle=rs("title")
GroupSetting=rs("GroupSetting")
GroupPic=rs("titlepic")
set rs = server.CreateObject ("Adodb.recordset")
sql="select * from dv_usergroups where usertitle='"&request.form("usertitle")&"'"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
rs.addnew
rs("usertitle")=request.form("usertitle")
rs("minarticle")=request.form("minarticle")
rs("grouppic")=request.form("titlepic")
rs("parentgid")=request.form("usergroupid")
rs("title")=GroupTitle
rs("GroupSetting")=GroupSetting
rs("isdisp")=0
rs("IsSetting")=0
rs("titlepic")=GroupPic
rs.update
else
	Errmsg=ErrMsg + "<BR><li>该等级名称已经存在。"
	dvbbs_error()
	exit sub
end if
rs.close
set rs=nothing
response.write "添加成功！建议您到更新用户数据中进行更新操作！"
end sub

Sub Del()
	Server.ScriptTimeout = 99999999
	Dim Minarticle, Minuserclass
	If Isnumeric(Request("Id")) Then
		If CLng(Request("id")) < 9 Then
			Errmsg = ErrMsg + "<BR><li>系统默认等级不能删除。"
			Dvbbs_Error()
			Exit Sub
		End If
		Set Rs = Dvbbs.Execute("SELECT * FROM Dv_UserGroups WHERE UserGroupId = " & Request("id"))
		Minarticle = Rs("Minarticle")
		Minuserclass = Rs("Usertitle")
		Rem 修正删除等级后等级设置的错误 2004-5-1 Dvbbs.YangZheng
		Set Rs = Dvbbs.Execute("SELECT TOP 1 * FROM Dv_Usergroups WHERE ParentGId = " & Request("id") & " AND NOT MinArticle = -1 ORDER BY Minarticle")
		If Not (Rs.Eof And Rs.Bof) Then
			Dvbbs.Execute("UPDATE [Dv_User] SET Userclass = '" & Rs("Usertitle") & "', Titlepic = '" & Rs("Grouppic") & "' WHERE Userclass = '" & Minuserclass & "'")
		Else
			Set Rs = Nothing
			Set Rs = Dvbbs.Execute("SELECT TOP 1 * FROM Dv_UserGroups WHERE ParentGId = 4 ORDER By Minarticle Desc")
			If Not (Rs.Eof And Rs.Bof) Then
				Dvbbs.Execute("UPDATE [Dv_User] SET UserGroupId = 4, Userclass = '" & Rs("Usertitle") & "', Titlepic = '" & Rs("Grouppic") & "' WHERE Userclass = '" & Minuserclass & "'")
			End If
		End If
		Dvbbs.Execute("DELETE FROM Dv_Usergroups WHERE Usergroupid = " & Request("id"))
		Response.Write "删除成功！"
		Set Rs = Nothing
	End If
End Sub

sub per()
if not isnumeric(request("groupid")) then
response.write "错误的参数！"
exit sub
end if
if request("groupaction")="yes" then
	dim GroupSetting,groupid,newgroupsetting
	GroupSetting=GetGroupPermission
	if request("isdefault")=1 then
		set rs=dvbbs.execute("select * from dv_usergroups where usergroupid="&request("groupid"))
		If Rs("ParentGID")=0 Then
			Dv_suc("您没有选择自定义等级选项，所做修改将无效")
			Exit Sub
		End If
		if rs("issetting")=1 then
		groupid=rs("parentgid")
		set rs=nothing
		set rs=dvbbs.execute("select * from dv_usergroups where usergroupid="&groupid)
		newgroupsetting=rs("groupsetting")
		Set Rs=Nothing
		dvbbs.execute("update dv_usergroups set issetting=0,groupsetting='"&newgroupsetting&"' where usergroupid="&request("regroupid"))
		'取消自定义设置，更新用户数据，更新为用户组ID
		dvbbs.execute("update [dv_user] set usergroupid="&groupid&" where userclass='"&request("usertitle")&"'")
		end if
		
	else
		dvbbs.execute("update dv_usergroups set issetting=1,groupsetting='"&GroupSetting&"' where usergroupid="&request("regroupid"))
		'更新用户数据
		dvbbs.execute("update [dv_user] set usergroupid="&request("regroupid")&" where userclass='"&request("usertitle")&"'")
	End If

	ReloadGroup(request("regroupid"))
	Dv_suc("修改等级自定义权限成功")
else
Dim reGroupSetting,founduserper,usergrade
If IsNumerIc(request("regroupid")) and request("regroupid")<>"" Then
	Set Rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&request("regroupid"))
	usergrade=rs("usertitle")
End If
founduserper=false
set rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "未找到用户等级"
exit sub
end if
If Rs("UserGroupID")<9 Then
	founduserper=false
Else
	If Rs("IsSetting")=1 Then
		founduserper=true
	Else
		founduserper=false
	End If
End If
reGroupSetting=split(rs("GroupSetting"),",")
%>
<FORM METHOD=POST ACTION="?action=per">
<input type=hidden name="groupid" value="<%=request("groupid")%>">
<input type=hidden name="regroupid" value="<%=request("regroupid")%>">
<input type=hidden name="usertitle" value="<%=usergrade%>">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><td colspan=3 height=25 class="Forumrow"><B>说明</B>：<BR>在这里您可以设置各个等级在论坛中的默认权限，<font color=blue>默认为使用该等级所属用户组权限，如果要让该等级有自定义权限，则修改时请选择自定义设置选项</font></td></tr>
<tr> 
<th height="23" colspan="3" >编辑论坛用户等级权限&nbsp;>> <%=rs("usertitle")%><%if usergrade<>"" then Response.Write "&nbsp;>> "&usergrade&""%></th>
</tr>
<tr> 
<td height="23" colspan="3" class=forumrow><input type=radio name="isdefault" value="1" <%if not founduserper then%>checked<%end if%>><B>使用用户组默认值</B> (注意: 这将删除任何之前所做的自定义设置)</td>
</tr>
<tr> 
<td height="23" colspan="3"  class=forumrow><input type=radio name="isdefault" value="0" <%if founduserper then%>checked<%end if%>><B>使用自定义设置</B>&nbsp;(选择自定义才能使以下设置生效) </td>
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

Function ReloadGroup(UserGroupID)
	Dim Rs,SQL
	SQL = "Select GroupSetting From [Dv_UserGroups] where UserGroupID = " & UserGroupID & ""
	Set Rs = Dvbbs.Execute(SQL)
	Dvbbs.Name="GroupSetting_"& UserGroupID
	Dvbbs.value=Rs(0)
	Set Rs = Nothing
End Function
%>
