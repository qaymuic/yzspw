<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
dim admin_flag
admin_flag=",8,"
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
else
call gradeinfo()
end if
end sub

sub gradeinfo()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<th width="100%" class="tableHeaderText" colspan=2>论坛门派管理
</th>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow><b>用户门派管理</b>：您可以添加修改或者删除论坛门派。</td>
</tr>
<%if request("action")="edit" then%>
<form method="POST" action=admin_menpai.asp?action=savedit>
<%
	set rs=Dvbbs.Execute("select * from dv_GroupName where id="&request("id"))
%>
<tr> 
<th width="100%" id=TableTitleLink colspan=2>修改门派 | <a href=admin_menpai.asp><b>添加门派<b></a></th>
</tr>
<tr> 
<td width="30%" class=forumrow><b>门派名称</b></td>
<td width="70%" class=forumrow> 
<input type="text" name="Groupname" size="35" value="<%=rs("Groupname")%>">&nbsp;<input type="submit" name="Submit" value="提 交">
<input type=hidden name=id value="<%=request("id")%>">
</td>
</tr>
<%set rs=nothing%>
<%else%>
<form method="POST" action=admin_menpai.asp?action=save>
<tr> 
<th width="100%" class="tableHeaderText" colspan=2><b>添加门派</b>
</th>
</tr>
<tr> 
<td width="30%" class=forumrow><b>门派</B>名称</td>
<td width="70%" class=forumrow>
<input type="text" name="Groupname" size="35">&nbsp;<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</form>
<%end if%>
<tr> 
<th height="23" colspan="2" ><b>管理门派</b></th>
</tr>
<%
	set rs=Dvbbs.Execute("select * from dv_GroupName")
	do while not rs.eof
%>
<tr> 
<td height="23" colspan="2" class=forumRowHighlight>
<a href="admin_menpai.asp?id=<%=rs("id")%>&action=edit">修改</a> | <a href="admin_menpai.asp?id=<%=rs("id")%>&action=del">删除</a> | <%=rs("GroupName")%>
</td>
</tr>
<%
	rs.movenext
	loop
	set rs=nothing
%>
</table>
<%
end sub
sub savegroup()
dim GroupName
GroupName=Dvbbs.Checkstr(trim(request("GroupName")))
set rs=Dvbbs.Execute("select top 1 id from dv_GroupName where   GroupName='"&GroupName&"' order by id desc")
if rs.eof and rs.bof then
	Dvbbs.Execute("insert into dv_GroupName (GroupName) values ('"&GroupName&"')")
else
	Errmsg=ErrMsg + "<BR><li>不能添加相同的组名。"
	dvbbs_error()
	exit sub
end if
set rs=nothing

%>
<center><p><b>添加成功！</b>
<%
end sub
sub savedit()
dim GroupName
GroupName=Dvbbs.Checkstr(trim(request("GroupName")))
set rs=Dvbbs.Execute("select top 1 id from dv_GroupName where   GroupName='"&GroupName&"' order by id desc")
if rs.eof and rs.bof then
		Dvbbs.Execute("update dv_GroupName set GroupName='"&GroupName&"' where id="&Dvbbs.Checkstr(request("id")))
else
	Errmsg=ErrMsg + "<BR><li>不能修改成已存在相同的组名。"
	dvbbs_error()
	exit sub
end if
set rs=nothing
%>
<center><p><b>修改成功！</b>
<%
end sub
sub del()
	Dvbbs.Execute("delete from dv_GroupName where id="&Dvbbs.Checkstr(request("id")))
%>
<center><p><b>删除成功！</b>
<%
end sub

%>