<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
dim admin_flag
admin_flag=",8,"
if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
<th width="100%" class="tableHeaderText" colspan=2>��̳���ɹ���
</th>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow><b>�û����ɹ���</b>������������޸Ļ���ɾ����̳���ɡ�</td>
</tr>
<%if request("action")="edit" then%>
<form method="POST" action=admin_menpai.asp?action=savedit>
<%
	set rs=Dvbbs.Execute("select * from dv_GroupName where id="&request("id"))
%>
<tr> 
<th width="100%" id=TableTitleLink colspan=2>�޸����� | <a href=admin_menpai.asp><b>�������<b></a></th>
</tr>
<tr> 
<td width="30%" class=forumrow><b>��������</b></td>
<td width="70%" class=forumrow> 
<input type="text" name="Groupname" size="35" value="<%=rs("Groupname")%>">&nbsp;<input type="submit" name="Submit" value="�� ��">
<input type=hidden name=id value="<%=request("id")%>">
</td>
</tr>
<%set rs=nothing%>
<%else%>
<form method="POST" action=admin_menpai.asp?action=save>
<tr> 
<th width="100%" class="tableHeaderText" colspan=2><b>�������</b>
</th>
</tr>
<tr> 
<td width="30%" class=forumrow><b>����</B>����</td>
<td width="70%" class=forumrow>
<input type="text" name="Groupname" size="35">&nbsp;<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</form>
<%end if%>
<tr> 
<th height="23" colspan="2" ><b>��������</b></th>
</tr>
<%
	set rs=Dvbbs.Execute("select * from dv_GroupName")
	do while not rs.eof
%>
<tr> 
<td height="23" colspan="2" class=forumRowHighlight>
<a href="admin_menpai.asp?id=<%=rs("id")%>&action=edit">�޸�</a> | <a href="admin_menpai.asp?id=<%=rs("id")%>&action=del">ɾ��</a> | <%=rs("GroupName")%>
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
	Errmsg=ErrMsg + "<BR><li>���������ͬ��������"
	dvbbs_error()
	exit sub
end if
set rs=nothing

%>
<center><p><b>��ӳɹ���</b>
<%
end sub
sub savedit()
dim GroupName
GroupName=Dvbbs.Checkstr(trim(request("GroupName")))
set rs=Dvbbs.Execute("select top 1 id from dv_GroupName where   GroupName='"&GroupName&"' order by id desc")
if rs.eof and rs.bof then
		Dvbbs.Execute("update dv_GroupName set GroupName='"&GroupName&"' where id="&Dvbbs.Checkstr(request("id")))
else
	Errmsg=ErrMsg + "<BR><li>�����޸ĳ��Ѵ�����ͬ��������"
	dvbbs_error()
	exit sub
end if
set rs=nothing
%>
<center><p><b>�޸ĳɹ���</b>
<%
end sub
sub del()
	Dvbbs.Execute("delete from dv_GroupName where id="&Dvbbs.Checkstr(request("id")))
%>
<center><p><b>ɾ���ɹ���</b>
<%
end sub

%>