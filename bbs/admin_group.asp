<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<!-- #include file="inc/GroupPermission.asp" -->
<%
	Head()
	dim admin_flag
	admin_flag=",15,"
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
<th height="23" colspan="5" >�û������&nbsp;&nbsp;<a href="?action=addgroup"><font color=#FFFFFF>[����û���]</font></a></th>
</tr>
<tr align=center>
<td height="23" width="30%" class=forumHeaderBackgroundAlternate><B>�û���(�ȼ�����)</B></td>
<td height="23" width="15%" class=forumHeaderBackgroundAlternate><B>�û�����</B></td>
<td height="23" width="20%" class=forumHeaderBackgroundAlternate><B>�༭Ȩ��</B></td>
<td height="23" width="15%" class=forumHeaderBackgroundAlternate><B>����</B></td>
<td height="23" width="20%" class=forumHeaderBackgroundAlternate><B>�г��û�</B></td>
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
<td height="23" width="20%" class="Forumrow"><a href="?action=editgroup&groupid=<%=rs("UserGroupID")%>">�༭</a><%if rs("UserGroupID")>8 then%> | <a href="?action=delgroup&groupid=<%=rs("UserGroupID")%>">ɾ��</a><%end if%></td>
<td height="23" width="15%" class="Forumrow">
<%
If Rs("IsSetting")=1 Then
	If Rs("ParentGID")=0 Then
		Response.Write "Ĭ����"
	Else
		Response.Write "�Զ���ȼ�"
	End If
End If
%></td>
<td height="23" width="20%" class="Forumrow"><a href="admin_user.asp?action=userSearch&userSearch=10&usergroupid=<%=rs("usergroupid")%>">�г������û�</a></td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
<tr><td colspan=5 height=25 class="ForumrowHighlight"><B>˵��</B>��<BR>�����������������ø����û�������̳�е�Ĭ��Ȩ�ޣ���̳Ĭ���û��鲻��ɾ���ͱ༭�û�����<BR>�ڿ��Խ�������û��������������Ȩ�ޣ����Խ��������û�ת�Ƶ����飬�뵽�û������н�����ز�������ĳ�����û����й������������û��б�<BR>�ۿ���ɾ���ͱ༭����ӵ��û���<BR>��<B>���ɾ���û��飬����û������������û����Զ�ת��ע���û��飬ͬʱɾ���ڵȼ��к͸��û�������ĵȼ��������¸��û������������û��ĵȼ�Ϊע���û��鰴�����¼���ĵȼ�</B><BR>�޸�������<font color=blue>��̳��ҳ�Լ������б��е�ͼ��</font>����ֱ�ӵ���༭�����޸�ҳ��</td></tr>
<tr> 
<th height="23" colspan="5" >�û�������</th>
<FORM METHOD=POST ACTION="admin_group.asp?action=saveorders">
<tr><td colspan=5 height=25 class="Forumrow">
<table width=100% >
˵������������Ҫ������̳��ҳ������ͼ����ʾ˳��
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
<input type=submit name=submit value="�� ��">
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
dv_suc("�����û�������ɹ�")
ReloadGroupTitle()
end sub

Sub Delgroup()
	If Not Isnumeric(Request("Groupid")) Then
		Response.Write "����Ĳ�����"
		Exit Sub
	End If
	If Clng(Request("Groupid")) < 9 Then
		Response.Write "ϵͳĬ�ϵ��û��鲻����ɾ����"
		Exit Sub
	End If
	'�����û��ȼ�����
	Server.ScriptTimeout = 999999
	Dim UserGrade
	Set Rs = Dvbbs.Execute("SELECT Userid, UserPost FROM [Dv_User] WHERE UserGroupID = " & Request("Groupid"))
	Do While Not Rs.Eof
		Rem ��ȡע���û����Ӧ�������ĵȼ� 2004-5-1 Dvbbs.YangZheng
		Set UserGrade = Dvbbs.Execute("SELECT TOP 1 Usertitle, Grouppic, UserGroupID From Dv_Usergroups WHERE ParentGID = 4 AND NOT MinArticle = -1 ORDER BY MinArticle")
		Dvbbs.Execute("UPDATE [Dv_User] SET Userclass = '" & UserGrade(0) & "', Titlepic = '" & UserGrade(1) & "', UserGroupid = 4 WHERE Userid = " & Rs(0))
		Rs.Movenext
	Loop
	Set Rs = Nothing
	'ɾ���û���
	Dvbbs.Execute("DELETE FROM Dv_Usergroups WHERE UserGroupID = " & Request("Groupid"))
	'ɾ��������ȼ�
	Dvbbs.Execute("DELETE FROM Dv_Usergroups WHERE ParentGID = " & Request("Groupid"))
	Response.Write "ɾ���ɹ���"
	ReloadGroupTitle()
End Sub

sub editgroup()
if not isnumeric(request("groupid")) then
response.write "����Ĳ�����"
exit sub
end if
if request("groupaction")="yes" then
	dim GroupSetting
	if request.form("title")="" then
	response.write "�������û������ƣ�"
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
	response.write "�޸ĳɹ���"
	ReloadGroup(request("groupid"))
else
Dim reGroupSetting
set rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "δ�ҵ����û��飡"
exit sub
end if
reGroupSetting=split(rs("GroupSetting"),",")
%>
<FORM METHOD=POST ACTION="?action=editgroup">
<input type=hidden name="groupid" value="<%=request("groupid")%>">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><td colspan=3 height=25 class="Forumrow"><B>˵��</B>��<BR>�����������������ø����û�������̳�е�Ĭ��Ȩ�ޣ���̳Ĭ���û��鲻��ɾ���ͱ༭�û�����<BR>�ڿ���ɾ���ͱ༭����ӵ��û���</td></tr>
<tr> 
<th height="23" colspan="3">�༭�û��飺<%=rs("title")%></th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�û�������</td>
<td height="23" width="40%" class=Forumrow colspan=2><input size=35 name="title" type=text value="<%=rs("title")%>"  <%if Cint(request("GroupID"))<9 then%>disabled<%end if%>></td>
</tr>
<%if Cint(request("GroupID"))<9 then%>
<input name="title" type=hidden value="<%=rs("title")%>">
<%end if%>
<tr>
<td height="23" width="60%" class=Forumrow>�Ƿ�����ҳ������ͼ������ʾ</td>
<td height="23" width="40%" class=Forumrow colspan=2>��<input name="isdisp" type=radio value="1" <%if rs("isdisp")=1 then%>checked<%end if%>>&nbsp;��<input name="isdisp" type=radio value="0" <%if rs("isdisp")=0 then%>checked<%end if%>></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��ҳ����ͼ��ͼƬ��ַ</td>
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
	response.write "�������û������ƣ�"
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
		Errmsg=ErrMsg + "<BR><li>���û��������Ѿ����ڡ�"
		dvbbs_error()
		exit sub
	end if
	rs.close
	set rs=nothing
	response.write "����û���ɹ����û�������ͬʱΪ�ȼ����ƣ��µȼ�������Ĭ�ϵ�ͼƬ���ã������Ե��ȼ������н����޸ģ�"
	Dim MaxID
	Set Rs=Dvbbs.Execute("Select Top 1 * from Dv_UserGroups Order by UserGroupID Desc")
	MaxID=rs(0)
	Set Rs=Nothing
	ReloadGroup(MaxID)
else
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<FORM METHOD=POST ACTION="?action=addgroup">
<tr><td colspan=3 height=25 class="Forumrow"><B>˵��</B>��<BR>�ٿ��Խ�������û��������������Ȩ�ޣ����Խ��������û�ת�Ƶ����飬�뵽�û������н�����ز�������ĳ�����û����й������������û��б�<BR>�ڿ���ɾ���ͱ༭����ӵ��û���<BR>��<B>����û���󣬸��û�����ͬʱΪ�ȼ�����</B></td></tr>
<tr> 
<th height="23" colspan="3" >����µ��û���</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�û�������</td>
<td height="23" width="40%" class=Forumrow colspan=2><input size=35 name="title" type=text></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�Ƿ�����ҳ������ͼ������ʾ</td>
<td height="23" width="40%" class=Forumrow colspan=2>��<input name="isdisp" type=radio value="1">&nbsp;��<input name="isdisp" type=radio value="0" checked></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��ҳ����ͼ��ͼƬ��ַ</td>
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