<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<!-- #include file="inc/GroupPermission.asp" -->

<%
	Head()
	dim admin_flag
	admin_flag=",17,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
<B>�����û��ȼ����õ�˵��������ϸ�Ķ���������</B>��<BR>
����û�������޶�Ӧ�ȼ����ƣ���ע���û��Զ�������������<BR>
����û���ĵȼ����ƿ��Ժ��û�������һ��<BR>
�����ȼ����趨���Զ���Ȩ�ޣ�Ȩ�����ͺ���Ȩ��һ�������ĳ���ȼ��趨���Զ���Ȩ�ޣ������õȼ�����̳���а���������Զ����Ȩ�ޣ����Ҹõȼ��������ڰ���Ȩ�޶������˵��У������Ϳ��趨ĳ���ȼ���ĳ�������е�Ȩ�ޣ���������صȼ��û�����ID����ɵȼ�ID
</td>
<td width="50%" valign=top>
<B>�ڵȼ����趨�û�����ʲô�ã�</B><BR>
һ����˵��ֻ��ע���û�ӵ�еȼ��������ڵȼ���������һ�㶼�趨��Ӧע���û��飬������óɱ���飬��ô���û�������������ȼ���ͬʱҲ���Զ����������õ���<BR>
�������������һ���û��飬���Ҹ���������û���ĳһЩȨ�ޣ���ô��������ôﵽһ���ȼ������ӣ����û��Զ����µ�����û�����ʹ������û����Ȩ�ޡ�<BR>�������ĳ���ȼ����û��������������������ȼ�����ô�Ͱ����ٷ�������Ϊ<B>-1</B>��һ��Ϊ�����û�����Ҫ���������ã�����ĳ���������ٷ���Ϊ<B>-1</B>�󣬸ü�����û������ܸ����������Ӷ�����������û�Ҳ�����Զ��������ü���ֻ�����û������з��ܸ����伶��
</td>
</tr>
</table>
</td>
</tr>
<tr> 
<th height="23" colspan="5" >�û��ȼ��趨</th>
</tr>
<tr> 
<td width="25%" class=forumrowHighlight><B>�ȼ�����</B></td>
<td width="15%" class=forumrowHighlight><B>���ٷ���</B></td>
<td width="20%" class=forumrowHighlight><B>ͼƬ</B></td>
<td width="25%" class=forumrowHighlight><B>������</B></td>
<td width="15%" class=forumrowHighlight><B>����</B></td>
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
	<td width="15%" class=Forumrow><%If Rs("UserGroupID")>8 Then%><a href="?action=del&id=<%=rs("UserGroupID")%>">ɾ��</a> | <%End If%>
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
	%>Ȩ��</a>
	<%If Rs("UserGroupID")>8 And Rs("IsSetting")=1 Then Response.Write " <font color=red>��</font>"%>
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
<input type="submit" name="Submit" value="�� ��">
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
			'������Զ���Ȩ�ޣ���UserGroupIDΪ�ȼ����ڵ�ID����֮��Ϊ��ID
			if rs("issetting")=1 then groupid=rs("usergroupid")
			Dvbbs.Execute("update [dv_user] set userclass='"&usertitle&"',titlepic='"&titlepic&"',usergroupid="&groupid&" where userclass='"&rs("usertitle")&"'")
		end if
		if rs("parentgid")=0 then groupid=0
		Dvbbs.Execute("update dv_usergroups set usertitle='"&usertitle&"',minarticle="&minarticle&",grouppic='"&titlepic&"',parentgid="&groupid&" where usergroupid="&usertitleID)
		end if
	next
	response.write "���óɹ����뷵�ء�"
	set rs=nothing
End Sub

sub add()
%>
<form method="POST" action=admin_grade.asp?action=savenew>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th colspan="2">����µ��û��ȼ�</th>
</tr>
<tr>
<td width="40%" class=forumrow><B>�����û���</B></td>
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
<td width="40%" class=forumrow><B>�ȼ�����</B></td>
<td width="60%" class=forumrow><input size=30 name="usertitle" type=text></td>
</tr>
<tr>
<td width="40%" class=forumrow><B>���ٷ���</B><BR>����õȼ��������ƺŻ��߹�����ݣ����������д-1����ʾ��������������������</td>
<td width="60%" class=forumrow><input size=30 name="minarticle" type=text></td>
</tr>
<tr>
<td width="40%" class=forumrow><B>�ȼ�ͼƬ</B></td>
<td width="60%" class=forumrow><input size=30 name="titlepic" type=text>&nbsp;�⽫����������������ʾ��ߵ��û�������</td>
</tr>
<tr> 
<td width="100%" colspan=2 class=forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</table>
</form>
<%
end sub
sub savenew()
if request.form("minarticle")="" then
	Errmsg=ErrMsg + "<BR><li>�������µĵȼ���Ҫ��������"
	dvbbs_error()
	exit sub
elseif not isnumeric(request.form("minarticle")) then
	Errmsg=ErrMsg + "<BR><li>�µĵȼ�������ֻ�������֡�"
	dvbbs_error()
	exit sub
end if
if request.form("titlepic")="" then
	Errmsg=ErrMsg + "<BR><li>�������µĵȼ�ͼƬ��"
	dvbbs_error()
	exit sub
end if
if request.form("usertitle")="" then
	Errmsg=ErrMsg + "<BR><li>�������µĵȼ����ơ�"
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
	Errmsg=ErrMsg + "<BR><li>�õȼ������Ѿ����ڡ�"
	dvbbs_error()
	exit sub
end if
rs.close
set rs=nothing
response.write "��ӳɹ����������������û������н��и��²�����"
end sub

Sub Del()
	Server.ScriptTimeout = 99999999
	Dim Minarticle, Minuserclass
	If Isnumeric(Request("Id")) Then
		If CLng(Request("id")) < 9 Then
			Errmsg = ErrMsg + "<BR><li>ϵͳĬ�ϵȼ�����ɾ����"
			Dvbbs_Error()
			Exit Sub
		End If
		Set Rs = Dvbbs.Execute("SELECT * FROM Dv_UserGroups WHERE UserGroupId = " & Request("id"))
		Minarticle = Rs("Minarticle")
		Minuserclass = Rs("Usertitle")
		Rem ����ɾ���ȼ���ȼ����õĴ��� 2004-5-1 Dvbbs.YangZheng
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
		Response.Write "ɾ���ɹ���"
		Set Rs = Nothing
	End If
End Sub

sub per()
if not isnumeric(request("groupid")) then
response.write "����Ĳ�����"
exit sub
end if
if request("groupaction")="yes" then
	dim GroupSetting,groupid,newgroupsetting
	GroupSetting=GetGroupPermission
	if request("isdefault")=1 then
		set rs=dvbbs.execute("select * from dv_usergroups where usergroupid="&request("groupid"))
		If Rs("ParentGID")=0 Then
			Dv_suc("��û��ѡ���Զ���ȼ�ѡ������޸Ľ���Ч")
			Exit Sub
		End If
		if rs("issetting")=1 then
		groupid=rs("parentgid")
		set rs=nothing
		set rs=dvbbs.execute("select * from dv_usergroups where usergroupid="&groupid)
		newgroupsetting=rs("groupsetting")
		Set Rs=Nothing
		dvbbs.execute("update dv_usergroups set issetting=0,groupsetting='"&newgroupsetting&"' where usergroupid="&request("regroupid"))
		'ȡ���Զ������ã������û����ݣ�����Ϊ�û���ID
		dvbbs.execute("update [dv_user] set usergroupid="&groupid&" where userclass='"&request("usertitle")&"'")
		end if
		
	else
		dvbbs.execute("update dv_usergroups set issetting=1,groupsetting='"&GroupSetting&"' where usergroupid="&request("regroupid"))
		'�����û�����
		dvbbs.execute("update [dv_user] set usergroupid="&request("regroupid")&" where userclass='"&request("usertitle")&"'")
	End If

	ReloadGroup(request("regroupid"))
	Dv_suc("�޸ĵȼ��Զ���Ȩ�޳ɹ�")
else
Dim reGroupSetting,founduserper,usergrade
If IsNumerIc(request("regroupid")) and request("regroupid")<>"" Then
	Set Rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&request("regroupid"))
	usergrade=rs("usertitle")
End If
founduserper=false
set rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "δ�ҵ��û��ȼ�"
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
<tr><td colspan=3 height=25 class="Forumrow"><B>˵��</B>��<BR>���������������ø����ȼ�����̳�е�Ĭ��Ȩ�ޣ�<font color=blue>Ĭ��Ϊʹ�øõȼ������û���Ȩ�ޣ����Ҫ�øõȼ����Զ���Ȩ�ޣ����޸�ʱ��ѡ���Զ�������ѡ��</font></td></tr>
<tr> 
<th height="23" colspan="3" >�༭��̳�û��ȼ�Ȩ��&nbsp;>> <%=rs("usertitle")%><%if usergrade<>"" then Response.Write "&nbsp;>> "&usergrade&""%></th>
</tr>
<tr> 
<td height="23" colspan="3" class=forumrow><input type=radio name="isdefault" value="1" <%if not founduserper then%>checked<%end if%>><B>ʹ���û���Ĭ��ֵ</B> (ע��: �⽫ɾ���κ�֮ǰ�������Զ�������)</td>
</tr>
<tr> 
<td height="23" colspan="3"  class=forumrow><input type=radio name="isdefault" value="0" <%if founduserper then%>checked<%end if%>><B>ʹ���Զ�������</B>&nbsp;(ѡ���Զ������ʹ����������Ч) </td>
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
