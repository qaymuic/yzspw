<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->

<%
Head()
Server.ScriptTimeout=9999999
dim iboardid(1000),idepth(1000),iboardname(1000)
dim k
dim n
dim admin_flag
admin_flag="22"
if not Dvbbs.master or instr(","&session("flag")&",",",22,")=0 then
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	dvbbs_error()
else
	dim body
	call main()
Footer()
end if
Erase iboardid
Erase idepth
Erase iboardname

sub main()
i=0
set rs=Dvbbs.Execute("select boardid,depth,boardtype from dv_board order by rootid,orders")
if rs.eof and rs.bof then
iboardid(0)=0
idepth(0)=0
iboardname(0)="û����̳"
else
do while not rs.eof
iboardid(i)=rs(0)
idepth(i)=rs(1)
iboardname(i)=rs(2)
i=i+1
rs.movenext
loop
end if
set rs=nothing
select case request("action")
case "alldel"
	call alldel()
case "userdel"
	call del()
case "alldelTopic"
	call alldelTopic()
case "delUser"
	call deluser()
case "moveinfo"
	call moveinfo()
case "MoveUserTopic"
	call moveusertopic()
case "MoveDateTopic"
	call movedatetopic()
case else
%>
<table cellpadding=3 cellspacing=1 border=0 width=95% align=center class="tableBorder">
	<tr>
    <td width="100%" valign=top class=forumrow>
<B>ע��</B>�����������������ɾ����̳���ӣ�<font color=red>�������в������ɻָ���</font>�����ȷ��������������ϸ������������Ϣ��
</td>
</tr>
</table><BR>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form action="admin_alldel.asp?action=alldel" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>ɾ��ָ������������</b>(�����ܲ��۳��û��������ͻ���)</th></tr>
            <tr>
            <td valign=middle width=40% class=forumrow>ɾ��������ǰ������(��д����)</td><td class=forumrow><input name="TimeLimited" value=100 size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>��̳����</td><td class=forumrow>
<select name="delboardid" size=1>
<%
for k=0 to i-1
	if iboardid(k)=0 then
		response.write "<option value=0>û����̳</option>"
	elseif k=0 then
		response.write "<option value=all>ȫ����̳</option>"
	end if
	response.write "<option value="&iboardid(k)&">"
	if idepth(k)>0 then
	for n=1 to idepth(k)
	response.write "��"
	next
	end if
	response.write iboardname(k)&"</option>"
next
%>
</select>
			</td></tr>
</form>
<form action="admin_alldel.asp?action=alldelTopic" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>ɾ��ָ��������û�лظ�������(�����ܲ��۳��û��������ͻ���)</th></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>ɾ��������ǰ������(��д����)</td><td class=forumrow><input name="TimeLimited" value=100 size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>��̳����</td><td class=forumrow>
<select name="delboardid" size=1>
<%
for k=0 to i-1
	if iboardid(k)=0 then
		response.write "<option value=0>û����̳</option>"
	elseif k=0 then
		response.write "<option value=all>ȫ����̳</option>"
	end if
	response.write "<option value="&iboardid(k)&">"
	if idepth(k)>0 then
	for n=1 to idepth(k)
	response.write "��"
	next
	end if
	response.write iboardname(k)&"</option>"
next
%>
</select>
			</td></tr>
</form>
<form action="admin_alldel.asp?action=userdel" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>ɾ��ĳ�û�����������</td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>�������û���</td><td class=forumrow><input type=text name="username" size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>��̳����</td><td class=forumrow>
<select name="delboardid" size=1>
<%
for k=0 to i-1
	if iboardid(k)=0 then
		response.write "<option value=0>û����̳</option>"
	elseif k=0 then
		response.write "<option value=all>ȫ����̳</option>"
	end if
	response.write "<option value="&iboardid(k)&">"
	if idepth(k)>0 then
	for n=1 to idepth(k)
	response.write "��"
	next
	end if
	response.write iboardname(k)&"</option>"
next
%>
</select>
			</td></tr>
</form>

<form action="admin_alldel.asp?action=delUser" method="post">
            <tr>
            <td class=forumrow valign=middle>ɾ��ָ��������û�е�¼���û�</td>
            <td class=forumrow valign=middle>
<select name=TimeLimited size=1> 
<option value=1>ɾ��һ��ǰ��
<option value=2>ɾ������ǰ��
<option value=7>ɾ��һ����ǰ��
<option value=15>ɾ�������ǰ��
<option value=30>ɾ��һ����ǰ��
<option value=60>ɾ��������ǰ��
<option value=180>ɾ������ǰ��
</select>
</select><input type=submit name="submit" value="�� ��"></td></tr></form>

</table>
<%end select%>
<%if founderr then Call dvbbs_error()%>
<%
end sub

Sub Moveinfo()
%>
<table cellpadding=3 cellspacing=1 border=0 width=95% align=center>
	<tr>
    <td width="100%" valign=top>
<B>ע��</B>������ֻ���ƶ����ӣ������ǿ�������ɾ����
            <br>���������ɾ��ԭ��̳���ӣ����ƶ�����ָ������̳�С������ȷ��������������ϸ������������Ϣ��<BR>�����Խ�һ����̳������̳�������ƶ����ϼ���̳��Ҳ���Խ��ϼ���̳�������ƶ����¼���̳������Ϊ�������̳������̳���úܿ��ܲ��ܷ������ӣ�ֻ�������
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form action="admin_alldel.asp?action=MoveDateTopic" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>�������ƶ�</th></tr>
            <tr>
            <td valign=middle width=40% class=forumrow>�ƶ�������ǰ������(��д����)</td><td class=forumrow><input name="TimeLimited" value=0 size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>ԭ��̳</td><td class=forumrow>
<select name="outboardid" size=1>
<%
for k=0 to i-1
	if iboardid(k)=0 then
		response.write "<option value=0>û����̳</option>"
	end if
	response.write "<option value="&iboardid(k)&">"
	if idepth(k)>0 then
	for n=1 to idepth(k)
	response.write "��"
	next
	end if
	response.write iboardname(k)&"</option>"
next
%>
</select>
			</td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>Ŀ����̳</td><td class=forumrow>
<select name="inboardid" size=1>
<%
for k=0 to i-1
	if iboardid(k)=0 then
		response.write "<option value=0>û����̳</option>"
	end if
	response.write "<option value="&iboardid(k)&">"
	if idepth(k)>0 then
	for n=1 to idepth(k)
	response.write "��"
	next
	end if
	response.write iboardname(k)&"</option>"
next
%>
</select>
			</td></tr>
</form>
<form action="admin_alldel.asp?action=MoveUserTopic" method="post">
            <tr>
            <th valign=middle colspan=2 height=23 align=left>���û��ƶ�</th></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>����д�û���</td><td class=forumrow><input name="username" size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>ԭ��̳</td><td class=forumrow>
<select name="outboardid" size=1>
<%
for k=0 to i-1
	if iboardid(k)=0 then
		response.write "<option value=0>û����̳</option>"
	end if
	response.write "<option value="&iboardid(k)&">"
	if idepth(k)>0 then
	for n=1 to idepth(k)
	response.write "��"
	next
	end if
	response.write iboardname(k)&"</option>"
next
%>
</select>
			</td></tr>
            <tr>
            <td valign=middle width=40%  class=forumrow>Ŀ����̳</td><td class=forumrow>
<select name="inboardid" size=1>
<%
for k=0 to i-1
	if iboardid(k)=0 then
		response.write "<option value=0>û����̳</option>"
	end if
	response.write "<option value="&iboardid(k)&">"
	if idepth(k)>0 then
	for n=1 to idepth(k)
	response.write "��"
	next
	end if
	response.write iboardname(k)&"</option>"
next
%>
</select>
			</td></tr>
</form>
</table>
<%
	end sub
	'ɾ��ĳ�û�����������
	sub del()
		dim titlenum,delboardid,PostUserID,delboardida
		if request("delboardid")="0" then
			founderr=true
			Errmsg=ErrMsg + "<BR><li>�Ƿ��İ��������"
			exit sub
		elseif request("delboardid")="all" then
			delboardid=""
			delboardida=""
		else
			delboardid=" boardid="&request("delboardid")&" and "
			delboardida=" F_boardid="&request("delboardid")&" and "
		end if
		if request("username")="" then
			founderr=true
			Errmsg=ErrMsg + "<BR><li>�����뱻����ɾ���û�����"
			exit sub
		end if
		Set Rs=Dvbbs.Execute("Select UserID,UserGroupID From Dv_User Where UserName='"&replace(request("username"),"'","")&"'")
		If Rs.Eof And Rs.Bof Then
			founderr=true
			Errmsg=ErrMsg + "<BR><li>Ŀ���û������ڣ����������롣"
			exit sub
		End If
		If Rs(1)=1 Or Rs(1)=2 Or Rs(1)=3 Then
			founderr=true
			Errmsg=ErrMsg + "<BR><li>�Թ���Ա���������������������Ӳ��ܽ�������ɾ��������"
			exit sub
		End If
		PostUserID=Rs(0)
		Rs.close:Set Rs=Nothing
		titlenum=0
		for i=0 to ubound(allposttable)
		set rs=Dvbbs.Execute("Select Count(*) from "&allposttable(i)&" where "&delboardid&" PostUserID="&PostUserID) 
   		titlenum=titlenum+rs(0)

		sql="Delete From "&allposttable(i)&" where "&delboardid&" PostUserID="&PostUserID
		Dvbbs.Execute(sql)
		next
		Rs.close:Set Rs=Nothing
		'����
		Dvbbs.Execute("delete from dv_besttopic where "&delboardid&" PostUserID="&PostUserID)
		'�ϴ�
		Dvbbs.Execute("delete from Dv_UpFile where "&delboardida&" F_UserID="&PostUserID)
		'���û���������⡢��������һ��ɾ��
		set rs=Dvbbs.Execute("select topicid,posttable from dv_topic where "&delboardid&" PostUserID="&PostUserID)
		do while not rs.eof
			Dvbbs.Execute("Delete From "&rs(1)&" where rootid="&rs(0))
		rs.movenext
		loop
		Rs.close:Set Rs=Nothing
		Dvbbs.Execute("Delete From dv_topic where "&delboardid&" PostUserID="&PostUserID)
		if isnull(titlenum) then titlenum=0
		sql="update [dv_user] set userpost=userpost-"&titlenum&",userWealth=userWealth-"&titlenum*Dvbbs.Forum_user(3)&",userEP=userEP-"&titlenum*Dvbbs.Forum_user(8)&",userCP=userCP-"&titlenum*Dvbbs.Forum_user(13)&" where UserID="&PostUserID
		Dvbbs.Execute(sql)
		response.write "ɾ���ɹ���<BR>��������������̳�����и���һ����̳���ݣ�����<a href=admin_alldel.asp>����</a>"
	end sub

'ɾ��ָ������������
Sub Alldel()
	Dim TimeLimited,Delboardid,DelSql
	If Request("delboardid")="0" Then
		'founderr=true
		Errmsg=ErrMsg + "<BR><li>�Ƿ��İ��������"
		Exit Sub
	Elseif Request("delboardid")="all" Then
		Delboardid=""
	Else
		'Delboardid="And boardid="&Clng(Request("delboardid"))
		Delboardid=" boardid="&Clng(Request("delboardid"))&" and "
	End If
	TimeLimited=Request.Form("TimeLimited")
	If Not Isnumeric(TimeLimited) Then
		'founderr=true
		Errmsg=ErrMsg + "<BR><li>�Ƿ��Ĳ�����"
		Exit Sub
	Else
		For i=0 to Ubound(allposttable)
			If IsSqlDataBase=1 Then
				Dvbbs.Execute("DELETE FROM "&Allposttable(i)&" WHERE "&Delboardid&" Datediff(d,DateAndTime,"&SqlNowString&")>"&TimeLimited)
			Else
				Dvbbs.Execute("DELETE FROM "&Allposttable(i)&" WHERE "&Delboardid&" Datediff('d',DateAndTime,"&SqlNowString&")>"&TimeLimited)
			End if
			Response.Write Allposttable(i)&"������ɾ����ɣ�<BR>"
			Response.Flush
		Next
		If IsSqlDataBase=1 Then
			Dvbbs.Execute("DELETE FROM Dv_topic WHERE "&Delboardid&" Datediff(d,DateAndTime,"&SqlNowString&")>"&TimeLimited)
			Dvbbs.Execute("delete from dv_besttopic where "&Delboardid&" datediff(d,DateAndTime,"&SqlNowString&")>"&TimeLimited)
		Else
			Dvbbs.Execute("DELETE FROM Dv_topic WHERE "&Delboardid&" Datediff('d',DateAndTime,"&SqlNowString&")>"&TimeLimited)
			Dvbbs.Execute("DELETE FROM Dv_besttopic WHERE "&Delboardid&" Datediff('d',DateAndTime,"&SqlNowString&") > "&TimeLimited)
		End If
			Response.Write "Dv_topic����ɾ����ɣ�<BR>"
			Response.Flush
	End if
	Response.write "ɾ���ɹ���<BR>��������������̳�����и���һ����̳���ݣ�����<a href=admin_alldel.asp>����</a>"
	Response.Flush
End sub

	sub alldelTopic()
	Dim TimeLimited,delboardid
	if request("delboardid")="0" then
		'founderr=true
		Errmsg=ErrMsg + "<BR><li>�Ƿ��İ��������"
		exit sub
	elseif request("delboardid")="all" then
		delboardid=""
	else
		delboardid=" boardid="&request("delboardid")&" and "
	end if
	TimeLimited=request.form("TimeLimited")
	if not isnumeric(TimeLimited) then
		'founderr=true
		Errmsg=ErrMsg + "<BR><li>�Ƿ��Ĳ�����"
		exit sub
	else
	if IsSqlDataBase=1 then
		set rs=Dvbbs.Execute("select Topicid,PostTable from dv_topic where "&delboardid&"   datediff(d,DateAndTime,"&SqlNowString&")>"&TimeLimited&" and Child=0")
	else
		set rs=Dvbbs.Execute("select Topicid,PostTable from dv_topic where "&delboardid&"   datediff('d',DateAndTime,"&SqlNowString&")>"&TimeLimited&" and Child=0")
	end if
	do while not rs.eof
		Dvbbs.Execute("Delete From "&rs(1)&" where rootid="&rs(0))
		Dvbbs.Execute("delete from dv_besttopic where rootid="&rs(0))
	rs.movenext
	loop
	if IsSqlDataBase=1 then
		Dvbbs.Execute("Delete From dv_topic where "&delboardid&"   datediff(d,DateAndTime,"&SqlNowString&")>"&TimeLimited&" and Child=0")
	else
		Dvbbs.Execute("Delete From dv_topic where "&delboardid&"   datediff('d',DateAndTime,"&SqlNowString&")>"&TimeLimited&" and Child=0")
	end if
	set rs=nothing
	end if
	response.write "ɾ���ɹ���<BR>��������������̳�����и���һ����̳���ݣ�����<a href=admin_alldel.asp>����</a>"
	end sub

	sub delUser()
	Dim TimeLimited
	TimeLimited=request.form("TimeLimited")
	if TimeLimited="all" then
	response.Write "���˰ɣ��뿪��ɣ���������������Ա��ɾ���ģ�"
	else
	if IsSqlDataBase=1 then
	set rs=Dvbbs.Execute("select userid,username,usergroupid from [dv_user] where datediff(d,LastLogin,"&SqlNowString&")>"&TimeLimited&"")
	else
	set rs=Dvbbs.Execute("select userid,username,usergroupid from [dv_user] where datediff('d',LastLogin,"&SqlNowString&")>"&TimeLimited&"")
	end if
	'shinzeal����ɾ���û���ͬʱ�Զ�ɾ�������ӣ��������������Ĺ���
	do while not rs.eof
		If rs(2)>3 then
		for i=0 to ubound(allposttable)
		sql="Delete From "&allposttable(i)&" where postuserid="&rs(0)
		Dvbbs.Execute(sql)
		next
		Dvbbs.Execute("delete from dv_besttopic where postuserid="&rs(0))
		Dvbbs.Execute("Delete From Dv_UpFile Where F_UserID="&rs(0))
		Dvbbs.Execute("Delete From Dv_Message Where Sender='"&Replace(Rs(1),"'","''")&"'")
		Dvbbs.Execute("Delete From Dv_Friend Where F_UserID="&rs(0))
		Dvbbs.Execute("Delete From Dv_BookMark Where UserName='"&Replace(Rs(1),"'","''")&"'")
		dim rrs
		set rrs=Dvbbs.Execute("select topicid,posttable from dv_topic where postuserid="&rs(0))
		do while not rrs.eof
		Dvbbs.Execute("Delete From "&rrs(1)&" where rootid="&rrs(0))
		rrs.movenext
		loop
		set rrs=nothing
		Dvbbs.Execute("Delete From dv_topic where postuserid="&rs(0))
		end if
	rs.movenext
	loop
	set rs=nothing
	if IsSqlDataBase=1 then
	Dvbbs.Execute("delete from [dv_user] where datediff(d,LastLogin,"&SqlNowString&")>"&TimeLimited&"")
	else
	Dvbbs.Execute("delete from [dv_user] where datediff('d',LastLogin,"&SqlNowString&")>"&TimeLimited&"")
	end if
	end if
	response.write "ɾ���ɹ���<BR>��������������̳�����и���һ����̳���ݣ�����<a href=admin_alldel.asp>����</a>"
	end sub

Sub MoveUserTopic()
	Dim PostUserID
	If Not Isnumeric(Request("Inboardid")) Then
		Response.Write "����İ��������"
		Exit Sub
	End If
	If Not Isnumeric(Request("Outboardid")) Then
		Response.Write "����İ��������"
		Exit Sub
	End If
	If Request("Username") = "" Then
		Response.Write "����д�û�����"
		Exit Sub
	End If
	If Cint(Request("Outboardid")) = Cint(Request("Inboardid")) Then
		Response.Write "��������ͬ��������ƶ�������"
		Exit Sub
	End If
	Set Rs = Dvbbs.Execute("Select UserID From Dv_User Where UserName = '" & Replace(Request("Username"), "'", "''") & "'")
	If Rs.Eof And Rs.Bof Then
		Response.Write "Ŀ���û����������ڣ����������룡"
		Exit Sub
	End If
	PostUserID = Rs(0)
	For i = 0 To Ubound(Allposttable)
		Dvbbs.Execute("UPDATE " & Allposttable(i) & " SET Boardid = " & Request("Inboardid") & " WHERE Boardid = " & Request("Outboardid") & " AND PostUserID = " & PostUserID)
	Next
	Rs.Close:Set Rs = Nothing
	REM �޸������ƶ���ʽ 2004-4-25 Dvbbs.YangZheng
	SET Rs = Dvbbs.Execute("SELECT Topicid, Posttable, Istop FROM Dv_Topic WHERE Boardid = " & Request("Outboardid") & " AND PostUserID = " & PostUserID)
	Rem Topicid:0, Posttable:1, Istop:2
	If Not(Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		Dim Yrs, TopstrinfoN, TopstrinfoO
		For i = 0 To Ubound(Sql,2)
			Dvbbs.Execute("UPDATE " & Sql(1,i) & " SET Boardid = " & Request("Inboardid") & " WHERE Rootid = " & Sql(0,i))
			Dvbbs.Execute("UPDATE Dv_Topic SET Boardid = " & Request("Inboardid") & " WHERE Boardid = " & Request("Outboardid") & " AND Topicid = " & Sql(0,i))
			If Sql(2,0) > 0 Then
				'��ȡ�¾ɰ���Ĺ̶���Ϣ
				Set Yrs = Dvbbs.Execute("SELECT BoardTopStr From Dv_Board Where Boardid = " & Request("Outboardid"))
				TopstrinfoO = Yrs(0)
				Set Yrs = Dvbbs.Execute("SELECT BoardTopStr From Dv_Board Where Boardid = " & Request("Inboardid"))
				TopstrinfoN = Yrs(0)
				Yrs.Close:Set Yrs = Nothing
				'ɾ��ԭ�̶�����ID
				TopstrinfoO = Replace(TopstrinfoO, Cstr(Sql(0,i))&",", "")
				TopstrinfoO = Replace(TopstrinfoO, ","&Cstr(Sql(0,i)), "")
				TopstrinfoO = Replace(TopstrinfoO, Cstr(Sql(0,i)), "")
				If TopstrinfoN = "" Or Isnull(TopstrinfoN) Then
					TopstrinfoN = Cstr(Sql(0,i))
				ElseIf TopstrinfoN = Cstr(Sql(0,i)) Then
					TopstrinfoN = TopstrinfoN
				ElseIf Instr(TopstrinfoN, ","&Cstr(Sql(0,i))) > 0 Then
					TopstrinfoN = TopstrinfoN
				Else
					TopstrinfoN = TopstrinfoN & "," & Cstr(Sql(0,i))
				End If
				'���µ�ǰ����̶���Ϣ������
				Sql = "UPDATE Dv_Board SET BoardTopStr = '" & TopstrinfoO & "' WHERE BoardID = " & Request("Outboardid")
				Dvbbs.Execute(Sql)
				Dvbbs.ReloadBoardInfo(Request("Outboardid"))
				'�����°���̶���Ϣ������
				Sql = "UPDATE Dv_Board SET BoardTopStr = '" & TopstrinfoN & "' WHERE Boardid = " & Request("Inboardid")
				Dvbbs.Execute(Sql)
				Dvbbs.ReloadBoardInfo(Request("Inboardid"))
			End If
		Next
	End If
	Dvbbs.Execute("UPDATE Dv_Besttopic SET Boardid = " & Request("Inboardid") & " WHERE Boardid = " & Request("Outboardid") & " AND PostUserID = " & PostUserID)
	'shinzeal�����ƶ��ϴ��ļ�����
	Dvbbs.Execute("UPDATE Dv_Upfile SET F_Boardid = " & Request("Inboardid") & " WHERE F_Boardid = " & Request("Outboardid") & " AND F_UserID = " & PostUserID)
	Response.Write "�ƶ��ɹ���<br>�ڡ��ؼ���̳���ݺ��޸����С�������̳���ݡ���"
End Sub

Sub MoveDateTopic()
	If Not Isnumeric(Request("TimeLimited")) Then
		Response.Write "��������ڲ�����"
		Exit Sub
	end if
	If Not Isnumeric(Request("Inboardid")) Then
		Response.Write "����İ��������"
		Exit Sub
	End If
	If Not Isnumeric(Request("Outboardid")) Then
		Response.Write "����İ��������"
		Exit Sub
	End If
	If Cint(Request("Outboardid")) = Cint(Request("Inboardid")) Then
		Response.Write "��������ͬ��������ƶ�������"
		Exit Sub
	End If
	Rem �޸��ƶ���ʽ 2004-4-25 Dvbbs.YangZheng
	If IsSqlDataBase = 1 Then
		Sql = "SELECT PostTable,Isbest,IsTop,TopicID FROM Dv_Topic WHERE Boardid = " & Request("Outboardid") & " AND DATEDIFF(d, DateAndTime, " & SqlNowString & ") > " & Request.Form("TimeLimited")
	Else
		Sql = "SELECT PostTable,Isbest,IsTop,TopicID FROM Dv_Topic WHERE Boardid = " & Request("Outboardid") & " AND DATEDIFF('d', DateAndTime, " & SqlNowString & ") > " & Request.Form("TimeLimited")
	End If
	Rem PostTable:0, Isbest:1, IsTop:2, TopicID:3
	Set Rs = Dvbbs.Execute(Sql)
	If Not(Rs.Eof And Rs.Bof) Then
		Sql = Rs.Getrows(-1)
		Rs.Close:Set Rs = Nothing
		For i = 0 To Ubound(Sql,2)
			Dvbbs.Execute("UPDATE " & Sql(0,i) & " SET BoardID = " & Request("Inboardid") & " WHERE BoardID = " & Request("Outboardid") & " AND RootID = " & Sql(3,i))
			Dvbbs.Execute("UPDATE Dv_Topic SET BoardID = " & Request("Inboardid") & " WHERE BoardID = " & Request("Outboardid") & " AND TopicID = " & Sql(3,i))
			If Sql(1,i) = 1 Then
				Dvbbs.Execute("UPDATE Dv_Besttopic Set BoardID = " & Request("Inboardid") & " WHERE BoardID = " & Request("Outboardid") & " AND RootID = " & Sql(3,i))
			End If
			If Sql(2,i) > 0 Then
				Dim Yrs, TopstrinfoN, TopstrinfoO
				'��ȡ�¾ɰ���Ĺ̶���Ϣ
				Set Yrs = Dvbbs.Execute("SELECT BoardTopStr From Dv_Board Where Boardid = " & Request("Outboardid"))
				TopstrinfoO = Yrs(0)
				Set Yrs = Dvbbs.Execute("SELECT BoardTopStr From Dv_Board Where Boardid = " & Request("Inboardid"))
				TopstrinfoN = Yrs(0)
				Yrs.Close:Set Yrs = Nothing
				'ɾ��ԭ�̶�����ID
				TopstrinfoO = Replace(TopstrinfoO, Cstr(Sql(3,i))&",", "")
				TopstrinfoO = Replace(TopstrinfoO, ","&Cstr(Sql(3,i)), "")
				TopstrinfoO = Replace(TopstrinfoO, Cstr(Sql(3,i)), "")
				If TopstrinfoN = "" Or Isnull(TopstrinfoN) Then
					TopstrinfoN = Cstr(Sql(3,i))
				ElseIf TopstrinfoN = Cstr(Sql(3,i)) Then
					TopstrinfoN = TopstrinfoN
				ElseIf Instr(TopstrinfoN, ","&Cstr(Sql(3,i))) > 0 Then
					TopstrinfoN = TopstrinfoN
				Else
					TopstrinfoN = TopstrinfoN & "," & Cstr(Sql(3,i))
				End If
				'����ԭ����̶���Ϣ������
				Sql = "UPDATE Dv_Board SET BoardTopStr = '" & TopstrinfoO & "' WHERE BoardID = " & Request("Outboardid")
				Dvbbs.Execute(Sql)
				Dvbbs.ReloadBoardInfo(Request("Outboardid"))
				'�����°���̶���Ϣ������
				Sql = "UPDATE Dv_Board SET BoardTopStr = '" & TopstrinfoN & "' WHERE Boardid = " & Request("Inboardid")
				Dvbbs.Execute(Sql)
				Dvbbs.ReloadBoardInfo(Request("Inboardid"))
			End If
		Next
	End If
	Response.Write "�ƶ��ɹ���<br>�ڡ��ؼ���̳���ݺ��޸����С�������̳���ݡ���"
End Sub
%>