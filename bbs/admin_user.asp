<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/md5.asp"-->
<!-- #include file="inc/DvADChar.asp" -->
<!-- #include file="inc/GroupPermission.asp" -->
<%
	Head()
	dim admin_flag,sqlstr,myrootid
	FoundErr=False 
	admin_flag=",14,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	End if
	if request("action")="fix" Then
	Call Fixuser()
	End If
dim trs
dim userinfo
dim usertitle
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr>
<th align=left colspan=8 height=23>�û�����</th>
</tr>
<tr>
<td width=20% class=forumrowHighlight>ע������</td>
<td width=80% class=forumrowHighlight colspan=7><li>�ٵ�ɾ����ť��ɾ����ѡ�����û����˲����ǲ�����ģ�<li>�������������ƶ��û�����Ӧ���飻<li>�۵��û���������Ӧ�����ϲ�����<li>�ܵ��û�����½IP�ɽ�������IP������<li>�ݵ��û�Email�������û�����Email<li>�޵��޸����ӽ����޸����û��������������ݲ���������������������ɾID�û������޸���</td>
</tr>
<tr>
<td width=100% class=forumrowHighlight colspan=8>
��ݷ�ʽ��<a href="admin_user.asp">�û�������ҳ</a> | <a href="?action=userSearch&userSearch=1"><%If Request("userSearch")="1" Then%><font color=red><%End If%>�����û�<%If Request("userSearch")="1" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=2"><%If Request("userSearch")="2" Then%><font color=red><%End If%>����TOP100<%If Request("userSearch")="2" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=3"><%If Request("userSearch")="3" Then%><font color=red><%End If%>����END100<%If Request("userSearch")="3" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=4"><%If Request("userSearch")="4" Then%><font color=red><%End If%>24H�ڵ�¼<%If Request("userSearch")="4" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=5"><%If Request("userSearch")="5" Then%><font color=red><%End If%>24H��ע��<%If Request("userSearch")="5" Then%></font><%End If%></a><BR>
����������<a href="?action=userSearch&userSearch=6"><%If Request("userSearch")="6" Then%><font color=red><%End If%>�ȴ�������֤<%If Request("userSearch")="6" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=7"><%If Request("userSearch")="7" Then%><font color=red><%End If%>�ȴ���֤<%If Request("userSearch")="7" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=8"><%If Request("userSearch")="8" Then%><font color=red><%End If%>�����Ŷ�<%If Request("userSearch")="8" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=11"><%If Request("userSearch")="11" Then%><font color=red><%End If%>���Ρ��û�<%If Request("userSearch")="11" Then%></font><%End If%></a> | <a href="?action=userSearch&userSearch=12"><%If Request("userSearch")="12" Then%><font color=red><%End If%>���� �û�<%If Request("userSearch")="12" Then%></font><%End If%></a><%If Dvbbs.Forum_ChanSetting(0)="1" Then%> | <a href="?action=userSearch&userSearch=13"><%If Request("userSearch")="13" Then%><font color=red><%End If%>���� ��Ա<%If Request("userSearch")="13" Then%></font><%End If%></a><%End If%> | <a href="?action=uniteuser">�ϲ��û�</a> | <a href="?action=userSearch&userSearch=14"><%If Request("userSearch")="13" Then%><font color=red><%End If%>�Զ���Ȩ���û�<%If Request("userSearch")="13" Then%></font><%End If%></a>
</td>
</tr>
<%if request("action")="" or request("userSearch")="0" then%>
<form action="?action=userSearch" method=post>
<tr>
<th align=left colspan=7 height=23>�߼���ѯ</th>
</tr>
<tr>
<td width=20% class=forumrow>ע������</td>
<td width=80% class=forumrow colspan=5>�ڼ�¼�ܶ���������������Խ���ѯԽ�����뾡�����ٲ�ѯ�����������ʾ��¼��Ҳ����ѡ�����</td>
</tr>
<tr>
<td width=20% class=forumrow>�����ʾ��¼��</td>
<td width=80% class=forumrow colspan=5><input size=45 name="searchMax" type=text value=100></td>
</tr>
<tr>
<td width=20% class=forumrow>�û���</td>
<td width=80% class=forumrow colspan=5><input size=45 name="username" type=text>&nbsp;<input type=checkbox name="usernamechk" value="yes" checked>�û�������ƥ��</td>
</tr>
<tr>
<td width=20% class=forumrow>�û���</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="usergroups">
<option value=0>����</option>
<%
set rs=Dvbbs.Execute("select usergroupid,title from dv_usergroups where ParentGID=0 order by usergroupid")
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(1)&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>�û��ȼ�</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="userclass">
<option value=0>����</option>
<%
set rs=Dvbbs.Execute("select usertitle from dv_usergroups order by usergroupid")
do while not rs.eof
response.write "<option value="&rs(0)&">"&rs(0)&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>Email����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userEmail" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>�û�IM����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userim" type=text> ������ҳ��OICQ��UC��ICQ��YAHOO��AIM��MSN</td>
</tr>
<tr>
<td width=20% class=forumrow>��¼IP����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="lastip" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>ͷ�ΰ���</td>
<td width=80% class=forumrow colspan=5><input size=45 name="usertitle" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>ǩ������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="sign" type=text></td>
</tr>
<tr>
<td width=20% class=forumrow>��ϸ���ϰ���</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userinfo" type=text></td>
</tr>
<!--shinzeal������������-->
<tr>
<th align=left colspan=7 height=23>�����ѯ&nbsp;��ע�⣺ <����> �� <����> ��Ĭ�ϰ��� <����>������������ʹ�ô����� ��</th>
</tr>
<tr>
<td width=100% class=forumrow colspan=6><table ID="Table1">
<tr>
<td width=50%>��¼����:<input type=radio value=more name="loginR" checked ID="Radio1">&nbsp;����&nbsp;<input type=radio value=less name="loginR" ID="Radio2">&nbsp;����&nbsp;&nbsp;<input size=5 name="loginT" type=text ID="Text1"> ��&nbsp;&nbsp;</td>
<td width=50%>��ʧ����:<input type=radio value=more name="vanishR" checked ID="Radio3">&nbsp;����&nbsp;<input type=radio value=less name="vanishR" ID="Radio4">&nbsp;����&nbsp;&nbsp;<input size=5 name="vanishT" type=text ID="Text2"> ��&nbsp;&nbsp;</td>
</tr>

<tr>
<td width=50%>ע������:<input type=radio value=more name="regR" checked ID="Radio5">&nbsp;����&nbsp;<input type=radio value=less name="regR" ID="Radio6">&nbsp;����&nbsp;&nbsp;<input size=5 name="regT" type=text ID="Text3"> ��&nbsp;&nbsp;</td>
<td width=50%>��������:<input type=radio value=more name="artcleR" checked ID="Radio7">&nbsp;����&nbsp;<input type=radio value=less name="artcleR" ID="Radio8">&nbsp;����&nbsp;&nbsp;<input size=5 name="artcleT" type=text ID="Text4"> ƪ&nbsp;&nbsp;</td>
</tr><!--������������-->
<tr>
<td width=100% class=forumrow align=center colspan=6><input name="submit" type=submit value="   ��  ��   "></td>
</tr>
<input type=hidden value="9" name="userSearch">
</form>
<%
elseif request("action")="userSearch" then
%>
<tr>
<th colspan=8 align=left height=23>�������</th>
</tr>
<%

	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	currentPage=request("page")
	if currentpage="" or not IsNumeric(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	Set rs= Server.CreateObject("ADODB.Recordset")
	Select Case Request("UserSearch")
	case 1
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid order by u.UserID desc"
	case 2
		sql="select top 100  u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid order by u.UserPost desc"
	case 3
		sql="select top 100 u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid order by u.UserPost"
	case 4
		If IsSqlDataBase=1 Then
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where datediff(hour,u.LastLogin,"&SqlNowString&")<25 order by u.lastlogin desc"
		else
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where datediff('h',u.LastLogin,"&SqlNowString&")<25 order by u.lastlogin desc"
		end if
	case 5
		If IsSqlDataBase=1 Then
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where datediff(hour,u.JoinDate,"&SqlNowString&")<25 order by u.UserID desc"
		else
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where datediff('h',u.JoinDate,"&SqlNowString&")<25 order by u.UserID desc"
		end if
	case 6
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid=5 order by u.UserID desc"
	case 7
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid=6 order by u.UserID desc"
	case 8
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid<4 order by u.usergroupid"
	case 11
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where u.lockuser=2 order by userid desc"
	case 12
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where u.lockuser=1 order by userid desc"
	case 13
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where u.IsChallenge=1 order by userid desc"
	case 14
		sql="select distinct u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u Inner Join Dv_UserAccess UC On u.UserID=UC.uc_UserID order by userid desc"
	case 9
		sqlstr=""
		if request("username")<>"" then
			if request("usernamechk")="yes" then
			sqlstr=" u.username='"&request("username")&"'"
			else
			sqlstr=" u.username like '%"&request("username")&"%'"
			end if
		end if
		if cint(request("usergroups"))>0 then
			if sqlstr="" then
			sqlstr=" u.usergroupid="&request("usergroups")&""
			else
			sqlstr=sqlstr & " and u.usergroupid="&request("usergroups")&""
			end if
		end if
		if request("userclass")<>"0" then
			if sqlstr="" then
			sqlstr=" u.userclass='"&request("userclass")&"'"
			else
			sqlstr=sqlstr & " and u.userclass='"&request("userclass")&"'"
			end if
		end if
		if request("useremail")<>"" then
			if sqlstr="" then
			sqlstr=" u.useremail like '%"&request("useremail")&"%'"
			else
			sqlstr=sqlstr & " and u.useremail like '%"&request("useremail")&"%'"
			end if
		end if
		if request("userim")<>"" then
			if sqlstr="" then
			sqlstr=" u.UserIM like '%"&request("userim")&"%'"
			else
			sqlstr=sqlstr & " and u.UserIM like '%"&request("userim")&"%'"
			end if
		end if
		if request("lastip")<>"" then
			if sqlstr="" then
			sqlstr=" u.UserLastIP like '%"&request("lastip")&"%'"
			else
			sqlstr=sqlstr & " and u.UserLastIP like '%"&request("lastip")&"%'"
			end if
		end if
		if request("userinfo")<>"" then
			if sqlstr="" then
			sqlstr=" u.UserInfo like '%"&request("userinfo")&"%'"
			else
			sqlstr=sqlstr & " and u.UserInfo like '%"&request("userinfo")&"%'"
			end if
		end if
		if request("title")<>"" then
			if sqlstr="" then
			sqlstr=" u.usertitle like '%"&request("title")&"%'"
			else
			sqlstr=sqlstr & " and u.usertitle like '%"&request("title")&"%'"
			end if
		end if
		if request("sign")<>"" then
			if sqlstr="" then
			sqlstr=" u.usersign like '%"&request("sign")&"%'"
			else
			sqlstr=sqlstr & " and u.usersign like '%"&request("sign")&"%'"
			end if
		end if

'======shinzeal������������=======
		dim Tsqlstr
		if request("loginT")<>"" then
		   	if request("loginR")="more" then
			 Tsqlstr=" u.userlogins >= "&request("loginT")&""
			else
			 Tsqlstr=" u.userlogins <= "&request("loginT")&""
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & "and" & Tsqlstr
			end if 
		end if

		if request("vanishT")<>"" then
		   	if request("vanishR")="more" then
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,u.lastlogin,"&SqlNowString&") >= "&request("vanishT")&""
				Else
					Tsqlstr=" datediff('d',u.lastlogin,"&SqlNowString&") >= "&request("vanishT")&""
				End If
			else
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,u.lastlogin,"&SqlNowString&") <= "&request("vanishT")&""
				Else
					Tsqlstr=" datediff('d',u.lastlogin,"&SqlNowString&") <= "&request("vanishT")&""
				End If
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & "and" & Tsqlstr
			end if 
		end if

		if request("regT")<>"" then
		   	if request("regR")="more" then
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,u.JoinDate,"&SqlNowString&") >= "&request("regT")&""
				Else
					Tsqlstr=" datediff('d',u.JoinDate,"&SqlNowString&") >= "&request("regT")&""
				End If
			else
				If IsSqlDataBase=1 Then
					Tsqlstr=" datediff(d,u.JoinDate,"&SqlNowString&") <= "&request("regT")&""
				Else
					Tsqlstr=" datediff('d',u.JoinDate,"&SqlNowString&") <= "&request("regT")&""
				End If
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & "and" & Tsqlstr
			end if 
		end if

		if request("artcleT")<>"" then
		   	if request("artcleR")="more" then
			 Tsqlstr=" u.UserPost >= "&request("artcleT")&""
			else
			 Tsqlstr=" u.UserPost <= "&request("artcleT")&""
			end if 	
			if sqlstr="" then 
			  sqlstr=Tsqlstr
			else
			  sqlstr=sqlstr & "and" & Tsqlstr
			end if 
		end if
'======������������======
		If Sqlstr = "" Then
			Response.Write "<tr><td colspan=8 class=forumrow>��ָ������������</td></tr>"
			Response.End
		End If
		If Request("Searchmax") = "" Or Not Isnumeric(Request("Searchmax")) Then
			Sql = "SELECT TOP 1 U.Userid, U.Username, U.Useremail, U.LastLogin, U.UserLastIP, U.UserPost, U.UserGroupID FROM [Dv_User] U INNER JOIN Dv_UserGroups G ON U.Usergroupid = G.Usergroupid WHERE " & Sqlstr & " ORDER BY U.UserID DESC"
		Else
			Sql = "SELECT TOP " & Request("Searchmax") & " U.Userid, U.Username, U.Useremail, U.LastLogin, U.UserLastIP, U.UserPost, U.UserGroupID FROM [Dv_User] U INNER JOIN Dv_UserGroups G ON U.Usergroupid = G.usergroupid WHERE " & Sqlstr & " ORDER BY U.UserID DESC"
		End If
	case 10
		sql="select u.userid,u.username,u.useremail,u.LastLogin,u.UserLastIP,u.UserPost,u.UserGroupID from [dv_user] u inner join dv_UserGroups G on u.usergroupid=g.usergroupid where u.usergroupid="&request("usergroupid")&" order by u.UserID desc"
	case else
		Response.Write "<tr><td colspan=8 class=forumrow>����Ĳ�����</td></tr>"
		Response.End
	End Select
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 class=forumrow>û���ҵ���ؼ�¼��</td></tr>"
	else
%>
<FORM METHOD=POST ACTION="?action=touser">
<tr align=center>
<td class=forumRowHighlight height=23><B>�û���</B></td>
<td class=forumRowHighlight><B>Email</B></td>
<td class=forumRowHighlight><B>Ȩ��</B></td>
<td class=forumRowHighlight><B>�����޸�</B></td>
<td class=forumRowHighlight><B>���IP</B></td>
<td class=forumRowHighlight><B>����¼</B></td>
<td class=forumRowHighlight><B>����</B></td>
</tr>
<%
		rs.PageSize = Cint(Dvbbs.Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Dvbbs.Forum_Setting(11)))
%>
<tr>
<td class=forumrow><a href="?action=modify&userid=<%=rs("userid")%>"><%=rs("username")%></a></td>
<td class=forumrow width=30% ><a href="mailto:<%=rs("useremail")%>"><%=rs("useremail")%></a></td>
<td class=forumrow width=8% align=center><a href="?action=UserPermission&userid=<%=rs("userid")%>&username=<%=rs("username")%>">�༭</a></td>
<td class=forumrow width=8% align=center><a href="?action=fix&userid=<%=rs("userid")%>&username=<%=rs("username")%>">�޸�</a></td>
<td class=forumrow width=20% ><a href="admin_lockIP.asp?userip=<%=rs("UserLastIP")%>" title="����������û�IP"><%=rs("userlastip")%></a></td>
<td class=forumrow width=15% ><%if rs("lastlogin")<>"" and isdate(rs("lastlogin")) then%><%=Formatdatetime(rs("lastlogin"),2)%><%end if%></td>
<td class=forumrow align=center><input type="checkbox" name="userid" value="<%=rs("userid")%>" <%if rs("userGroupid")=1 then response.write "disabled"%>></td>
</tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
Pcount=rs.PageCount
%>
<tr><td colspan=7 class=forumrow align=center>��ҳ��
<%
	if currentpage > 4 then
'shinzeal�������������ķ�ҳ����
	response.write "<a href=""?page=1&userSearch="&request("userSearch")&"&username="&request("username")&"&useremail="&request("useremail")&"&homepage="&request("homepage")&"&oicq="&request("oicq")&"&icq="&request("icq")&"&msn="&request("msn")&"&title="&request("title")&"&sign="&request("sign")&"&userclass="&request("userclass")&"&usergroups="&request("usergroups")&"&action="&request("action")&"&usergroupid="&request("usergroupid")&"&loginR="&request("loginR")&"&loginT="&request("loginT")&"&vanishR="&request("vanishR")&"&vanishT="&request("vanishT")&"&regR="&request("regR")&"&regT="&request("regT")&"&artcleR="&request("artcleR")&"&artcleT="&request("artcleT")&"&searchmax="&request("searchmax")&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=red>["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&"&userSearch="&request("userSearch")&"&username="&request("username")&"&useremail="&request("useremail")&"&homepage="&request("homepage")&"&oicq="&request("oicq")&"&icq="&request("icq")&"&msn="&request("msn")&"&title="&request("title")&"&sign="&request("sign")&"&userclass="&request("userclass")&"&usergroups="&request("usergroups")&"&action="&request("action")&"&usergroupid="&request("usergroupid")&"&loginR="&request("loginR")&"&loginT="&request("loginT")&"&vanishR="&request("vanishR")&"&vanishT="&request("vanishT")&"&regR="&request("regR")&"&regT="&request("regT")&"&artcleR="&request("artcleR")&"&artcleT="&request("artcleT")&"&searchmax="&request("searchmax")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&userSearch="&request("userSearch")&"&username="&request("username")&"&useremail="&request("useremail")&"&homepage="&request("homepage")&"&oicq="&request("oicq")&"&icq="&request("icq")&"&msn="&request("msn")&"&title="&request("title")&"&sign="&request("sign")&"&userclass="&request("userclass")&"&usergroups="&request("usergroups")&"&action="&request("action")&"&usergroupid="&request("usergroupid")&"&loginR="&request("loginR")&"&loginT="&request("loginT")&"&vanishR="&request("vanishR")&"&vanishT="&request("vanishT")&"&regR="&request("regR")&"&regT="&request("regT")&"&artcleR="&request("artcleR")&"&artcleT="&request("artcleT")&"&searchmax="&request("searchmax")&""">["&Pcount&"]</a>"
'shinzeal����������ҳ����������
	end if
%>
</td></tr>
<tr><td colspan=5 class=forumrow align=center><B>��ѡ������Ҫ���еĲ���</B>��ɾ��<input type="radio" name="useraction" value=1>&nbsp;&nbsp;ɾ���û���������<input type="radio" name="useraction" value=3>&nbsp;&nbsp;�ƶ����û���<input type="radio" name="useraction" value=2 checked>
<select size=1 name="selusergroup">
<%
set trs=Dvbbs.Execute("select usergroupid,title from dv_usergroups where not (usergroupid=1 or usergroupid=7) and ParentGID=0 order by usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)&">"&trs(1)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing

%>
</select>
</td>
<td class=forumrow colspan=8 align=center>ȫ��ѡ��<input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)">
</td>
</tr>
<tr><td colspan=8 class=forumrow align=center>
<input type=submit name=submit value="ִ��ѡ���Ĳ���"  onclick="{if(confirm('ȷ��ִ��ѡ��Ĳ�����?')){return true;}return false;}">
</td></tr>
</FORM>
<%
	end if
	rs.close
	set rs=nothing	
elseif request("action")="touser" then
	response.write "<tr><th colspan=8 height=23 align=left>ִ�н��</th></tr>"
	if request("useraction")="" then
		response.write "<tr><td colspan=8 class=forumrow>��ָ����ز�����</td></tr>"
		founderr=true
	end if
	if request("userid")="" then
		response.write "<tr><td colspan=8 class=forumrow>��ѡ������û���</td></tr>"
		founderr=true
	end if
	if not founderr then
		if request("useraction")=1 then
			'------------------shinzeal����ɾ���û��Ķ���-------------------------
			dim uid
			for i=1 to request("userid").count
				if request("userid").count=1 then
				uID=request("userid")
				else
				uID=replace(request.form("userid")(i),"'","")
				end if
				set rs=Dvbbs.Execute("select username from [dv_User] where userid="&uid&"")
				if not (rs.eof and rs.bof) then
					Dvbbs.Execute("update dv_message set delR=1 where incept='"&trim(rs(0))&"' and delR=0")
					Dvbbs.Execute("update dv_message set delS=1 where sender='"&trim(rs(0))&"' and delS=0 and issend=0")
					Dvbbs.Execute("update dv_message set delS=1 where sender='"&trim(rs(0))&"' and delS=0 and issend=1")
					Dvbbs.Execute("delete from dv_message where incept='"&rs(0)&"' and delR=1") 
					Dvbbs.Execute("update dv_message set delS=2 where sender='"&trim(rs(0))&"' and delS=1")
					Dvbbs.Execute("delete from dv_friend where F_username='"&rs(0)&"'") 
					Dvbbs.Execute("delete from dv_bookmark where username='"&rs(0)&"'") 
				end if 
				rs.close
			next
			'-------------------ɾ���û��Ķ���------------------------
			'ɾ���û������Ӻ;���
			Dvbbs.Execute("delete from dv_topic where PostUserID in ("&replace(request("userid"),"'","")&")")
			for i=0 to ubound(allposttable)
				Dvbbs.Execute("delete from "&allposttable(i)&" where PostUserID in ("&replace(request("userid"),"'","")&")")
			next
			Dvbbs.Execute("delete from dv_besttopic where PostUserID in ("&replace(request("userid"),"'","")&")")
			'ɾ���û��ϴ���
			Dvbbs.Execute("delete from dv_upfile where F_UserID in ("&replace(request("userid"),"'","")&")")
			Dvbbs.Execute("delete from [dv_user] where userid in ("&replace(request("userid"),"'","")&")")
			response.write "<tr><td colspan=8 class=forumrow>�����ɹ���</td></tr>"
		elseif request("useraction")=2 then
			dim userclass,usertitlepic
			set rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&request("selusergroup")&" order by minarticle")
			if not (rs.eof and rs.bof) then
				userclass=rs("usertitle")
				usertitlepic=rs("grouppic")
			end if
			Dvbbs.Execute("update [dv_user] set UserGroupID="&replace(request("selusergroup"),"'","")&",userclass='"&userclass&"',titlepic='"&usertitlepic&"' where userid in ("&replace(request("userid"),"'","")&")")
			response.write "<tr><td colspan=8 class=forumrow>�����ɹ���</td></tr>"
		elseif request("useraction")=3 then
			dim titlenum
			if request("userid")="" then
				response.write "<tr><td colspan=8 class=forumrow>�����뱻ɾ�������û�����</td></tr>"
			end if
			titlenum=0
			for i=0 to ubound(allposttable)
			set rs=Dvbbs.Execute("Select Count(announceID) from "&allposttable(i)&" where postuserid in ("&replace(request("userid"),"'","")&")") 
   			titlenum=titlenum+rs(0)
			sql="update "&allposttable(i)&" set locktopic=boardid,boardid=444,isbest=0 where postuserid in ("&replace(request("userid"),"'","")&")"
			Dvbbs.Execute(sql)
			next
			Dvbbs.Execute("delete from dv_besttopic where postuserid in ("&replace(request("userid"),"'","")&")")
			set rs=Dvbbs.Execute("select topicid,posttable from dv_topic where postuserid in ("&replace(request("userid"),"'","")&")")
			do while not rs.eof
			Dvbbs.Execute("update "&rs(1)&" set locktopic=boardid,boardid=444,isbest=0 where rootid="&rs(0))
			rs.movenext
			loop
			set rs=nothing
			Dvbbs.Execute("update dv_topic set locktopic=boardid,boardid=444,isbest=0 where postuserid in ("&replace(request("userid"),"'","")&")")
			if isnull(titlenum) then titlenum=0
			sql="update [dv_user] set UserPost=UserPost-"&titlenum&",userWealth=userWealth-"&titlenum*Dvbbs.Forum_user(3)&",userEP=userEP-"&titlenum*Dvbbs.Forum_user(8)&",userCP=userCP-"&titlenum*Dvbbs.Forum_user(13)&" where userid in ("&replace(request("userid"),"'","")&")"
			Dvbbs.Execute(sql)
			response.write "<tr><td colspan=8 class=forumrow>ɾ���ɹ������Ҫ��ȫɾ�������뵽��̳����վ<BR>��������������̳�����и���һ����̳���ݣ�����<a href=admin_alldel.asp>����</a></td></tr>"
		else
			response.write "<tr><td colspan=8 class=forumrow>����Ĳ�����</td></tr>"
		end if
	end if
elseif request("action")="modify" then
dim realname,character,personal,country,province,city,shengxiao,blood,belief,occupation,marital, education,college,userphone,iaddress
Dim UserIM
	response.write "<tr><th colspan=8 height=23 align=left>�û����ϲ���</th></tr>"
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=8 class=forumrow>������û�������</td></tr>"
		founderr=true
	end if
	if not founderr then
		Set rs= Server.CreateObject("ADODB.Recordset")
		sql="select * from [dv_user] where userid="&request("userid")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 class=forumrow>û���ҵ�����û���</td></tr>"
		founderr=true
		else
if rs("userinfo")<>"" then
	userinfo=split(Server.HtmlEncode(rs("userinfo")),"|||")
	if ubound(userinfo)=14 then
		realname=userinfo(0)
		character=userinfo(1)
		personal=userinfo(2)
		country=userinfo(3)
		province=userinfo(4)
		city=userinfo(5)
		shengxiao=userinfo(6)
		blood=userinfo(7)
		belief=userinfo(8)
		occupation=userinfo(9)
		marital=userinfo(10)
		education=userinfo(11)
		college=userinfo(12)
		userphone=userinfo(13)
		iaddress=userinfo(14)
	else
		realname=""
		character=""
		personal=""
		country=""
		province=""
		city=""
		shengxiao=""
		blood=""
		belief=""
		occupation=""
		marital=""
		education=""
		college=""
		userphone=""
		iaddress=""
	end if
else
	realname=""
	character=""
	personal=""
	country=""
	province=""
	city=""
	shengxiao=""
	blood=""
	belief=""
	occupation=""
	marital=""
	education=""
	college=""
	userphone=""
	iaddress=""
end if
UserIM = Split(Rs("UserIM"),"|||")
%>
<FORM METHOD=POST ACTION="?action=saveuserinfo">
<tr>
<td width=100% class=forumrow valign=top colspan=8>�� <%=rs("username")%> �û��������ѡ�<BR><BR>
<a href="mailto:<%=rs("useremail")%>">���ʼ�</a> | <a href="messanger.asp?action=new&touser=<%=rs("username")%>" target=_blank>������</a> | <a href="dispuser.asp?id=<%=rs("userid")%>" target=_blank>Ԥ���û�����</a> | <a href="Query.asp?stype=1&nSearch=3&keyword=<%=rs("username")%>&SearchDate=30" target=_blank>�û�����</a> | <a href="Query.asp?stype=6&nSearch=0&pSearch=0&keyword=<%=rs("username")%>" target=_blank>�û�����</a> | <a href="Query.asp?stype=4&nSearch=0&pSearch=0&keyword=<%=rs("username")%>" target=_blank>�û�����</a> | <a href="show.asp?username=admin" target=_blank>�û�չ��</a> | <a href="?action=UserPermission&userid=<%=rs("userid")%>&username=<%=rs("username")%>">�༭Ȩ��</a> | <a href="look_ip.asp?action=lookip&ip=<%=Rs("UserLastIP")%>" target=_blank>�����Դ</a> | <a href="?action=touser&useraction=1&userid=<%=rs("userid")%>" onclick="{if(confirm('ɾ�������ɻָ������ҽ�ɾ�����û�����̳��������Ϣ��ȷ��ɾ����?')){return true;}return false;}">ɾ���û�</a>
</td>
</tr>
<tr><th colspan=6 height=23 align=left>�û����������޸ģ���<%=rs("username")%></th></tr>
<tr>
<td width=20% class=forumrow>�û���</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="usergroups">
<%
set trs=Dvbbs.Execute("select usergroupid,title,parentgid from dv_usergroups where IsSetting=1 order by usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)
if rs("usergroupid")=trs(0) then response.write " selected "
response.write ">"&trs(1)
if trs(2)>0 then response.write "(�Զ���ȼ�)"
response.write "</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
</select>
</td>
</tr>
<input name="userid" type=hidden value="<%=rs("userid")%>">
<tr>
<td width=20% class=forumrow>�û���</td>
<td width=80% class=forumrow colspan=5><input size=45 name="username" type=text value="<%=Server.HtmlEncode(rs("username"))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��  ��</td>
<td width=80% class=forumrow colspan=5><input size=45 name="password" type=text>&nbsp;������޸�������</td>
</tr>
<tr>
<td width=20% class=forumrow>��������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="quesion" type=text value="<%If Trim(rs("userquesion"))<>"" Then Response.Write Server.HtmlEncode(rs("userquesion"))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>�����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="answer" type=text>&nbsp;������޸�������</td>
</tr>
<tr>
<td width=20% class=forumrow>�û��ȼ�</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name="userclass">
<%
set trs=Dvbbs.Execute("select usertitle from dv_usergroups order by usergroupid")
do while not trs.eof
response.write "<option value="&trs(0)
if rs("userclass")=trs(0) then response.write " selected "
response.write ">"&trs(0)&"</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>Email</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userEmail" type=text value="<%If Trim(rs("useremail"))<>"" Then Response.Write Server.HtmlEncode(rs("useremail"))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>������ҳ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="homepage" type=text value="<%=Server.HtmlEncode(UserIM(0))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>ͷ��</td>
<td width=80% class=forumrow colspan=5><input size=45 name="face" type=text value="<%If Trim(Rs("UserFace"))<>"" Then Response.Write Server.HtmlEncode(rs("userface"))%>">&nbsp;��ȣ�<input size=3 name="width" type=text value="<%=rs("userwidth")%>">&nbsp;�߶ȣ�<input size=3 name="height" type=text value="<%=rs("userheight")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>OICQ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="oicq" type=text value="<%=Server.HtmlEncode(UserIM(1))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>ICQ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="icq" type=text value="<%=Server.HtmlEncode(UserIM(2))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>MSN</td>
<td width=80% class=forumrow colspan=5><input size=45 name="msn" type=text value="<%=Server.HtmlEncode(UserIM(3))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>AIM</td>
<td width=80% class=forumrow colspan=5><input size=45 name="aim" type=text value="<%=Server.HtmlEncode(UserIM(4))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>YaHoo</td>
<td width=80% class=forumrow colspan=5><input size=45 name="yahoo" type=text value="<%=Server.HtmlEncode(UserIM(5))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>UC</td>
<td width=80% class=forumrow colspan=5><input size=45 name="uc" type=text value="<%=Server.HtmlEncode(UserIM(6))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>ͷ��</td>
<td width=80% class=forumrow colspan=5><input size=45 name="usertitle" type=text value="<%If Trim(Rs("UserTitle"))<>"" Then Response.Write Server.HtmlEncode(rs("usertitle"))%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>�ȼ�ͼƬ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="titlepic" type=text value="<%=rs("titlepic")%>"></td>
</tr>
<tr><th colspan=6 height=23 align=left>�û���ֵ�����޸�</th></tr>
<tr>
<td width=20% class=forumrow>��������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="article" type=text value="<%=rs("UserPost")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��ɾ����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="Userdel" type=text value="<%=rs("userdel")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userisbest" type=text value="<%=rs("userisbest")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��Ǯ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userwealth" type=text value="<%=rs("userwealth")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userep" type=text value="<%=rs("userep")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="usercp" type=text value="<%=rs("usercp")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userpower" type=text value="<%=rs("userpower")%>"></td>
</tr>
<tr><th colspan=6 height=23 align=left>�������</th></tr>
<tr>
<td width=20% class=forumrow>����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="birthday" type=text value="<%=rs("userbirthday")%>">&nbsp;��ʽ��2001-2-2</td>
</tr>
<tr>
<td width=20% class=forumrow>ע��ʱ��</td>
<td width=80% class=forumrow colspan=5><input size=45 name="adddate" type=text value="<%=rs("JoinDate")%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>����¼</td>
<td width=80% class=forumrow colspan=5><input size=45 name="lastlogin" type=text value="<%=rs("lastlogin")%>"></td>
</tr>
<tr><th colspan=6 height=23 align=left>�û���ϸ����</th></tr>
<tr>
<td width=20% class=forumrow>��ʵ����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="realname" type=text value="<%=realname%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="country" type=text value="<%=country%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>��ϵ�绰</td>
<td width=80% class=forumrow colspan=5><input size=45 name="userphone" type=text value="<%=userphone%>"></td>
</tr><tr>
<td width=20% class=forumrow>ͨ�ŵ�ַ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="address" type=text value="<%=iaddress%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>ʡ������</td>
<td width=80% class=forumrow colspan=5><input size=45 name="province" type=text value="<%=province%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>�ǡ�����</td>
<td width=80% class=forumrow colspan=5><input size=45 name="city" type=text value="<%=city%>"></td>
</tr><tr>
<td width=20% class=forumrow>������Ф</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=shengxiao>
<option <%if shengxiao="" then%>selected<%end if%>></option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=ţ <%if shengxiao="ţ" then%>selected<%end if%>>ţ</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
<option value=�� <%if shengxiao="��" then%>selected<%end if%>>��</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>Ѫ������</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=blood>
<option <%if blood="" then%>selected<%end if%>></option>
<option value=A <%if blood="A" then%>selected<%end if%>>A</option>
<option value=B <%if blood="B" then%>selected<%end if%>>B</option>
<option value=AB <%if blood="AB" then%>selected<%end if%>>AB</option>
<option value=O <%if blood="O" then%>selected<%end if%>>O</option>
<option value=���� <%if blood="����" then%>selected<%end if%>>����</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>�š�����</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=belief>
<option <%if belief="" then%>selected<%end if%>></option>
<option value=��� <%if belief="���" then%>selected<%end if%>>���</option>
<option value=���� <%if belief="����" then%>selected<%end if%>>����</option>
<option value=������ <%if belief="������" then%>selected<%end if%>>������</option>
<option value=������ <%if belief="������" then%>selected<%end if%>>������</option>
<option value=�ؽ� <%if belief="�ؽ�" then%>selected<%end if%>>�ؽ�</option>
<option value=�������� <%if belief="��������" then%>selected<%end if%>>��������</option>
<option value=���������� <%if belief="����������" then%>selected<%end if%>>����������</option>
<option value=���� <%if belief="����" then%>selected<%end if%>>����</option>
</select>
</td>
</tr><tr>
<td width=20% class=forumrow>ְ����ҵ</td>
<td width=80% class=forumrow colspan=5>
<select name=occupation>
<option <%if occupation="" then%>selected<%end if%>> </option>
<option value="�ƻ�/����" <%if occupation="�ƻ�/����" then%>selected<%end if%>>�ƻ�/����</option>
<option value=����ʦ <%if occupation="����ʦ" then%>selected<%end if%>>����ʦ</option>
<option value=���� <%if occupation="����" then%>selected<%end if%>>����</option>
<option value=����������ҵ <%if occupation="����������ҵ" then%>selected<%end if%>>����������ҵ</option>
<option value=��ͥ���� <%if occupation="��ͥ����" then%>selected<%end if%>>��ͥ����</option>
<option value="����/��ѵ" <%if occupation="����/��ѵ" then%>selected<%end if%>>����/��ѵ</option>
<option value="�ͻ�����/֧��" <%if occupation="�ͻ�����/֧��" then%>selected<%end if%>>�ͻ�����/֧��</option>
<option value="������/�ֹ�����" <%if occupation="������/�ֹ�����" then%>selected<%end if%>>������/�ֹ�����</option>
<option value=���� <%if occupation="����" then%>selected<%end if%>>����</option>
<option value=��ְҵ <%if occupation="��ְҵ" then%>selected<%end if%>>��ְҵ</option>
<option value="����/�г�/���" <%if occupation="����/�г�/���" then%>selected<%end if%>>����/�г�/���</option>
<option value=ѧ�� <%if occupation="ѧ��" then%>selected<%end if%>>ѧ��</option>
<option value=�о��Ϳ��� <%if occupation="�о��Ϳ���" then%>selected<%end if%>>�о��Ϳ���</option>
<option value="һ�����/�ල" <%if occupation="һ�����/�ල" then%>selected<%end if%>>һ�����/�ල</option>
<option value="����/����" <%if occupation="����/����" then%>selected<%end if%>>����/����</option>
<option value="ִ�й�/�߼�����" <%if occupation="ִ�й�/�߼�����" then%>selected<%end if%>>ִ�й�/�߼�����</option>
<option value="����/����/����" <%if occupation="����/����/����" then%>selected<%end if%>>����/����/����</option>
<option value=רҵ��Ա <%if occupation="רҵ��Ա" then%>selected<%end if%>>רҵ��Ա</option>
<option value="�Թ�/ҵ��" <%if occupation="�Թ�/ҵ��" then%>selected<%end if%>>�Թ�/ҵ��</option>
<option value=���� <%if occupation="����" then%>selected<%end if%>>����</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>����״��</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=marital>
<option <%if marital="" then%>selected<%end if%>></option>
<option value=δ�� <%if marital="δ��" then%>selected<%end if%>>δ��</option>
<option value=�ѻ� <%if marital="�ѻ�" then%>selected<%end if%>>�ѻ�</option>
<option value=���� <%if marital="����" then%>selected<%end if%>>����</option>
<option value=ɥż <%if marital="ɥż" then%>selected<%end if%>>ɥż</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>���ѧ��</td>
<td width=80% class=forumrow colspan=5>
<select size=1 name=education>
<option <%if education="" then%>selected<%end if%>></option>
<option value=Сѧ <%if education="Сѧ" then%>selected<%end if%>>Сѧ</option>
<option value=���� <%if education="����" then%>selected<%end if%>>����</option>
<option value=���� <%if education="����" then%>selected<%end if%>>����</option>
<option value=��ѧ <%if education="��ѧ" then%>selected<%end if%>>��ѧ</option>
<option value=˶ʿ <%if education="˶ʿ" then%>selected<%end if%>>˶ʿ</option>
<option value=��ʿ <%if education="��ʿ" then%>selected<%end if%>>��ʿ</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>��ҵԺУ</td>
<td width=80% class=forumrow colspan=5><input size=45 name="college" type=text value="<%=college%>"></td>
</tr>
<tr>
<td width=20% class=forumrow>�ԡ���</td>
<td width=80% class=forumrow colspan=5>
<textarea name=character rows=4 cols=80><%=character%></textarea>
</td>
</tr><tr>
<td width=20% class=forumrow>���˼��</td>
<td width=80% class=forumrow colspan=5>
<textarea name=personal rows=4 cols=80><%=personal%></textarea>
</td>
</tr><tr>
<td width=20% class=forumrow>�û�ǩ��</td>
<td width=80% class=forumrow colspan=5>
<textarea name="sign" rows=4 cols=80><%If Trim(Rs("UserSign"))<>"" Then Response.Write Server.HtmlEncode(rs("usersign"))%></textarea>
</td>
</tr>
<tr><th colspan=6 height=23 align=left>�û�����</th></tr>
<tr>
<td width=20% class=forumrow>�û�״̬</td>
<td width=80% class=forumrow colspan=5>
���� <input type="radio" value="0" <%if rs("lockuser")=0 then%>checked<%end if%> name="lockuser">&nbsp;
���� <input type="radio" value="1" <%if rs("lockuser")=1 then%>checked<%end if%> name="lockuser">&nbsp;
���� <input type="radio" value="2" <%if rs("lockuser")=2 then%>checked<%end if%> name="lockuser">
</td>
</tr>
<tr>
<td width=100% class=forumrow align=center colspan=6><input name="submit" type=submit value="   ��  ��   "></td>
</tr>
</FORM>
<%
		end if
		rs.close
		set rs=nothing
	end if
elseif request("action")="saveuserinfo" then
	response.write "<tr><th colspan=8 height=23 align=left>�����û�����</th></tr>"
	userinfo=checkreal(request.Form("realname")) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
	dim myuserim
	myuserim=checkreal(request.Form("homepage")) & "|||" & checkreal(request.Form("oicq")) & "|||" & checkreal(request.Form("icq")) & "|||" & checkreal(request.Form("msn")) & "|||" & checkreal(request.Form("aim")) & "|||" & checkreal(request.Form("yahoo")) & "|||" & request.Form("uc")
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=8 class=forumrow>������û�������</td></tr>"
		founderr=true
	end if
	if not founderr then
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select * from [dv_user] where userid="&request("userid")
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 class=forumrow>û���ҵ�����û���</td></tr>"
		founderr=true
	else
		rs("username")=request.form("username")
		if request.form("password")<>"" then
		rs("userpassword")=md5(request.form("password"),16)
		end if
		rs("usergroupid")=request.form("usergroups")
		rs("userquesion")=request.form("quesion")
		if request.form("answer")<>"" then rs("useranswer")=md5(request.form("answer"),16)
		rs("userclass")=request.form("userclass")
		rs("useremail")=request.form("useremail")
		rs("userim")=myuserim
		rs("userface")=request.form("face")
		if isnumeric(request.form("width")) then rs("userwidth")=request.form("width")
		if isnumeric(request.form("height")) then rs("userheight")=request.form("height")
		rs("usertitle")=request.form("usertitle")
		rs("titlepic")=request.form("titlepic")
		if isnumeric(request.form("article")) then rs("UserPost")=request.form("article")
		if isnumeric(request.form("userdel")) then rs("userdel")=request.form("userdel")
		if isnumeric(request.form("userisbest")) then rs("userisbest")=request.form("userisbest")
		if isnumeric(request.form("userpower")) then rs("userpower")=request.form("userpower")
		if isnumeric(request.form("userwealth")) then rs("userwealth")=request.form("userwealth")
		if isnumeric(request.form("userep")) then rs("userep")=request.form("userep")
		if isnumeric(request.form("usercp")) then rs("usercp")=request.form("usercp")
		if isdate(request.form("birthday")) then rs("userbirthday")=request.form("birthday")
		if isdate(request.form("adddate")) then rs("JoinDate")=request.form("adddate")
		if isdate(request.form("lastlogin")) then rs("lastlogin")=request.form("lastlogin")
		if isnumeric(request.form("lockuser")) then rs("lockuser")=request.form("lockuser")
		rs("usersign")=request.form("sign")
		rs("userinfo")=userinfo
		rs.update
	end if
	rs.close
	set rs=nothing
	end if
	if founderr then
		response.write "<tr><td colspan=8 class=forumrow>����ʧ�ܡ�</td></tr>"
	else
		response.write "<tr><td colspan=8 class=forumrow>�����û����ݳɹ���</td></tr>"
	end if
ElseIf Request("Action") = "UserPermission" Then
	Response.Write "<tr><th colspan=8 height=23 align=left>�༭" & Request("Username") & "��̳Ȩ�ޣ���ɫ��ʾ���û��ڸð������Զ���Ȩ�ޣ�</th></tr>"
	If Not Isnumeric(Request("Userid")) Then
		Response.Write "<tr><td colspan=8 class=forumrow>������û�������</td></tr>"
		Founderr = True
	End If
	If Not Founderr Then
		Response.Write "<tr><td colspan=8 class=forumrow height=25>�����������ø��û��ڲ�ͬ��̳�ڵ�Ȩ�ޣ���ɫ��ʾΪ���û���ʹ�õ����û��Զ�������<BR>�ڸ�Ȩ�޲��ܼ̳У�������������һ�������¼���̳�İ��棬��ôֻ�������õİ�����Ч��������������̳��Ч<BR>���������������Ч������������ҳ��<B>ѡ���Զ�������</B>��ѡ�����Զ������ú��������õ�Ȩ�޽�<B>����</B>���û������ú���̳Ȩ�����ã������û���Ĭ�ϻ���̳Ȩ�����ø��û��鲻�ܹ������ӣ������������˸��û��ɹ������ӣ���ô���û����������Ϳ��Թ�������</td></tr>"
		Response.Write "<tr><td colspan=8 class=forumrow height=25><a href=?action=userBoardPermission&boardid=0&userid=" & Request("Userid") & ">�༭���û�������ҳ���Ȩ��</a>����Ҫ��Զ��Ų������ã�</td></tr>"
'----------------------boardinfo--------------------
		Response.Write "<tr><td colspan=8 class=forumrow><B>�����̳���ƽ���༭״̬</B><BR>"
		Rem �����������ѭ����ѯ 2004-5-6 Dvbbs.YangZheng
		Dim Bn
		Sql = "SELECT Depth, Child, Boardid, Parentid, Boardtype FROM Dv_Board ORDER BY Rootid, Orders"
		Set Rs = Dvbbs.Execute(Sql)
		If Not (Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
			Rs.Close:Set Rs = Nothing
			For Bn = 0 To Ubound(Sql,2)
				If Sql(0,Bn) > 0 Then
					For i = 1 To Sql(0,Bn)
						Response.Write "&nbsp;"
					Next
				End If
				If Sql(1,Bn) > 0 Then
					Response.Write "<img src=""skins/default/plus.gif"">"
				Else
					Response.Write "<img src=""skins/default/nofollow.gif"">"
				End If
%>
<a href="?action=UserBoardPermission&boardid=<%=Sql(2,Bn)%>&userid=<%=Request("Userid")%>">
<%
				Set Trs = Dvbbs.Execute("SELECT Uc_UserId FROM Dv_UserAccess WHERE Uc_Boardid = " & Sql(2,Bn) & " AND Uc_Userid = " & Request("Userid"))
				If Not (Trs.Eof And Trs.Bof) Then
					Response.Write "<font color=red>"
				End If
				If Sql(3,Bn) = 0 Then Response.Write "<b>"
				Response.Write Sql(4,Bn)
				If Sql(3,Bn) = 0 Then Response.Write "</b>"
				If Sql(1,Bn) > 0 Then Response.Write "(" & Sql(1,Bn) & ")"
				Response.Write "</font></a><BR>"
			Next
		End If
		Response.Write "</td></tr>"
'-------------------end-------------------
	End If
ElseIf Request("Action") = "UserBoardPermission" Then
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=8 class=forumrow>������û�������</td></tr>"
		founderr=true
	end if
	if not isnumeric(request("boardid")) then
		response.write "<tr><td colspan=8 class=forumrow>����İ��������</td></tr>"
		founderr=true
	end if
	if not founderr then
	set rs=Dvbbs.Execute("select u.UserGroupID,ug.title,u.username from [dv_user] u inner join dv_UserGroups UG on u.userGroupID=ug.userGroupID where u.userid="&request("userid"))
	Dvbbs.UserGroupID=rs(0)
	usertitle=rs(1)
	Dvbbs.membername=rs(2)
	dim boardtype
	set rs=Dvbbs.Execute("select boardtype from dv_board where boardid="&request("boardid"))
	if rs.eof and rs.bof then
	boardtype="��̳����ҳ��"
	else
	boardtype=rs(0)
	end if
	response.write "<tr><th colspan=8 height=23 align=left>�༭ "&Dvbbs.membername&" �� "&boardtype&" Ȩ��</th></tr>"
	response.write "<tr><td colspan=8 height=25 class=forumrow>ע�⣺���û����� <B>"&usertitle&"</B> �û����У�����������������Զ���Ȩ�ޣ�����û�Ȩ�޽����Զ���Ȩ��Ϊ��</td></tr>"
%>
<tr><td colspan=8 class=forumrow>
<%
Dim reGroupSetting
Dim FoundGroup,FoundUserPermission,FoundGroupPermission
FoundGroup=false
FoundUserPermission=false
FoundGroupPermission=false

set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
if not (rs.eof and rs.bof) then
	reGroupSetting=rs("uc_Setting")
	FoundGroup=true
	FoundUserPermission=true
end if

if not foundgroup then
set rs=Dvbbs.Execute("select * from dv_BoardPermission where boardid="&request("boardid")&" and groupid="&DVbbs.UserGroupID)
if not(rs.eof and rs.bof) then
	reGroupSetting=rs("PSetting")
	FoundGroup=true
	FoundGroupPermission=true
end if
end if

if not foundgroup then
set rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&DVbbs.UserGroupID)
if rs.eof and rs.bof then
	response.write "δ�ҵ����û��飡"
	response.end
else
	FoundGroup=true
	FoundGroupPermission=true
	reGroupSetting=rs("GroupSetting")
end if
end if
%>
<table width="100%" border="0" cellspacing="1" cellpadding="0"  align=center>
<FORM METHOD=POST ACTION="?action=saveuserpermission">
<input type=hidden name="userid" value="<%=request("userid")%>">
<input type=hidden name="BoardID" value="<%=request("boardid")%>">
<input type=hidden name="username" value="<%=Dvbbs.membername%>">

<tr> 
<td width="100%" class=Forumrow colspan=2 height=25>
<font color=blue>����Ŀ��</font>��<input type=radio name="savetype" value=0 checked>�ð���&nbsp;<input type=radio name="savetype" value=1>���а���&nbsp;<input type=radio name="savetype" value=2>��ͬ���������а��棨���������ࣩ&nbsp;<input type=radio name="savetype" value=3>��ͬ���������а��棨�������ࣩ&nbsp;<input type=radio name="savetype" value=4>ͬ����ͬ�������
</td>
</tr>
<tr> 
<td width="100%" class=Forumrow colspan=2 height=25>
<font color=blue>
����ָ�ķ����ָһ�����࣬�����Ǹð�����ϼ�����</font>��������Ŀǰ���õ���һ���弶���棬ѡ������ͬ���������а��涼���£���ô���ｫ���°����÷����һ�����������������ļ����а��棬��������ĸ��·�Χ̫�󣬿���ѡ�����ͬ����ͬ������档
</td>
</tr>
<tr> 
<td height="23" colspan="2" class=forumrow><input type=radio name="isdefault" value="1" <%if FoundGroupPermission then%>checked<%end if%>><B>ʹ���û���Ĭ��ֵ</B> (ע��: �⽫ɾ���κ�֮ǰ�������Զ�������)</td>
</tr>
<tr> 
<td height="23" colspan="2"  class=forumrow><input type=radio name="isdefault" value="0" <%if FoundUserPermission then%>checked<%end if%>><B>ʹ���Զ�������</B> &nbsp;(<font color=blue>ѡ���Զ������ʹ����������Ч</font>)</td>
</tr>
<%
GroupPermission(reGroupSetting)
%>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
</td></tr>
<%
	end if
elseif request("action")="saveuserpermission" then
	response.write "<tr><th colspan=8 height=23 align=left>�༭�û� "&request("username")&" Ȩ��</th></tr>"
	if not isnumeric(request("userid")) then
		response.write "<tr><td colspan=8 class=forumrow>������û�������</td></tr>"
		founderr=true
	end if
	if not isnumeric(request("boardid")) then
		response.write "<tr><td colspan=8 class=forumrow>����İ��������</td></tr>"
		founderr=true
	end if
	if not founderr then
	dim myGroupSetting
	Dim IsGroupSetting,MyIsGroupSetting,FoundSetting
	myGroupSetting=GetGroupPermission
	select case request("savetype")
	'��ǰ����
	case "0"
		if request("isdefault")=1 then
			Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
			Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
			FoundSetting=Rs(0)
			If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&request("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = ""
			Else
				IsGroupSetting = "," & Rs(0) & ","
				If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0","")
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&request("boardid"))
		else
			set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
			if rs.eof and rs.bof then
				Dvbbs.Execute("insert into dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&request("boardid")&",'"&myGroupSetting&"')")
			else
				Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&request("boardid")&" and uc_userid="&request("userid"))
			end if
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&request("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = 0
			Else
				IsGroupSetting = "," & Rs(0) & ","
				IsGroupSetting = Replace(IsGroupSetting,",0","")
				IsGroupSetting = IsGroupSetting & "0,"
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&request("boardid"))
			Set Rs=Nothing
		end if
		Dvbbs.ReloadBoardInfo(request("boardid"))
	'���а���
	case "1"
		set trs=Dvbbs.Execute("select * from dv_board")
		do while not trs.eof
		if request("isdefault")=1 then
			Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			FoundSetting=Rs(0)
			If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = ""
			Else
				IsGroupSetting = "," & Rs(0) & ","
				If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0","")
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			FoundSetting=""
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		else
			set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			if rs.eof and rs.bof then
				Dvbbs.Execute("insert into dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&trs("boardid")&",'"&myGroupSetting&"')")
			else
				Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			end if
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = 0
			Else
				IsGroupSetting = "," & Rs(0) & ","
				IsGroupSetting = Replace(IsGroupSetting,",0","")
				IsGroupSetting = IsGroupSetting & "0,"
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		end if
		Dvbbs.ReloadBoardInfo(trs("boardid"))
		trs.movenext
		loop
		trs.close
		set trs=nothing
		Set Rs=Nothing
	'��ͬ���������а��棨���������ࣩ
	case "2"
		set trs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("boardid"))
		myrootid=trs(0)
		set trs=Dvbbs.Execute("select * from dv_board where (Not ParentID=0) and rootid="&myrootid)
		do while not trs.eof
		if request("isdefault")=1 then
			Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			FoundSetting=Rs(0)
			If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = ""
			Else
				IsGroupSetting = "," & Rs(0) & ","
				If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0","")
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			FoundSetting=""
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		else
			set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			if rs.eof and rs.bof then
				Dvbbs.Execute("insert into dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&trs("boardid")&",'"&myGroupSetting&"')")
			else
				Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			end if
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = 0
			Else
				IsGroupSetting = "," & Rs(0) & ","
				IsGroupSetting = Replace(IsGroupSetting,",0","")
				IsGroupSetting = IsGroupSetting & "0,"
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		end if
		Dvbbs.ReloadBoardInfo(trs("boardid"))
		trs.movenext
		loop
		trs.close
		set trs=nothing
		Set Rs=Nothing
	'��ͬ���������а��棨�������ࣩ
	case "3"
		set trs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("boardid"))
		myrootid=trs(0)
		set trs=Dvbbs.Execute("select * from dv_board where rootid="&myrootid)
		do while not trs.eof
		if request("isdefault")=1 then
			Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			FoundSetting=Rs(0)
			If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = ""
			Else
				IsGroupSetting = "," & Rs(0) & ","
				If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0","")
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			FoundSetting=""
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		else
			set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			if rs.eof and rs.bof then
				Dvbbs.Execute("insert into dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&trs("boardid")&",'"&myGroupSetting&"')")
			else
				Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			end if
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = 0
			Else
				IsGroupSetting = "," & Rs(0) & ","
				IsGroupSetting = Replace(IsGroupSetting,",0","")
				IsGroupSetting = IsGroupSetting & "0,"
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		end if
		Dvbbs.ReloadBoardInfo(trs("boardid"))
		trs.movenext
		loop
		trs.close
		set trs=nothing
		Set Rs=Nothing
	'ͬ����ͬ�������
	case "4"
		dim myparentid,myparentstr
		set trs=Dvbbs.Execute("select rootid,ParentStr,ParentID from dv_board where boardid="&request("boardid"))
		myrootid=trs(0)
		myparentid=trs(1)
		myparentstr=trs(2)
		set trs=Dvbbs.Execute("select * from dv_board where rootid="&myrootid&" and ParentID="&myparentid&" and ParentStr='"&myparentstr&"'")
		do while not trs.eof
		if request("isdefault")=1 then
			Dvbbs.Execute("delete from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			Set Rs=Dvbbs.Execute("Select Count(*) from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			FoundSetting=Rs(0)
			If IsNull(FoundSetting) Or FoundSetting="" Then FoundSetting=0
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = ""
			Else
				IsGroupSetting = "," & Rs(0) & ","
				If FoundSetting=0 Then IsGroupSetting = Replace(IsGroupSetting,",0","")
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			FoundSetting=""
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		else
			set rs=Dvbbs.Execute("select * from dv_UserAccess where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			if rs.eof and rs.bof then
				Dvbbs.Execute("insert into dv_UserAccess (uc_userid,uc_boardid,uc_setting) values ("&request("userid")&","&trs("boardid")&",'"&myGroupSetting&"')")
			else
				Dvbbs.Execute("update dv_UserAccess set uc_setting='"&myGroupSetting&"' where uc_boardid="&trs("boardid")&" and uc_userid="&request("userid"))
			end if
			Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&trs("boardid"))
			If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
				MyIsGroupSetting = 0
			Else
				IsGroupSetting = "," & Rs(0) & ","
				IsGroupSetting = Replace(IsGroupSetting,",0","")
				IsGroupSetting = IsGroupSetting & "0,"
				IsGroupSetting = Split(IsGroupSetting,",")
				For i=1 To Ubound(IsGroupSetting)-1
					If i=1 Then
						MyIsGroupSetting = IsGroupSetting(i)
					Else
						MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
					End If
				Next
			End If
			Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&trs("boardid"))
		end if
		Dvbbs.ReloadBoardInfo(trs("boardid"))
		trs.movenext
		loop
		trs.close
		set trs=nothing
		Set Rs=Nothing
	end select
	if founderr then
		response.write "<tr><td colspan=8 class=forumrow>����ʧ�ܡ�</td></tr>"
	else
		response.write "<tr><td colspan=8 class=forumrow>�����û�Ȩ�޳ɹ���</td></tr>"
	end if
	End if
elseif request("action")="uniteuser" then
	if request("auser")<>"" and request("buser")<>"" then
		dim auserid,buserid
		dim c1,c2,c3,c4,c5,c6,c7,c8,c9
		set rs=dvbbs.execute("select userid,userpost,usertopic,userviews,userwealth,userep,usercp,userpower,userisbest,userdel,usergroupid from dv_user where username='"&replace(request("auser"),"'","''")&"'")
		if rs.eof and rs.bof then
			errmsg = errmsg + "<tr><td colspan=8 class=forumrow>û���ҵ����ϲ��û�</td></tr>"
			founderr=true
		else
			auserid=rs(0)
			c1=rs(1)
			c2=rs(2)
			c3=rs(3)
			c4=rs(4)
			c5=rs(5)
			c6=rs(6)
			c7=rs(7)
			c8=rs(8)
			c9=rs(9)
			if rs(10)<>4 then
				errmsg = errmsg + "<tr><td colspan=8 class=forumrow>ֻ�����ע���û�����кϲ��û�����</td></tr>"
				founderr=true
			end if
		end if
		set rs=dvbbs.execute("select userid from dv_user where username='"&replace(request("buser"),"'","''")&"'")
		if rs.eof and rs.bof then
			errmsg = errmsg + "<tr><td colspan=8 class=forumrow>û���ҵ��ϲ���Ŀ���û�</td></tr>"
			founderr=true
		else
			buserid=rs(0)
		end if
		if auserid=buserid then
			errmsg = errmsg + "<tr><td colspan=8 class=forumrow>��ͬ�û����ܽ��кϲ�</td></tr>"
			founderr=true
		end if
		if founderr then
			Response.Write errmsg
		else
			'�ϲ��û�������
			dvbbs.execute("update dv_user set userpost=userpost+"&c1&",usertopic=usertopic+"&c2&",userviews=userviews+"&c3&",userwealth=userwealth+"&c4&",userep=userep+"&c5&",usercp=usercp+"&c6&",userpower=userpower+"&c7&",userisbest=userisbest+"&c8&",userdel=userdel+"&c9&" where userid="&buserid)
			'������������
			for i=0 to ubound(allposttable)
				dvbbs.execute("update "&allposttable(i)&" set postuserid="&buserid&",username='"&replace(request("buser"),"'","''")&"' where postuserid="&auserid)
			next
			dvbbs.execute("update dv_topic set postuserid="&buserid&",postusername='"&replace(request("buser"),"'","''")&"' where postuserid="&auserid)
			'���¶�������
			Dvbbs.Execute("update dv_message set incept='"&replace(request("buser"),"'","''")&"' where incept='"&replace(request("auser"),"'","''")&"'")
			Dvbbs.Execute("update dv_message set sender='"&replace(request("buser"),"'","''")&"' where sender='"&replace(request("auser"),"'","''")&"'")
			Dvbbs.Execute("update dv_friend set F_username='"&replace(request("buser"),"'","''")&"' where F_username='"&replace(request("auser"),"'","''")&"'") 
			Dvbbs.Execute("update dv_bookmark set username='"&replace(request("buser"),"'","''")&"' where username='"&replace(request("auser"),"'","''")&"'") 

			Dvbbs.Execute("update dv_besttopic set PostUserID="&buserid&",postusername='"&replace(request("buser"),"'","''")&"' where PostUserID="&auserid)
			'�����û��ϴ���
			Dvbbs.Execute("update dv_upfile set F_UserID="&buserid&",F_Username='"&replace(request("buser"),"'","''")&"' where F_UserID="&auserid)
			response.write "<tr><td colspan=8 class=forumrow>�ϲ��û����ݳɹ���</td></tr>"
		end if
	else
%>
<form action="?action=uniteuser" method=post>
<tr>
<th align=left colspan=7 height=23>�ϲ��û�</th>
</tr>
<tr>
<td width=20% class=forumrow>ע������</td>
<td width=80% class=forumrow colspan=5>���ϲ��û�����̳�е��������ӣ����������������š��ϴ����ղص����Ͻ��ϲ�����ָ�����û���</td>
</tr>
<tr>
<td width=20% class=forumrow>ѡ��</td>
<td width=80% class=forumrow colspan=5>���û� <input size=25 name="auser" type=text> ���Ϻϲ��� <input size=25 name="buser" type=text> �û� <input type=submit name=submit value="�ύ"></td>
</tr>
</form>
<%
	end if
end if
function checkreal(v)
	dim w
	if not isnull(v) then
	w=replace(v,"|||","����")
	checkreal=w
	end if
end function

%>
</table>
<p></p>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>
<% footer()%>
<%
Sub Fixuser()
	Dim Userid
	Userid = Request("Userid")
	If Not IsNumeric(Userid) Then
	Errmsg = ErrMsg + "<BR><li>��������!"
		Dvbbs_Error()
		Exit Sub
	End If
	Userid = CLng(Userid)
	Dim Rs, Username, UserArticle, UserIsBest
	UserArticle = 0
	Set Rs = Dvbbs.Execute("SELECT Username FROM [Dv_User] WHERE Userid = " & Userid & "")
	If Rs.Eof Or Rs.Bof Then
		Errmsg = ErrMsg + "<BR><li>�Ҳ������û�����ɾ�û���Ҫ������ԭ��������ע��ſ����޸�����!"
		Dvbbs_Error()
		Exit Sub
	Else
		Username = Rs(0)
		Rs.Close:Set Rs = Nothing
		'�޸������
		Dvbbs.Execute ("Update Dv_Topic Set PostUserID = " & Userid & " WHERE PostUserName = '" & Username & "'")
		'�޸��������ݱ�
		For i = 0 To Ubound(AllPostTable)
			Dvbbs.Execute ("Update " & AllPostTable(i) & " Set Postuserid = " & Userid & " WHERE UserName = '" & Username & "'")
			'�����û�����
			Set Rs = Dvbbs.Execute("SELECT COUNT(*) FROM " & AllPostTable(i) & " WHERE Postuserid = " & Userid & "")
			UserArticle = UserArticle + Rs(0)
			Rs.Close:Set Rs = Nothing
		Next
		'�޸�����
		Dvbbs.Execute ("UPDATE Dv_BestTopic Set PostUserID = " & Userid & " WHERE PostUserName = '" & Username & "'")
		Set Rs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_BestTopic WHERE Postuserid = " & Userid &"")
		UserIsBest = Rs(0)
		Rs.Close:Set Rs = Nothing
		'�޸��ϴ��ļ��б�
		Dvbbs.Execute ("UPDATE DV_Upfile SET F_UserID = " & Userid & " WHERE F_Username = '" & Username & "'")
		'���·�����
		Dvbbs.Execute ("UPDATE [Dv_User] SET UserPost = " & UserArticle & ", UserIsBest = " & UserIsBest & " WHERE Userid = " & Userid & "")
	End If
	Set Rs = Nothing
	Dv_Suc("�û�<b>" & Username & "</b>�����޸��ɹ���")
	Footer()
	Response.End
End Sub
%>