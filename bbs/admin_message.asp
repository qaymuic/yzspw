<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
Head()
Server.ScriptTimeout=9999999
Dim admin_flag
Dim Numc
admin_flag = ",6,"
If Not Dvbbs.Master or instr(","&session("flag")&",",admin_flag)=0 then
	Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	dvbbs_error()
Else
	Dim Body
	If Request("action") = "add" Then
		Call Savemsg()
	Elseif Request("action")="del" Then
		Call Del()
	Elseif Request("action")="delall" Then
		Call Delall()
	Elseif Request("action")="delchk" Then
		Call Delchk()
	Else
		Call Sendmsg()
	End if
%>
<p align=center><%=body%></p></font>
<%
	Footer()
End If

Sub Savemsg()
	Dim Sendtime,sender,userlist,message,isshow
	isshow=Request("isshow")
	message=Request("message")
	message=Dvbbs.checkStr(message)
	If Len(message)>255 Then
		Response.Write "消息内容不能多于255字节"
		Exit Sub			
	End If 
	sendtime=Now()
	sender=Dvbbs.Forum_info(0)
	Select case request("stype")
	case 1
		Sql = "SELECT Count(*) FROM [dv_online] where userid>0"
		Set Rs = Dvbbs.execute(Sql)
		Numc = Rs(0)
		sql="select username from dv_online where userid>0"
	Case 2
		Sql = "SELECT Count(*) FROM [dv_user] where usergroupid=8"
		Set Rs = Dvbbs.execute(Sql)
		Numc = Rs(0)
		sql = "select username from [dv_user] where usergroupid=8 order by userid desc"
	Case 3
		Sql = "SELECT Count(*) FROM [dv_user] where usergroupid=3"
		Set Rs = Dvbbs.execute(Sql)
		Numc = Rs(0)
		sql = "select username from [dv_user] where usergroupid=3 order by userid desc"
	Case 4
   		Sql = "SELECT Count(*) FROM [dv_user] where usergroupid=1"
		Set Rs = Dvbbs.execute(Sql)
		Numc = Rs(0)
   		sql = "select username from [dv_user] where usergroupid=1 order by userid desc"
	Case 5
   		Sql = "SELECT Count(*) FROM [dv_user] where usergroupid<4"
		Set Rs = Dvbbs.execute(Sql)
		Numc = Rs(0)
   		sql = "select username from [dv_user] where usergroupid<4 order by userid desc"
	Case 6
		Sql = "SELECT Count(*) FROM [Dv_user]"
		Set Rs = Dvbbs.execute(Sql)
		Numc = Rs(0)
		Rs.Close
	    Sql = "SELECT Username FROM [Dv_user] ORDER BY Userid DESC"
	Case 7
		Sql = "SELECT COUNT(*) FROM [Dv_User] WHERE UserGroupID = 2"
		Set Rs = Dvbbs.Execute(Sql)
		Numc = Rs(0)
		sql = "SELECT UserName FROM [Dv_User] WHERE UserGroupID = 2 ORDER BY UserID DESC"
	Case Else
		REM 加入自定义用户组群发短信功能 2004-5-19 Dv.Yz
		Sql = "SELECT COUNT(*) FROM [Dv_User] WHERE Usergroupid = " & Cint(Request("stype"))
		Set Rs = Dvbbs.Execute(Sql)
		Numc = Rs(0)
		Sql = "SELECT Username FROM [Dv_User] WHERE Usergroupid = " & Cint(Request("stype")) & " ORDER BY Userid DESC"
	End Select
%>
<br><table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<tr><td colspan=2 class=forumrow>
下面开始发送短消息，预计本次发送<%=Numc%>个用户。
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=ffffff height=9><img src="skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span></td></tr>
</table>
<%
Response.Flush
	Set rs=Dvbbs.Execute(SQL)
	userlist=Rs.GetRows(-1)
	Set rs=Nothing
		Response.Write "<script>img2.width=" & Fix((i/Numc) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""正在发送，..."";" & VbCrLf
		Response.Write "img2.title=""发送短信给...."";</script>" & VbCrLf
		Response.Flush
		For i=0 to UBound(userlist,2)
			userlist(0,i)=Dvbbs.checkStr(userlist(0,i))
			If Response.IsClientConnected Then
				If isshow="1" Then
					Response.Write "<script>img2.width=" & Fix((i/Numc) * 400) & ";" & VbCrLf
					Response.Write "txt2.innerHTML=""" & FormatNumber(i/Numc*100,4,-1) & "%，发送短信给" & userlist(0,i) & "成功！"";" & VbCrLf
					Response.Write "img2.title=""发送短信给" & userlist(0,i)  & "成功！"";</script>" & VbCrLf
					Response.Flush
				End If
				Sql = "INSERT into dv_message(incept, sender, title, content, sendtime, flag, issend) values('"&userlist(0,i) &"', '"&sender&"', '"&TRim(Request("title"))&"', '"&Trim(message)&"', "&SqlNowString&",0,1)"
				Dvbbs.Execute(Sql)
				Update_user_msg(userlist(0,i))
				userlist(0,i)=""
			End If 
		Next 
	Response.Write "<script>img2.width=400;" & VbCrLf
	Response.Write "txt2.innerHTML=""100%，发送完成"";" & VbCrLf
	Response.Write "img2.title=""发送短信给...."";</script>" & VbCrLf
	Response.Flush
	body=body+"<br>"+"操作成功！请继续别的操作。"
end sub

sub sendmsg()
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0"  align=center class="tableBorder">
                <tr> 
                  <th colspan="2" height=24>论坛短信管理
                  </th>
                </tr>
            <form action="admin_message.asp?action=del" method=post>
                <tr> 
                  <td colspan="2" class=ForumrowHighLight>
                      批量删除某用户短消息（主要用于删除系统批量信息：动网小精灵）：<br><input type="text" name="username" size="20">
			<input type="submit" name="Submit" value="提 交">
                  </td>
                </tr>
            </form>
			<form action="admin_message.asp?action=delall" method=post>
                <tr> 
                  <td colspan="2" class=ForumrowHighLight>
                      批量删除用户指定日期内短消息（默认为删除已读信息）：<br>
					  <select name="delDate" size=1>
						<option value=7>一个星期前</option>
						<option value=30>一个月前</option>
						<option value=60>两个月前</option>
						<option value=180>半年前</option>
						<option value="all">所有信息</option>
					  </select>
					  &nbsp;<input type="checkbox" name="isread" value="yes">包括未读信息
			<input type="submit" name="Submit" value="提 交">
                  </td>
                </tr>
            </form>
			<form action="admin_message.asp?action=delchk" method=post>
                <tr> 
                  <td colspan="2" class=ForumrowHighLight>
				  批量删除含有某关键字短信（注意：本操作将删除所有已读和未读信息）：<br>
				  关键字：<input type="text" name="keyword" size=30>&nbsp;在
					  <select name="selaction" size=1>
						<option value=1>标题中</option>
						<option value=2>内容中</option>
					  </select>
					  &nbsp;<input type="submit" name="Submit" value="提 交">
                  </td>
                </tr>
            </form>
                <tr> 
                  <th colspan="2" height=24>论坛短信广播
                  </th>
                </tr>
            <form action="admin_message.asp?action=add" method=post>
                <tr> 
                  <td width="22%" class=Forumrow>消息标题</td>
                  <td width="78%" class=Forumrow> 
                    <input type="text" name="title" size="70">
                  </td>
                </tr>
                <tr> 
                  <td width="22%" class=Forumrow>接收方选择</td>
                  <td width="78%" class=Forumrow> 
                    <select name=stype size=1>
					<option value="1">所有在线用户</option>
					<option value="2">所有贵宾</option>
					<option value="3">所有版主</option>
					<option value="4">所有管理员</option>
					<option value="5">版主/超版/管理员</option>
					<option value="6">所有用户</option>
					<option value="7">所有超版</option>
<%
	Sql = "SELECT UserGroupID, Title From Dv_UserGroups WHERE UserGroupID > 8 AND ParentGID = 0 ORDER BY UserGroupID"
	Set Rs = Dvbbs.Execute(Sql)
	If Not (Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		For i = 0 To Ubound(Sql,2)
%>
					<option value="<%=Cint(Sql(0,i))%>"><%=Dvbbs.HtmlEnCode(Sql(1,i))%></option>
<%
		Next
	End If
%>
					</select>
                  </td>
                </tr>
                <tr> 
                  <td width="22%" height="20" valign="top" class=Forumrow>
                    <p>消息内容</p>
                    <p>(<font color="red">HTML代码支持</font>)</p>
                  </td>
                  <td width="78%" height="20" class=Forumrow> 
                    <textarea name="message" cols="80" rows="10"></textarea>
                    <br><input type="radio" name="isshow" value="1" checked>显示发送过程 <input type="radio" name="isshow" value="0" > 不显示发送过程（速度较快）
                  </td>
                </tr>
                <tr> 
                  <td width="22%" height="23" valign="top" align="center" class=Forumrow> 
                    <div align="left"> </div>
                  </td>
                  <td width="78%" height="23" class=Forumrow> 
                    <div align="center"> 
                      <input type="submit" name="Submit" value="发送消息">
                      <input type="reset" name="Submit2" value="重新填写">
                    </div>
                  </td>
                </tr>
            </form>
              </table>
<%
end sub

Sub Del()
	If Request("username") = "" Then
		Body = Body + "<br>" + "请输入要批量删除的用户名。"
		Exit Sub
	End If
	Sql = "DELETE FROM Dv_Message WHERE Sender = '" & Request("username") & "'"
	Dvbbs.Execute(Sql)
	Body = Body + "<br>" + "操作成功！请继续别的操作。"
End Sub

Sub Delall()
	REM 改数组循环避免删除论坛短信超时 2004-5-11 Dvbbs.YangZheng
	Dim Selflag, Summid
	If Request("isread") = "yes" Then
		Selflag = " ORDER BY Id"
	Else
		Selflag = " AND Flag = 1 ORDER BY Id"
	End If
	Select Case Request("delDate")
	Case "all"
		Sql = "SELECT Id FROM Dv_Message WHERE Id > 0 " & Selflag
	Case 7
		If IsSqlDataBase = 1 Then
			Sql = "SELECT Id From Dv_Message WHERE DATEDIFF(d, Sendtime, " & SqlNowString & ") > 7 " & Selflag
		Else
			Sql = "SELECT Id FROM Dv_Message WHERE DATEDIFF('d', Sendtime, " & SqlNowString & ") > 7 " & Selflag
		End If
	Case 30
		If IsSqlDataBase = 1 Then
			Sql = "SELECT Id From Dv_Message WHERE DATEDIFF(d, Sendtime, " & SqlNowString & ") > 30 " & Selflag
		Else
			Sql = "SELECT Id FROM Dv_Message WHERE DATEDIFF('d', Sendtime, " & SqlNowString & ") > 30 " & Selflag
		End If
	Case 60
		If IsSqlDataBase = 1 Then
			Sql = "SELECT Id From Dv_Message WHERE DATEDIFF(d, Sendtime, " & SqlNowString & ") > 60 " & Selflag
		Else
			Sql = "SELECT Id FROM Dv_Message WHERE DATEDIFF('d', Sendtime, " & SqlNowString & ") > 60 " & Selflag
		End If
	Case 180
		If IsSqlDataBase = 1 Then
			Sql = "SELECT Id From Dv_Message WHERE DATEDIFF(d, Sendtime, " & SqlNowString & ") > 180 " & Selflag
		Else
			Sql = "SELECT Id FROM Dv_Message WHERE DATEDIFF('d', Sendtime, " & SqlNowString & ") > 180 " & Selflag
		End If
	End Select
	Set Rs = Dvbbs.Execute(Sql)
	Summid = 0
	If Not (Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		For i = 0 To Ubound(Sql,2)
			Dvbbs.Execute("DELETE FROM Dv_Message Where Id = " & Sql(0,i))
			Summid = Summid + 1
		Next
	End If
	Body = Body + "<br>" + "操作删除" & Summid & "条论坛短信成功！请继续别的操作。"
End Sub

sub delchk()
	if request.form("keyword")="" then
	body="请输入关键字！"
	exit sub
	end if
	if request.form("selaction")=1 then
	Dvbbs.Execute("delete from dv_message where title like '%"&replace(request.form("keyword"),"'","")&"%'")
	body="操作成功！请继续别的操作。"
	elseif request.form("selaction")=2 then
	Dvbbs.Execute("delete from dv_message where content like '%"&replace(request.form("keyword"),"'","")&"%'")
	body="操作成功！请继续别的操作。"
	else
	body="未指定相关参数！"
	exit sub
	end if
End Sub
Function inceptid(stype,iusername)
	Dim ars
	set ars=Dvbbs.Execute("Select top 1 id,sender from dv_Message Where flag=0 and issend=1 and delR=0 And incept ='"& iusername &"'")
	if stype=1 then
	inceptid=ars(0)
	else
	inceptid=ars(1)
	end if
	set ars=nothing
End Function
Function update_user_msg(username)
	Dim msginfo
	If newincept(username)>0 Then
		msginfo=newincept(username) & "||" & inceptid(1,username) & "||" & inceptid(2,username)
	Else
		msginfo="0||0||null"
	End If
	Dvbbs.Execute("update [dv_user] set UserMsg='"&dvbbs.CheckStr(msginfo)&"' where username='"&dvbbs.CheckStr(username)&"'")
End Function
'统计留言
Function newincept(iusername)
	Dim rs
	Rs=Dvbbs.Execute("Select Count(id) from dv_Message Where flag=0 and issend=1 and delR=0 And incept='"& iusername &"'")
	newincept=Rs(0)
	Set Rs=Nothing
	If IsNull(newincept) Then newincept=0
End Function
%>