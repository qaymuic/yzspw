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
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
		Response.Write "��Ϣ���ݲ��ܶ���255�ֽ�"
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
		REM �����Զ����û���Ⱥ�����Ź��� 2004-5-19 Dv.Yz
		Sql = "SELECT COUNT(*) FROM [Dv_User] WHERE Usergroupid = " & Cint(Request("stype"))
		Set Rs = Dvbbs.Execute(Sql)
		Numc = Rs(0)
		Sql = "SELECT Username FROM [Dv_User] WHERE Usergroupid = " & Cint(Request("stype")) & " ORDER BY Userid DESC"
	End Select
%>
<br><table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<tr><td colspan=2 class=forumrow>
���濪ʼ���Ͷ���Ϣ��Ԥ�Ʊ��η���<%=Numc%>���û���
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
		Response.Write "txt2.innerHTML=""���ڷ��ͣ�..."";" & VbCrLf
		Response.Write "img2.title=""���Ͷ��Ÿ�...."";</script>" & VbCrLf
		Response.Flush
		For i=0 to UBound(userlist,2)
			userlist(0,i)=Dvbbs.checkStr(userlist(0,i))
			If Response.IsClientConnected Then
				If isshow="1" Then
					Response.Write "<script>img2.width=" & Fix((i/Numc) * 400) & ";" & VbCrLf
					Response.Write "txt2.innerHTML=""" & FormatNumber(i/Numc*100,4,-1) & "%�����Ͷ��Ÿ�" & userlist(0,i) & "�ɹ���"";" & VbCrLf
					Response.Write "img2.title=""���Ͷ��Ÿ�" & userlist(0,i)  & "�ɹ���"";</script>" & VbCrLf
					Response.Flush
				End If
				Sql = "INSERT into dv_message(incept, sender, title, content, sendtime, flag, issend) values('"&userlist(0,i) &"', '"&sender&"', '"&TRim(Request("title"))&"', '"&Trim(message)&"', "&SqlNowString&",0,1)"
				Dvbbs.Execute(Sql)
				Update_user_msg(userlist(0,i))
				userlist(0,i)=""
			End If 
		Next 
	Response.Write "<script>img2.width=400;" & VbCrLf
	Response.Write "txt2.innerHTML=""100%���������"";" & VbCrLf
	Response.Write "img2.title=""���Ͷ��Ÿ�...."";</script>" & VbCrLf
	Response.Flush
	body=body+"<br>"+"�����ɹ����������Ĳ�����"
end sub

sub sendmsg()
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0"  align=center class="tableBorder">
                <tr> 
                  <th colspan="2" height=24>��̳���Ź���
                  </th>
                </tr>
            <form action="admin_message.asp?action=del" method=post>
                <tr> 
                  <td colspan="2" class=ForumrowHighLight>
                      ����ɾ��ĳ�û�����Ϣ����Ҫ����ɾ��ϵͳ������Ϣ������С���飩��<br><input type="text" name="username" size="20">
			<input type="submit" name="Submit" value="�� ��">
                  </td>
                </tr>
            </form>
			<form action="admin_message.asp?action=delall" method=post>
                <tr> 
                  <td colspan="2" class=ForumrowHighLight>
                      ����ɾ���û�ָ�������ڶ���Ϣ��Ĭ��Ϊɾ���Ѷ���Ϣ����<br>
					  <select name="delDate" size=1>
						<option value=7>һ������ǰ</option>
						<option value=30>һ����ǰ</option>
						<option value=60>������ǰ</option>
						<option value=180>����ǰ</option>
						<option value="all">������Ϣ</option>
					  </select>
					  &nbsp;<input type="checkbox" name="isread" value="yes">����δ����Ϣ
			<input type="submit" name="Submit" value="�� ��">
                  </td>
                </tr>
            </form>
			<form action="admin_message.asp?action=delchk" method=post>
                <tr> 
                  <td colspan="2" class=ForumrowHighLight>
				  ����ɾ������ĳ�ؼ��ֶ��ţ�ע�⣺��������ɾ�������Ѷ���δ����Ϣ����<br>
				  �ؼ��֣�<input type="text" name="keyword" size=30>&nbsp;��
					  <select name="selaction" size=1>
						<option value=1>������</option>
						<option value=2>������</option>
					  </select>
					  &nbsp;<input type="submit" name="Submit" value="�� ��">
                  </td>
                </tr>
            </form>
                <tr> 
                  <th colspan="2" height=24>��̳���Ź㲥
                  </th>
                </tr>
            <form action="admin_message.asp?action=add" method=post>
                <tr> 
                  <td width="22%" class=Forumrow>��Ϣ����</td>
                  <td width="78%" class=Forumrow> 
                    <input type="text" name="title" size="70">
                  </td>
                </tr>
                <tr> 
                  <td width="22%" class=Forumrow>���շ�ѡ��</td>
                  <td width="78%" class=Forumrow> 
                    <select name=stype size=1>
					<option value="1">���������û�</option>
					<option value="2">���й��</option>
					<option value="3">���а���</option>
					<option value="4">���й���Ա</option>
					<option value="5">����/����/����Ա</option>
					<option value="6">�����û�</option>
					<option value="7">���г���</option>
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
                    <p>��Ϣ����</p>
                    <p>(<font color="red">HTML����֧��</font>)</p>
                  </td>
                  <td width="78%" height="20" class=Forumrow> 
                    <textarea name="message" cols="80" rows="10"></textarea>
                    <br><input type="radio" name="isshow" value="1" checked>��ʾ���͹��� <input type="radio" name="isshow" value="0" > ����ʾ���͹��̣��ٶȽϿ죩
                  </td>
                </tr>
                <tr> 
                  <td width="22%" height="23" valign="top" align="center" class=Forumrow> 
                    <div align="left"> </div>
                  </td>
                  <td width="78%" height="23" class=Forumrow> 
                    <div align="center"> 
                      <input type="submit" name="Submit" value="������Ϣ">
                      <input type="reset" name="Submit2" value="������д">
                    </div>
                  </td>
                </tr>
            </form>
              </table>
<%
end sub

Sub Del()
	If Request("username") = "" Then
		Body = Body + "<br>" + "������Ҫ����ɾ�����û�����"
		Exit Sub
	End If
	Sql = "DELETE FROM Dv_Message WHERE Sender = '" & Request("username") & "'"
	Dvbbs.Execute(Sql)
	Body = Body + "<br>" + "�����ɹ����������Ĳ�����"
End Sub

Sub Delall()
	REM ������ѭ������ɾ����̳���ų�ʱ 2004-5-11 Dvbbs.YangZheng
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
	Body = Body + "<br>" + "����ɾ��" & Summid & "����̳���ųɹ����������Ĳ�����"
End Sub

sub delchk()
	if request.form("keyword")="" then
	body="������ؼ��֣�"
	exit sub
	end if
	if request.form("selaction")=1 then
	Dvbbs.Execute("delete from dv_message where title like '%"&replace(request.form("keyword"),"'","")&"%'")
	body="�����ɹ����������Ĳ�����"
	elseif request.form("selaction")=2 then
	Dvbbs.Execute("delete from dv_message where content like '%"&replace(request.form("keyword"),"'","")&"%'")
	body="�����ɹ����������Ĳ�����"
	else
	body="δָ����ز�����"
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
'ͳ������
Function newincept(iusername)
	Dim rs
	Rs=Dvbbs.Execute("Select Count(id) from dv_Message Where flag=0 and issend=1 and delR=0 And incept='"& iusername &"'")
	newincept=Rs(0)
	Set Rs=Nothing
	If IsNull(newincept) Then newincept=0
End Function
%>