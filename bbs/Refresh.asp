<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> Update_Userid </TITLE>
<meta NAME=GENERATOR CHARSET=GB2312>
<meta http-equiv="refresh" content="20">
</HEAD>
<%
'Response.Write "<iframe width=0 height=0 src=""Refresh.asp"" name=""Refresh""></iframe>"
IF Dvbbs.userid<>0 Then UPDATE_User_Msg(Dvbbs.membername)

'更新用户短信通知信息（新短信条数||新短讯ID||发信人名）
Sub UPDATE_User_Msg(username)
	Dim msginfo,i,UP_UserInfo
	If newincept(username)>0 Then
		msginfo=newincept(username) & "||" & inceptid(1,username) & "||" & inceptid(2,username)
	Else
		msginfo="0||0||null"
	End If
	Dvbbs.execute("UPDATE [Dv_User] Set UserMsg='"&Dvbbs.CheckStr(msginfo)&"' WHERE username='"&Dvbbs.CheckStr(username)&"'")
	If username=Dvbbs.MemberName Then 
		UP_UserInfo=Session(Dvbbs.CacheName & "UserID")
		UP_UserInfo(30)=msginfo
		Session(Dvbbs.CacheName & "UserID")=UP_UserInfo
	End If
End Sub

'统计留言
Function newincept(iusername)
Dim Rs
Rs=Dvbbs.execute("SELECT Count(id) FROM Dv_Message WHERE flag=0 And issend=1 And DelR=0 And incept='"& iusername &"'")
    newincept=Rs(0)
	Set Rs=nothing
	If isnull(newincept) Then newincept=0
End Function

Function inceptid(stype,iusername)
	Set Rs=Dvbbs.execute("SELECT top 1 id,sender FROM Dv_Message WHERE flag=0 And issend=1 And DelR=0 And incept='"& iusername &"'")
	If not rs.eof Then
		If stype=1 Then
			inceptid=Rs(0)
		Else
			inceptid=Rs(1)
		End If
	Else
		If stype=1 Then
			inceptid=0
		Else
			inceptid="null"
		End If
	End If
	Set Rs=nothing
End Function
%>
<BODY>
</BODY>
</HTML>
