<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dim voteid
Dim title,votevalue,votevaluestr,voteoption
Dim TempArray,TempStr,TempStr1,TempStr2,TempStr3
Dvbbs.Loadtemplates("dispbbs")
Dvbbs.Stats=template.Strings(12)
Dvbbs.head
If request("id")="" then
	Dvbbs.AddErrCode(30)
ElseIf Not IsNumeric(request("id")) then
	Dvbbs.AddErrCode(30)
Else
	VoteID=request("id")
End If
Dvbbs.ShowErr

TempArray = Split(template.html(15),"||")
TempStr = TempArray(0)

Dim Rs,i

Set Rs=Dvbbs.Execute("select vote from dv_vote where voteid="&voteid)
If Not (rs.eof And rs.bof) Then
	votevalue=split(rs(0),"|")
Else
	Dvbbs.AddErrCode(30)
End If
Dvbbs.ShowErr

Set Rs=Dvbbs.Execute("select title from dv_topic where pollid="&voteid)
If Not (Rs.EOF And rs.bof) Then
	title=Dvbbs.HtmlEncode(rs(0))
End If

TempStr = Replace(TempStr,"{$title}",title)

Set Rs=Dvbbs.Execute("select v.*,u.username from dv_voteuser v inner join [dv_user] u on v.userid=u.userid where voteid="&voteid)
If Rs.Eof And Rs.Bof Then
	TempStr = Replace(TempStr,"{$voteinfo}",TempArray(1))
Else
	TempStr1 = TempArray(2)
	Do While Not Rs.EOF
		TempStr2 = TempStr1
		TempStr2 = Replace(TempStr2,"{$userid}",Rs("UserID"))
		TempStr2 = Replace(TempStr2,"{$username}",Rs("UserName"))
		voteoption = Split(rs("voteoption"),",")
		For i = 0 To Ubound(voteoption)
			If IsNumeric(voteoption(i)) Then
				If i<>0 Then votevaluestr = votevaluestr & "<BR>"
				votevaluestr = votevaluestr & votevalue(voteoption(i))
			End If
		Next
		TempStr2 = Replace(TempStr2,"{$uservote}",votevaluestr)
		votevaluestr = ""

		TempStr3 = TempStr3 & TempStr2
	Rs.MoveNext
	Loop
	TempStr = Replace(TempStr,"{$voteinfo}",TempStr3)
End If
Rs.Close
Set Rs =Nothing
Response.Write TempStr
%>