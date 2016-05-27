<!-- #include file="Conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
dvbbs.stats="参与投票"
dim voteid
Dim announceid
announceid=CheckBoardInfo
If Dvbbs.IsReadonly()  And Not Dvbbs.Master Then Response.redirect "showerr.asp?action=readonly&boardid="&dvbbs.boardID&"" 
dim action
dim vote,votenum
dim postvote(200)
dim postvote1
dim j,votenum_1,votenumlen
dim vrs
dim postnum,postoption

If Dvbbs.UserID=0 then Dvbbs.AddErrCode(34)
	
if request("id")="" then
	Dvbbs.AddErrCode(35)
ElseIf Not IsNumeric(request("id")) then
	Dvbbs.AddErrCode(35)
Else
	AnnounceID=request("id")
End If
If request("voteid")="" then
	Dvbbs.AddErrCode(35)
ElseIf not IsNumeric(request("voteid")) then
	Dvbbs.AddErrCode(35)
Else
	voteID=request("voteid")
end If
	
If CInt(Dvbbs.GroupSetting(9))=0 then Dvbbs.AddErrCode(56)
Dvbbs.ShowErr

Main
Dvbbs.ShowErr
Sub main()
	Dim Rs,SQL,i
	set rs=Dvbbs.Execute("select locktopic from dv_topic where topicid="&AnnounceID)
	If Not (rs.eof and rs.bof) then
		If Rs(0)=1 Then
			Dvbbs.AddErrCode(57)
			Exit Sub
		End If
	End If
	Set rs=server.createobject("adodb.recordset")
	sql="select * from dv_vote where voteid="&voteid
	rs.open sql,conn,1,3
	If rs.eof and rs.bof Then
		Dvbbs.AddErrCode(32)
		Exit Sub
	Else
		If Not (Dvbbs.Master Or Dvbbs.SuperBoardMaster Or Dvbbs.BoardMaster) Then
		'文章
		If Clng(Rs("UArticle"))>Clng(Dvbbs.MyUserInfo(8)) Then Response.redirect "showerr.asp?ErrCodes=<li>本投票设置了用户发贴最少为 <B>"&Rs("UArticle")&"</B> 才能投票&action=OtherErr"
		'金钱
		If Clng(Rs("UWealth"))>Clng(Dvbbs.MyUserInfo(21)) Then Response.redirect "showerr.asp?ErrCodes=<li>本投票设置了用户金钱最少为 <B>"&Rs("UWealth")&"</B> 才能投票&action=OtherErr"
		'经验
		If Clng(Rs("UEP"))>Clng(Dvbbs.MyUserInfo(22)) Then Response.redirect "showerr.asp?ErrCodes=<li>本投票设置了用户积分最少为 <B>"&Rs("UEP")&"</B> 才能投票&action=OtherErr"
		'魅力
		If Clng(Rs("UCP"))>Clng(Dvbbs.MyUserInfo(23)) Then Response.redirect "showerr.asp?ErrCodes=<li>本投票设置了用户魅力最少为 <B>"&Rs("UCP")&"</B> 才能投票&action=OtherErr"
		'威望
		If Clng(Rs("UPower"))>Clng(Dvbbs.MyUserInfo(24)) Then Response.redirect "showerr.asp?ErrCodes=<li>本投票设置了用户威望最少为 <B>"&Rs("UPower")&"</B> 才能投票&action=OtherErr"
		End If
		Set vrs=Dvbbs.Execute("select userid from dv_voteuser where voteid="&voteID&" and userid="&Dvbbs.userid)
		If Not(vrs.eof and vrs.bof) Then
			Dvbbs.AddErrCode(58)
			Exit Sub
		Else 
			votenum=split(rs("votenum"),"|")
			If Rs("votetype")=1 Then
				For i = 0 to UBound(votenum)
					postvote(i)=request("postvote_"&i&"")
				Next
			End If 
			For j = 0 to UBound(votenum)
				If rs("votetype")=0 Then
					if cint(request("postvote"))=j Then
						votenum(j)=votenum(j)+1
						postoption=j
					End If
					votenum_1=""&votenum_1&""&votenum(j)&"|"
					postnum=1
				Else
					If postvote(j)<>"" Then
						If cint(postvote(j))=j Then
							votenum(j)=votenum(j)+1
							postnum=postnum+1
							postoption=postoption & j & ","
						End If
					End If
					votenum_1=""&votenum_1&""&votenum(j)&"|"
				End If
			Next
			If postnum="" or isnull(postnum) then
				Dvbbs.AddErrCode(59)
				Exit Sub
			End If
			votenumlen=len(votenum_1)
			votenum_1=left(votenum_1,votenumlen-1)
			rs("votenum")=votenum_1
			rs("voters")=rs("voters")+1
			rs.update
			Dvbbs.Execute("update dv_Topic set VoteTotal=voteTotal+"&postnum&" where topicid="&Announceid)
			Dvbbs.Execute("insert into dv_voteuser (voteid,userid,voteoption) values ("&voteid&","&Dvbbs.userid&",'"&postoption&"')")
		End If
	End If 
	
	Rs.Close
	Set Rs=Nothing
	If Dvbbs.Board_Setting(53)<>"0" Then
		SQL="update dv_topic set LastPostTime="&SqlNowString&" where Topicid="&announceid&" and istop=0"
		Dvbbs.Execute(SQL)
	End If
	response.redirect Request.ServerVariables("HTTP_REFERER")
end Sub
%>