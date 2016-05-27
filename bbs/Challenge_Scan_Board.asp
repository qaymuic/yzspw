<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Response.Clear
dim rechallengeWord,retokerWord
dim remobile
dim issuc
dim trs
dim myboarduser,myboardtype
dim errcode
dim rs,i
issuc=false

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(8)=1) Then
	Response.Write "本论坛没有开启VIP收费论坛功能。"
	Response.End
End If

rechallengeWord=trim(Dvbbs.CheckStr(request("challengWord")))
retokerWord=trim(request("tokenWord"))
remobile=Dvbbs.CheckStr(trim(request("mobile")))

if md5(rechallengeWord & ":" & Dvbbs.CacheData(21,0),32)=retokerWord then
	issuc=false
	set rs=Dvbbs.Execute("select username from [dv_user] where UserMobile='"&remobile&"'")
	'set rs=Dvbbs.Execute("select * from DV_ChanOrders where O_isApply=1 and O_type=2 and O_Paycode='"&apaycode(i)&"'")
	if rs.eof and rs.bof then
		issuc=false
	else
		set trs=Dvbbs.Execute("select * from dv_board where boardid="&dvbbs.boardid)
		if not (trs.eof and trs.bof) then
			myboarduser=trs("boarduser")
			myboardtype=trs("boardtype")
			if myboarduser="" then
				myboarduser=rs("Username")
			else
				if instr("," & lcase(myboarduser) & ",","," & lcase(rs("Username")) & ",")=0 then
					myboarduser=myboarduser & "," & rs("Username")
				end if
			end if
			Dvbbs.Execute("update dv_board set boarduser='"&myboarduser&"' where boardid="&dvbbs.boardid)
			Dvbbs.ReloadBoardInfo(Dvbbs.BoardID)
			issuc=true
			dvbbs.membername=rs("Username")
		end if
		set trs=nothing
	end if
	set rs=nothing
	if issuc then
		Response.Write "100"
		Dvbbs.Execute("insert into DV_ChanOrders (O_type,O_mobile,O_Username,O_isApply,O_issuc,O_PayMoney,O_BoardID) values (2,'"&remobile&"','"&dvbbs.membername&"',1,1,"&dvbbs.Board_Setting(20)&","&dvbbs.boardid&")")
	else
		Response.Write "201"
	end if
else
Response.Write "local tokenword:" & md5(rechallengeWord & ":" & Dvbbs.CacheData(21,0),32)
Response.Write ",ray tokenword:" & retokerword
Response.Write ",ray cword:" & rechallengeWord
	Response.Write "201"
end if
%>