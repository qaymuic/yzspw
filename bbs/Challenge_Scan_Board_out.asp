<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Response.Clear
dim rechallengeWord,retokerWord
dim paycode,apaycode
dim issuc
dim trs
dim myboarduser,myboardtype
dim errcode
dim myboarduserlen
dim rs,i
issuc=false

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(8)=1) Then
	Response.Write "本论坛没有开启VIP收费论坛功能。"
	Response.End
End If

rechallengeWord=trim(dvbbs.CheckStr(request("challengeWord")))
retokerWord=trim(request("tokenWord"))
paycode=dvbbs.CheckStr(trim(request("postData")))

if md5(rechallengeWord & ":" & Dvbbs.CacheData(21,0),32)=retokerWord then

Response.Write "<results>"
apaycode=split(paycode,":")
for i=0 to ubound(apaycode)
	issuc=false
	set rs=Dvbbs.Execute("select * from DV_ChanOrders where O_issuc=1 and O_type=2 and O_Paycode='"&apaycode(i)&"'")
	if rs.eof and rs.bof then
		issuc=false
	else
		set trs=Dvbbs.Execute("selete * from dv_board where boardid="&rs("O_boardid"))
		if not (trs.eof and trs.bof) then
			myboarduser=trs("boarduser")
			myboardtype=trs("boardtype")
			if myboarduser<>"" then
				myboarduser="," & myboarduser & ","
				myboarduser=replace(lcase(myboarduser),"," & lcase(rs("O_Username")) & ",",",")
				'去掉附加的右边逗号
				myboarduserlen=len(myboarduser)
				myboarduser=left(myboarduser,myboarduserlen-1)
				'去掉附件的左边逗号
				myboarduserlen=len(myboarduser)
				myboarduser=right(myboarduser,myboarduserlen-1)
			end if
			Dvbbs.Execute("update dv_board set boarduser='"&myboarduser&"' where boardid="&rs("O_boardid"))
			Dvbbs.ReloadBoardInfo(rs("O_boardid"))
			Dvbbs.Execute("delete from DV_ChanOrders where O_id="&rs("o_id"))
			issuc=true
		end if
		set trs=nothing
	end if
	set rs=nothing
next
else
	issuc=false
end if

if issuc then
	errcode="100"
else
	errcode="201"
end if
Response.Write errcode
%>