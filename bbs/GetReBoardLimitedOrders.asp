<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Response.Clear
dim rechallengeWord,retokerWord,repayid,paycode
dim challengeWord_key,rechallengeWord_key
dim trs,boarduser
dim rs,i

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(8)=1) Then
	Response.Write "����̳û�п���VIP�շ���̳���ܡ�"
	Response.End
End If

repayid=trim(Dvbbs.CheckStr(request("postdata")))
rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
retokerWord=trim(request("tokenWord"))

challengeWord_key=session("challengeWord_key")
if challengeWord_key=retokerWord then
	'type=1�������⣬type=2������̳
	paycode=split(repayid,":")
	for i=0 to ubound(paycode)
		set rs=Dvbbs.Execute("select * from DV_ChanOrders where O_Paycode='"&trim(paycode(i))&"'")
		if not (rs.eof and rs.bof) then
			Dvbbs.Execute("update DV_ChanOrders set o_issuc=1 where o_id="&rs("o_id"))
			set trs=Dvbbs.Execute("select * from dv_board where boardid="&rs("o_boardid"))
			if not (trs.eof and trs.bof) then
				boarduser=rs("boarduser")
				if isnull(boarduser) or boarduser="" then
				boarduser=rs("o_username")
				else
				boarduser=boarduser & "," & rs("o_username")
				end if
				Dvbbs.Execute("update dv_board set boarduser='"&boarduser&"' where boardid="&rs("o_boardid"))
			end if
			set trs=nothing
		end if
		set rs=nothing
	next
	'���سɹ���Ϣ
else
	'����ʧ����Ϣ
end if
%>