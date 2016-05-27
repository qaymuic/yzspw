<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<%
	Dim times,Rs
	Dvbbs.stats="Ìø×ªÖ÷Ìâ"
	If request("sid")="" then
		Dvbbs.AddErrCode(43)
	Elseif not Isnumeric(request("sid")) then
		Dvbbs.AddErrCode(35)
	Else
		times=request("sid")
	End If
	Dvbbs.ShowErr()
If request("action")="next" Then
	set rs=Dvbbs.execute("select top 1 topicid from Dv_topic where boardid="&Dvbbs.boardid&" and topicid>"&times&" and not locktopic=2 order by Dateandtime")
	If rs.eof and rs.bof Then
		Dvbbs.AddErrCode(44)
	Else
		response.redirect "dispbbs.asp?boardid="&Dvbbs.boardid&"&ID="&rs(0)
	End If
Else
	Set rs=Dvbbs.execute("select top 1 topicid from Dv_topic where boardid="&Dvbbs.boardid&" and  topicid<"&times&" and not locktopic=2 order by Dateandtime desc")
	If rs.eof and rs.bof Then
		Dvbbs.AddErrCode(45)
		Set rs=Nothing
	Else	
		response.redirect "dispbbs.asp?boardid="&Dvbbs.boardid&"&ID="&rs(0)
	End If
End If
If Dvbbs.ErrCodes<>"" Then Dvbbs.ShowErr
%>