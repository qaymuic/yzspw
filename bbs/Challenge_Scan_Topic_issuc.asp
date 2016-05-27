<!--#include file="Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Response.Clear
Dvbbs.Loadtemplates("")
dim i
dim retokenword,subsubject,falsubject
dim ssubsubject,sfalsubject
dim challengeWord
challengeWord=request("challengeWord")
retokenWord=request("forumChanWord")
subsubject=request("subsubject")
falsubject=request("falsubject")

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(7)=1) Then
	Response.Write "本论坛没有开启主题订阅手机短信功能。"
	Response.End
End If
'Response.Write Dvbbs.CacheData(21,0)
if md5(challengeWord & ":" & Dvbbs.CacheData(21,0),32)=trim(retokenWord) then

	if subsubject<>"" then
		ssubsubject=Split(subsubject,";")
		for i=0 to ubound(ssubsubject)
			if isnumeric(ssubsubject(i)) then
				Dvbbs.Execute("update dv_topic set LastSmsTime="&SqlNowString&" where topicid="&ssubsubject(i))
			end if
		next
	end if

	if falsubject<>"" then
		sfalsubject=split(falsubject,";")
		for i=0 to ubound(sfalsubject)
			if isnumeric(sfalsubject(i)) then
				Dvbbs.Execute("update dv_topic set IsSmsTopic=0,LastSmsTime="&SqlNowString&" where topicid="&sfalsubject(i))
			end if
		next
	end if

end if
%>