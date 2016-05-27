<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<%
dim announceid,SQl,Rs
dim rootid
dim topic
dim url
Dvbbs.LoadTemplates("postjob")
Dvbbs.stats=template.Strings(7)
Dvbbs.nav()
If Dvbbs.userid=0 Then
	Dvbbs.AddErrCode(6)
End If
If request("id")="" Then
	Dvbbs.AddErrCode(43)
ElseIf Not Isnumeric(request("id")) Then 
	Dvbbs.AddErrCode(30)
Else
	AnnounceID=Clng(request("id"))
End If
Dvbbs.ShowErr()
url="dispbbs.asp?"
url=url & "boardid="&Dvbbs.boardid&"&id="&announceid
Call chkurl()
Dvbbs.ShowErr()
Call favadd()
Dvbbs.ShowErr()
Dvbbs.head_var 1,Dvbbs.Board_Data(4,0),"",""
Dvbbs.Dvbbs_suc("<li>"&template.Strings(8))
Dvbbs.activeonline()
Dvbbs.footer()

Sub chkurl()
	sql="select title from Dv_topic where topicid="&announceid
	set rs=Dvbbs.execute(sql)
	if rs.eof and rs.bof then
		Dvbbs.AddErrCode(48)
	Else
		topic=Dvbbs.HtmlEnCode(rs(0))
	End If
	Rs.Close:set rs=nothing
End sub
sub favadd()
	sql="select * from Dv_bookmark where username='"&Dvbbs.membername&"' and url='"&trim(url)&"'"
	set rs=server.createobject("adodb.recordset")
	If Not IsObject(Conn) Then ConnectionDatabase
	rs.open sql,conn,1,3
	if not (rs.eof and rs.bof) then
		Dvbbs.AddErrCode(53)
	Else
		rs.addnew
		rs("username")=Dvbbs.membername
		rs("topic")=Dvbbs.checkStr(trim(topic))
		rs("url")=Dvbbs.checkStr(trim(url))
		rs("addtime")=Now()
		rs.update
	end if
	rs.close:set rs=nothing
end sub
%>