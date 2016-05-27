<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Response.Clear
Dvbbs.Loadtemplates("")
dim rs
dim raychanword,tokenword
dim LastPost_ray,LastID,res,trs
dim topicid,title,body,nikename,posttime
rayChanWord=request("rayChanWord")
tokenword=md5(request("rayChanWord") & ":" & Dvbbs.CacheData(21,0),32)

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(7)=1) Then
	Response.Write "本论坛没有开启主题订阅手机短信功能。"
	Response.End
End If

set rs=Dvbbs.Execute("select top 1 * from Dv_ChallengeInfo")
Dim MyForumID
MyForumID=rs("D_ForumID")

'挑战随机数
Dim MaxUserID,MaxLength
MaxLength=12
set rs=Dvbbs.Execute("select Max(userid) from [dv_user]")
MaxUserID=rs(0)

Dim num1,rndnum
Randomize
Do While Len(rndnum)<4
	num1=CStr(Chr((57-48)*rnd+48))
	rndnum=rndnum&num1
loop
MaxUserID=rndnum & MaxUserID
MaxLength=MaxLength-len(MaxUserID)
select case MaxLength
case 7
	MaxUserID="0000000" & MaxUserID
case 6
	MaxUserID="000000" & MaxUserID
case 5
	MaxUserID="00000" & MaxUserID
case 4
	MaxUserID="0000" & MaxUserID
case 3
	MaxUserID="000" & MaxUserID
case 2
	MaxUserID="00" & MaxUserID
case 1
	MaxUserID="0" & MaxUserID
case 0
	MaxUserID=MaxUserID
end select
Session("challengeWord")=MaxUserID

session("challengeWord_key")=md5(Session("challengeWord") & ":" & Dvbbs.CacheData(21,0),32)
%>
<res rayChanWord="<%=raychanword%>" tokenWord="<%=tokenword%>" forumId="<%=MyForumID%>" challengeWord="<%=MaxUserID%>">
<%
'response.write datediff("d","2003-6-5",now)
if IsSqlDataBase=1 then
	set rs=Dvbbs.Execute("select * from dv_topic where IsSmsTopic=1 and datediff(mi,LastSmsTime,LastPostTime)>=0")
else
	set rs=Dvbbs.Execute("select * from dv_topic where IsSmsTopic=1 and datediff('s',LastSmsTime,LastPostTime)>=60")
end if
do while not rs.eof
Response.Write "<msg>"
	TopicID=rs("TopicID")
	Response.Write "<id>"&topicid&"</id>"
	title=left(rs("title"),30)
	Response.Write "<sub>"&Dvbbs.CheckStr(replace(replace(title,"<",""),">",""))&"</sub>"
	LastPost_ray=split(Rs("lastpost"),"$")
	LastID=LastPost_ray(1)
	set trs=Dvbbs.Execute("select * from "&rs("PostTable")&" where AnnounceID="&LastID)
	if not (trs.eof and trs.bof) then
	body=left(trs("body"),30)
	nikename=left(trs("username"),20)
	posttime=replace(replace(replace(trs("dateandtime"),"-",""),":","")," ","")
	end if
	set trs=nothing
	Response.Write "<sender>"&Dvbbs.CheckStr(Dvbbs.htmlencode(nikename))&"</sender>"
	Response.Write "<sendTime>"&Dvbbs.CheckStr(Dvbbs.htmlencode(posttime))&"</sendTime>"
	Response.Write "<content>"&Dvbbs.CheckStr(replace(replace(body,"<",""),">",""))&"</content>"
Response.Write "</msg>"
rs.movenext
loop
set rs=nothing
%>
</res>