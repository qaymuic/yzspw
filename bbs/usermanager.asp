<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Dvbbs.LoadTemplates("usermanager")
Dvbbs.Stats=Dvbbs.MemberName&template.Strings(0)
Dvbbs.Nav()
Dvbbs.Head_var 0,0,template.Strings(0),"usermanager.asp"

Dim Sql,Rs,TempStr
If dvbbs.userid=0 Then
	Dvbbs.AddErrCode(6)
	Dvbbs.Showerr()
Else
	Main()
End If
Dvbbs.ActiveOnline()
Dvbbs.Footer()

Sub Main()
Dim MainTable,User_info,i,UserFace
User_info=Session(Dvbbs.CacheName & "UserID")
UserFace="<img src="""&Dvbbs.HtmlEncode(User_info(11))&""" width="&User_info(12)&" height="&User_info(13)&" align=absmiddle >"
	MainTable=Template.Html(1)
	MainTable=Replace(MainTable,"{$TableWidth}",Dvbbs.mainsetting(0))
	MainTable=Replace(MainTable,"{$color}",Dvbbs.mainsetting(1))
	MainTable=Replace(MainTable,"{$user_Article}",User_info(8))
	MainTable=Replace(MainTable,"{$user_Group}",User_info(18))
	MainTable=Replace(MainTable,"{$user_Face}",Dv_FilterJS(UserFace))
	MainTable=Replace(MainTable,"{$user_Title}",Dvbbs.htmlencode(User_info(34)))
	MainTable=Replace(MainTable,"{$user_Wealth}",User_info(21))
	MainTable=Replace(MainTable,"{$user_EP}",User_info(22))
	MainTable=Replace(MainTable,"{$user_CP}",User_info(23))
	MainTable=Replace(MainTable,"{$user_IsBest}",User_info(28))
	MainTable=Replace(MainTable,"{$user_AddDate}",User_info(14))
	MainTable=Replace(MainTable,"{$user_Logins}",User_info(16))
	MainTable=Replace(MainTable,"{$username}",Dvbbs.Membername)
	MainTable=Replace(MainTable,"{$msg_newincept}",Dvbbs.sendmsgnum)
	MainTable=Replace(MainTable,"{$msg_incept}",incept())
	MainTable=Replace(MainTable,"{$msg_send}",allsend())
	MainTable=Replace(MainTable,"{$friend_Info}",friendlist())			'好友在线显示
	MainTable=Replace(MainTable,"{$msglist}",msg_list(5))				'短信列
	MainTable=Replace(MainTable,"{$filelist}",Fileuplist(5))			'上传文件列
	'MainTable=Replace(MainTable,"{$topiclist}",NewTopic(5))			'新帖列
	Response.Write Template.Html(0)
	Response.write MainTable
End Sub

Function msg_list(str)			'短信列
Dim Tempwrite,Msgsrc,msgflag,tablebody,i
sql="Select id,sender,title,content,flag,sendtime from DV_Message where incept='"&Dvbbs.checkstr(Dvbbs.MemberName)&"' and issend=1 and delR=0 order by flag,sendtime desc"
Set Rs=Dvbbs.Execute(sql)
If Rs.eof and Rs.bof Then
	msg_list="<tr><td class=tablebody1 align=center valign=middle colspan=5>"&template.Strings(7)&"</td></tr>"
	Exit Function
Else
SQL=Rs.GetRows(cint(str))
Rs.close:set Rs=nothing
For i=0 to Ubound(SQL,2)
	Tempwrite=Template.Html(2)
	msgflag=SQL(4,i)
	If msgflag=0 Then
		tablebody="tablebody2"
		Msgsrc="<img src="&template.pic(11)&" >"
	Else
		tablebody="tablebody1"
		Msgsrc="<img src="&template.pic(12)&" >"
	End If
	Tempwrite=Replace(Tempwrite,"{$tablebody}",tablebody)
	Tempwrite=Replace(Tempwrite,"{$msg_pic}",Msgsrc)
	Tempwrite=Replace(Tempwrite,"{$msg_name}",Dvbbs.htmlencode(SQL(1,i)))
	Tempwrite=Replace(Tempwrite,"{$msg_id}",SQL(0,i))
	Tempwrite=Replace(Tempwrite,"{$msg_topic}",Dvbbs.htmlencode(SQL(2,i)))
	If Isnull(Sql(5,i)) Or SQL(5,i) = "" Then
		Tempwrite=Replace(Tempwrite,"{$msg_time}",Now())
	Else
		Tempwrite=Replace(Tempwrite,"{$msg_time}",SQL(5,i))
	End If
	Tempwrite=Replace(Tempwrite,"{$msg_size}",len(SQL(3,i)))
	msg_list=msg_list+Tempwrite
Next
End If
End Function

Function Fileuplist(str)		'上传文件列
	Dim Tempwrite,F_imgsrc,i
	Set Rs=Dvbbs.Execute("Select F_ID,F_Filename,F_FileType,F_Type,F_Flag,F_FileSize,F_AddTime from [DV_Upfile] where F_UserID="&dvbbs.userid&" order by F_ID desc ")
	If Rs.eof and Rs.bof Then
		Fileuplist="<tr><td class=tablebody1 align=center valign=middle colspan=5>"&template.Strings(7)&"</td></tr>"
		Exit Function
	Else
		SQL=Rs.GetRows(cint(str))
		Rs.close:set Rs=nothing
	For i=0 to Ubound(SQL,2)
		Tempwrite=Template.Html(3)
		Tempwrite=Replace(Tempwrite,"{$tablebody}","class=tablebody1")
		F_imgsrc="<img src='Skins/Default/filetype/"&SQL(2,i)&".gif' border=0>"
		Tempwrite=Replace(Tempwrite,"{$file_incept}",F_imgsrc)
		Tempwrite=Replace(Tempwrite,"{$file_size}",GetSize(SQL(5,i)))
		Tempwrite=Replace(Tempwrite,"{$file_time}",SQL(6,i))
		Tempwrite=Replace(Tempwrite,"{$file_type}",F_Typename(SQL(3,i)))
		Tempwrite=Replace(Tempwrite,"{$file_topic}",Dvbbs.Htmlencode(SQL(1,i)))
		Fileuplist=Fileuplist+Tempwrite
	Next
	End If
End Function

'<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:100%">
'<tr><th height=25 align=left>-=> 最近发表的文章</th></tr>
'{$topiclist}
'</table><br>
Function NewTopic(str)			'新帖列
	Dim Tempwrite,topic,i
	Set Rs=Dvbbs.Execute("Select announceid,rootid,boardid,dateandtime,topic,body from "&Dvbbs.NowUseBbs&" where PostUserID="&dvbbs.userid&" and locktopic<2 order by announceid desc")
	If Not Rs.eof Then
		SQL=Rs.GetRows(cint(str))
	Rs.close:set Rs=nothing
	For i=0 to Ubound(SQL,2)
		topic=replace(SQL(4,i)," ","")
		If topic<>"" Then
			topic=topic
		Else
			topic=SQL(5,i)
			topic=replace(topic,chr(13),"")
			topic=replace(topic,chr(10),"")
		End If
		If Len(topic)>30 Then
			topic=left(topic,30)&"..."
		End If
		topic=Dvbbs.Htmlencode(topic)
		Tempwrite=Template.Html(4)
		Tempwrite=Replace(Tempwrite,"{$topic_title}","<a href=dispbbs.asp?boardid="&SQL(2,i)&"&id="&SQL(1,i)&"&replyid="&SQL(0,i)&"#"&SQL(0,i)&">"&topic&"</a>")
		Tempwrite=Replace(Tempwrite,"{$topic_posttime}",SQL(3,i))
		NewTopic=NewTopic+Tempwrite
	Next
	Else
		Tempwrite=Template.Html(4)
		Tempwrite=Replace(Tempwrite,"{$topic_title}","")
		Tempwrite=Replace(Tempwrite,"{$topic_posttime}","")
		NewTopic=NewTopic+Tempwrite
	End If
End Function

Function friendlist()
Dim FRs,OnlineTime,i,F_friend
If Dvbbs.Boardmaster or Dvbbs.Master Then
Set FRs=Dvbbs.Execute("Select F_friend,(Select top 1 startime from Dv_online where username = DV_Friend.F_friend) From DV_Friend Where F_mod=1 AND F_userid="&Dvbbs.Userid&" order by F_mod desc")
Else
Set FRs=Dvbbs.Execute("Select F_friend,(Select top 1 startime from Dv_online where userhidden=2 and username = DV_Friend.F_friend) From DV_Friend Where F_mod=1 AND F_userid="&Dvbbs.Userid&" order by F_mod desc")
End If
If FRs.eof and FRs.bof Then
	friendlist=template.Strings(8)
	Exit Function
Else
	SQL=FRs.GetRows(10)
End If
FRs.close:set FRs=nothing
For i=0 To Ubound(SQL,2)
	F_friend=Dvbbs.checkstr(SQL(0,i))
	If SQL(1,i)="" or isNull(SQL(1,i)) Then
		OnlineTime=Template.Strings(9)
	Else
		OnlineTime=template.Strings(10)
		OnlineTime=Replace(OnlineTime,"{$color}",Dvbbs.mainsetting(1))
		OnlineTime=Replace(OnlineTime,"{$OnlineTime}",DatedIff("n",SQL(1,i),Now()))
	End If
	friendlist=friendlist & "<a href=""javascript:openScript('messanger.asp?action=new&touser="&F_friend&"',500,400)"" ><img src="""&Dvbbs.mainpic(15)&""" alt='给好友发送短讯' border=0></a>&nbsp;<a href=dispuser.asp?name="&F_friend&" >"&F_friend&"</a> "&OnlineTime&" <br>"
Next
End Function

REM 统计已发出新短信
Function allsend()
	Set Rs=Dvbbs.Execute("Select Count(id) From DV_Message Where flag=0 and issend=1 And sender='"& Dvbbs.checkstr(Dvbbs.MemberName) &"'")
	allsend=Rs(0)
	Rs.close
	If isnull(allsend) Then allsend=0
End Function
REM 统计收件箱中的短信
Function incept()
	incept=0
	Set Rs=Dvbbs.Execute("Select Count(id) From DV_Message Where issend=1 and delR=0 And incept='"& Dvbbs.checkstr(Dvbbs.MemberName) &"'")
	incept=Rs(0)
	Rs.close
	If isnull(incept) Then incept=0
End Function

Function F_Typename(str)
DIM TempName
TempName=split(Dvbbs.lanstr(5),"||")
	If not IsEmpty(str) and isNumeric(str) Then
		Select case str
		case 1
			F_Typename=TempName(1)
		case 2
			F_Typename=TempName(2)
		case 3
			F_Typename=TempName(3)
		case 4
			F_Typename=TempName(4)
		case Else
			F_Typename=TempName(0)
		End Select
	End If
End Function

Function Dv_FilterJS(v)
	If  Not Isnull(V) Then
		Dim t
		Dim re
		Dim reContent
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(&#)"
		t=re.Replace(v,"<I>&#</I>")
		re.Pattern="(script)"
		t=re.Replace(t,"<I>script</I>")
		re.Pattern="(js:)"
		t=re.Replace(t,"<I>js:</I>")
		re.Pattern="(value)"
		t=re.Replace(t,"<I>value</I>")
		re.Pattern="(about:)"
		t=re.Replace(t,"<I>about:</I>")
		re.Pattern="(file:)"
		t=re.Replace(t,"<I>file:</I>")
		re.Pattern="(Document.cookie)"
		t=re.Replace(t,"<I>Documents.cookie</I>")
		re.Pattern="(vbs:)"
		t=re.Replace(t,"<I>vbs:</I>")
		re.Pattern="(on(mouse|Exit|error|click|key))"
		t=re.Replace(t,"<I>on$2</I>")
		Dv_FilterJS=t
		Set Re=Nothing
	End If 
End Function


Function GetSize(size)
if isEmpty(size) then exit function
	if size>1024 then
 		   size=(size\1024)
 		   GetSize=size & "&nbsp;KB"
	else
		   GetSize=size & "&nbsp;Byte"
 	end if
 	if size>1024 then
 		   size=(size/1024)
 		   GetSize=Formatnumber(size,2) & "&nbsp;MB"		
 	end if
 	if size>1024 then
 		   size=(size/1024)
 		   GetSize=Formatnumber(size,2) & "&nbsp;GB"	   
 	end if   
End Function

'Function SplitStr(str1,str2,str3)
'If IsEmpty(str1) or IsEmpty(str2) or not isNumeric(str3) Then Exit Function
'SplitStr=split(str1,str2)(str3)
'End Function
%>