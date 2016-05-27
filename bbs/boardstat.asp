<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<%
Dvbbs.LoadTemplates("boardstat")
Dim sinfo,allnum
Dim boardstat_info(2),i
Dim Sql,Yrs
Dim BoardHide

If request("action")="lastbbsnum" then
	sinfo="PostNum"
	allnum=Dvbbs.cachedata(8,0)
	Dvbbs.stats=template.Strings(1)
Elseif request("action")="lasttopicnum" then
	sinfo="TopicNum"
	allnum=Dvbbs.CacheData(7,0)
	Dvbbs.stats=template.Strings(2)
Elseif request("reaction")<>"" then
	sinfo="todayNum"
	allnum=MyBoardOnline.Forum_Online
	If Dvbbs.boardid=0 then
		Dvbbs.stats=template.Strings(3)
	Else 
		Dvbbs.stats=Dvbbs.Boardtype&template.Strings(4)
	End If 
Else
	sinfo="todayNum"
	allnum=Dvbbs.CacheData(9,0)
	Dvbbs.stats=template.Strings(5)
End If
Dim Stats
Stats=Dvbbs.stats
Dvbbs.Nav()
Dvbbs.Showerr()

If Dvbbs.boardid=0 Then
	dvbbs.Head_var 0,0,template.Strings(0),"boardstat.asp"
Else
	Dvbbs.head_var 1,Dvbbs.Board_Data(4,0),"",""
End if
Dim Redcolor
Redcolor=Dvbbs.mainsetting(1)
Main()
Dvbbs.ActiveOnline()
Dvbbs.Footer()

Sub Main()
	dim MainTable,Toplist
	Toplist=Template.Html(0)
	Toplist=Replace(Toplist,"{$boardid}",Dvbbs.boardid)
	Toplist=Replace(Toplist,"{$boardinfolist}",stateinfo_list(Dvbbs.boardid))
	MainTable=Template.Html(1)
	If request("reaction")="onlineinfo" then
		MainTable=Replace(Maintable,"{$statelist}",Onlinemain(Dvbbs.boardid))
	Elseif request("reaction")="onlineUserinfo" then
		MainTable=Replace(Maintable,"{$statelist}",userinfo)
	Elseif request("reaction")="online" then
		MainTable=Replace(Maintable,"{$statelist}",Onlinestat)
	Else
		MainTable=Replace(Maintable,"{$statelist}",Statelist(sinfo))
	End if
	Response.Write Toplist
	Response.write MainTable
End sub

Function stateinfo_list(str)
	Dim Tempwrite
	If Dvbbs.boardid=0 then
		sql="select Forum_TopicNum,Forum_PostNum,Forum_TodayNum from [Dv_setup]"
	Else
		sql="select TopicNum,PostNum,TodayNum from [Dv_board] where boardid="&Dvbbs.boardid
	End if
	Set	Yrs=Dvbbs.Execute(sql)
	IF not Yrs.eof then
		Sql=Yrs.GetString(,,"|||","","")
	End IF
	Yrs.close:Set Yrs=nothing
	Sql=split(Sql,"|||")
	For i=0 To 2
		boardstat_info(i)=Sql(i)
	Next
	If request("reaction")<>"" then
		If clng(str)=0 then
			Tempwrite=template.Strings(6)
			Tempwrite=Replace(Tempwrite,"{$Allonlinenum}",MyBoardOnline.Forum_Online)
		Else
			Tempwrite=template.Strings(6)&Template.Strings(7)
			Tempwrite=Replace(Tempwrite,"{$Allonlinenum}",MyBoardOnline.Forum_Online)
			Tempwrite=Replace(Tempwrite,"{$boardtype}",Dvbbs.Boardtype)
			Tempwrite=Replace(Tempwrite,"{$boardusernum}",MyBoardOnline.Board_UserOnline)
			Tempwrite=Replace(Tempwrite,"{$boardguestnum}",MyBoardOnline.Board_GuestOnline)
		End if
		stateinfo_list=Tempwrite
	else
		Tempwrite=template.Strings(8)
		if clng(str)=0 then
			Tempwrite=Replace(Tempwrite,"{$boardtype}",Dvbbs.Forum_info(0))
		else
			Tempwrite=Replace(Tempwrite,"{$boardtype}",Dvbbs.Boardtype)
		end if
		Tempwrite=Replace(Tempwrite,"{$TopicNum}",boardstat_info(0))
		Tempwrite=Replace(Tempwrite,"{$BbsNum}",boardstat_info(1))
		Tempwrite=Replace(Tempwrite,"{$TodayNum}",boardstat_info(2))
		stateinfo_list=Tempwrite
	end if
End function

Function Statelist(sinfo)
	Dim Tempslist,Tempslist1,i,k
	sql="Select "&sinfo&",boardid,BoardType,Board_Setting from Dv_board order by "&sinfo&" desc"
	Set Yrs=Dvbbs.execute(Sql)
	If Yrs.eof then
		Tempslist=Template.Html(2)
		Tempslist=Replace(Tempslist,"{$Stattitle1}",template.Strings(9))
		Tempslist=Replace(Tempslist,"{$Stattitle2}",template.Strings(10))
		Tempslist=Replace(Tempslist,"{$Stattitle3}",Stats)
		Statelist="<tr><td colspan=3 class=tablebody1 align=center>"
		Statelist=Statelist&template.Strings(11)
		Statelist=Statelist&"</td></tr>"
		Statelist=Tempslist&Statelist
	Else
		Sql=YRs.GetRows(-1)
		Yrs.close:set Yrs=nothing
		k=11
		For i=0 to Ubound(sql,2)
			k=k+1
			BoardHide=Split(Sql(3,i),",")(1)
			Tempslist1=Tempslist1&template.html(3)
			If Dvbbs.boardid=Clng(Sql(1,i)) then
				Tempslist1=Replace(Tempslist1,"{$selboard}",2)
			Else
				Tempslist1=Replace(Tempslist1,"{$selboard}",1)
			End if
			If LoadHideBBS then
				Tempslist1=Replace(Tempslist1,"{$boardtype}","<a href=list.asp?boardid="&Sql(1,i)&">"&Sql(2,i)&"</a>")
			Else
				Tempslist1=Replace(Tempslist1,"{$boardtype}",template.Strings(13))
			End if
			Tempslist1=Replace(Tempslist1,"{$statpic}",template.pic(k))
			If k>20 then k=11
			If allnum=0 then
				Tempslist1=Replace(Tempslist1,"{$statwidth}",1)
			Else
				Tempslist1=Replace(Tempslist1,"{$statwidth}",Cint(replace(FormatPercent(Sql(0,i)/allnum),"%",""))*5.1)
			End if
			Tempslist1=Replace(Tempslist1,"{$statnum}",Sql(0,i))
		Next
		Tempslist=Template.Html(2)
		Tempslist=Replace(Tempslist,"{$Stattitle1}",template.Strings(9))
		Tempslist=Replace(Tempslist,"{$Stattitle2}",template.Strings(10))
		Tempslist=Replace(Tempslist,"{$Stattitle3}",Stats)
		Tempslist=Tempslist&Tempslist1
		Statelist=Tempslist
	End if
End function

Function Onlinestat()
	Dim Tempslist,Tempslist1,Tempslist2,i,k
	Dim Othernum,OnlineNum
	Othernum=Allnum
	Sql="select boardid,BoardType,board_setting,(select count(id) from dv_online where boardid=dv_board.boardid) from [Dv_board] order by Boardid"
	Set Yrs=Dvbbs.execute(Sql)
	If Yrs.eof then
		Tempslist=Template.Html(2)
		Tempslist=Replace(Tempslist,"{$Stattitle1}",template.Strings(9))
		Tempslist=Replace(Tempslist,"{$Stattitle2}",template.Strings(10))
		Tempslist=Replace(Tempslist,"{$Stattitle3}",template.Strings(4))
		Onlinestat="<tr><td colspan=3 class=tablebody1 align=center>"
		Onlinestat=Onlinestat&template.Strings(11)
		Onlinestat=Onlinestat&"</td></tr>"
		Onlinestat=Tempslist&Onlinestat
	Else
		Sql=YRs.GetRows(-1)
		Yrs.close:Set Yrs=nothing
		k=11
		For i=0 to Ubound(sql,2)
			k=k+1
			BoardHide=Split(sql(2,i),",")(1)
			OnlineNum=sql(3,i)
			If isnull(onlineNum) then onlineNum=0
			Othernum=Othernum-onlineNum
			Tempslist1=Tempslist1&template.html(3)
			If Dvbbs.boardid=Clng(sql(0,i)) then
				Tempslist1=Replace(Tempslist1,"{$selboard}",2)
			Else
				Tempslist1=Replace(Tempslist1,"{$selboard}",1)
			End if
			If LoadHideBBS then
				Tempslist1=Replace(Tempslist1,"{$boardtype}","<a href=list.asp?boardid="&sql(0,i)&">"&sql(1,i)&"</a>")
			Else
				Tempslist1=Replace(Tempslist1,"{$boardtype}",template.Strings(13))
			End if
			Tempslist1=Replace(Tempslist1,"{$statpic}",template.pic(k))
			If k>20 then k=11
			If allnum=0 then
				Tempslist1=Replace(Tempslist1,"{$statwidth}",1)
			Else
				Tempslist1=Replace(Tempslist1,"{$statwidth}",Cint(replace(FormatPercent(onlineNum/allnum),"%",""))*5.1)
			End if
			Tempslist1=Replace(Tempslist1,"{$statnum}",onlineNum)
		Next
		IF Othernum>0 then
			Tempslist2="<tr><td class=""tablebody2"" align=""middle"">"
			Tempslist2=Tempslist2&template.Strings(12)
			Tempslist2=Tempslist2&"</td><td class=""tablebody1""><img height=""8"" src="""
			Tempslist2=Tempslist2&template.pic(k)
			Tempslist2=Tempslist2&""" width="""
			Tempslist2=Tempslist2&Cint(replace(FormatPercent(OtherNum/allnum),"%",""))*5.1
			Tempslist2=Tempslist2&"""></td><td class=""tablebody2"" align=""middle"">"
			Tempslist2=Tempslist2&Othernum
			Tempslist2=Tempslist2&"</td></tr>"
			Tempslist1=Tempslist2&Tempslist1
		End if
		Tempslist=Template.Html(2)
		Tempslist=Replace(Tempslist,"{$Stattitle1}",template.Strings(9))
		Tempslist=Replace(Tempslist,"{$Stattitle2}",template.Strings(10))
		Tempslist=Replace(Tempslist,"{$Stattitle3}",template.Strings(4))
		Tempslist=Tempslist&Tempslist1
		Onlinestat=Tempslist
	End if
End function

Function onlinemain(str)
	Dim Tempslist,Tempslist1
	dim page_count,Pcount
	dim totalrec,PageListNum
	dim onlinename,ipinfo,sysinfo,actiontime
	Dim titlepic,maslink1,maslink2
	PageListNum=Cint(Dvbbs.Forum_Setting(11))
	Tempslist=Template.html(4)	

	If Clng(str)=0 then
		sql="select username,stats,browser,ip,userhidden,lastimebk,startime,usergroupid,id,(select TitlePic from [Dv_UserGroups] where UserGroupID=Dv_online.usergroupid) from [Dv_online] order by startime"
	Else
		sql="select username,stats,browser,ip,userhidden,lastimebk,startime,usergroupid,id,(select TitlePic from [Dv_UserGroups] where UserGroupID=Dv_online.usergroupid) from [Dv_online] where Boardid="&Clng(str)&" order by startime"
	End if
	dim totalPages,currentPage
	currentPage=request.querystring("page")
	If currentpage="" or not isnumeric(currentpage) then
		currentpage=1
	Else
		currentpage=clng(currentpage)
	End if
	If Not IsObject(Conn) Then ConnectionDatabase
	Set Yrs=server.createobject("adodb.recordset")
	Yrs.open Sql,conn,1,1
	Dvbbs.SqlQueryNum=Dvbbs.SqlQueryNum+1
	If Yrs.eof or Yrs.bof then
		onlinemain="<tr><td colspan=6 class=tablebody1 align=center>"
		onlinemain=onlinemain&template.Strings(11)
		onlinemain=onlinemain&"</td></tr>"
		onlinemain=Tempslist&onlinemain
	Else
		page_count=0
      	totalrec=Yrs.recordcount
	  	If totalrec mod PageListNum=0 then
				Pcount= totalrec \ PageListNum
	  	Else
				Pcount= totalrec \ PageListNum+1
	  	End if
		if currentpage > Pcount then currentpage = Pcount
		if currentpage<1 then currentpage=1
		Yrs.Move (currentpage-1) * Cint(PageListNum)
		While (not Yrs.eof) and (not page_count = Cint(PageListNum))
			If Clng(Yrs("userhidden"))=1 then
				If Dvbbs.master or Dvbbs.superboardmaster or trim(Yrs(0))=Dvbbs.membername then
					onlinename="<a href=dispuser.asp?name="&Dvbbs.HtmlEncode(Yrs("username"))&" target=_blank>"&Dvbbs.HtmlEncode(Yrs("username"))&"</a>"
					If Clng(Yrs("usergroupid"))=9999 then
						titlepic=template.pic(0)
					Else
						titlepic=Yrs(9)
					End if
					Maslink1="<a title="&template.Strings(18)&" href=""javascript:openScript('messanger.asp?action=new&touser="&Dvbbs.htmlencode(Yrs("username"))&"',500,400)"">"
					Maslink2="</a>"
				Else
					onlinename=template.Strings(14)
					titlepic=template.pic(7)
					Maslink1=""
					Maslink2=""
				End if
			Else
				If Clng(Yrs("usergroupid"))=9999 then
					titlepic=template.pic(0)
				Else
					titlepic=Yrs(9)
				End if
				If Clng(Yrs("Usergroupid"))=7 then
					onlinename=Dvbbs.HtmlEncode(Yrs("username"))
					Maslink1=""
					Maslink2=""
				Else
					onlinename="<a href=dispuser.asp?name="&Dvbbs.HtmlEncode(Yrs("username"))&" target=_blank>"&Dvbbs.HtmlEncode(Yrs("username"))&"</a>"
					Maslink1="<a title="&template.Strings(18)&" href=""javascript:openScript('messanger.asp?action=new&touser="&Dvbbs.htmlencode(Yrs("username"))&"',500,400)"">"
					Maslink2="</a>"
				End if
			End if
			If trim(Yrs("username"))=Dvbbs.membername then
				onlinename="<font color="&Redcolor&">"&onlinename&"</font>"
				Maslink1=""
				Maslink2=""
			Elseif Dvbbs.userid=0 then
				Maslink1=""
				Maslink2=""
			Elseif Clng(Dvbbs.UserGroupId)=7 then
				If Clng(Session(Dvbbs.CacheName & "UserID")(0))=Clng(Yrs("id")) then onlinename="<font color="&Redcolor&">"&onlinename&"</font>"
			End if
			If Cint(Dvbbs.GroupSetting(30))=1 then
				ipinfo="<a href=look_ip.asp?action=lookip&ip="&Yrs("ip")&">"&Yrs("ip")&"</a>"
			Else
				ipinfo=template.Strings(15)
			End if
			sysinfo=Replace(replace(Yrs("browser"),"|",","),template.Strings(16),"IE")
			actiontime=GetTimeStr(Yrs("startime"),Yrs("lastimebk"))
			If page_count=0 then
				Tempslist1=template.html(5)
			Else
				Tempslist1=Tempslist1&template.html(5)
			End if
			Tempslist1=Replace(Tempslist1,"{$maslink1}",Maslink1)
			Tempslist1=Replace(Tempslist1,"{$maslink2}",Maslink2)
			REM ：因新建组可能没有填写等级图片 2004-5-11 YZ
			If Titlepic = "" Or Isnull(Titlepic) Then
				Titlepic = Template.Pic(4)
			End If
			Tempslist1=Replace(Tempslist1,"{$titlepic}",titlepic)
			Tempslist1=Replace(Tempslist1,"{$onlinename}",onlinename)
			Tempslist1=Replace(Tempslist1,"{$onlinestat}",Dvbbs.HtmlEncode(Yrs("stats")))
			Tempslist1=Replace(Tempslist1,"{$onlinesys}",sysinfo)
			Tempslist1=Replace(Tempslist1,"{$onlineip}",ipinfo)
			Tempslist1=Replace(Tempslist1,"{$onlinetime}",actiontime)
			Yrs.movenext
		  	page_count = page_count + 1
		Wend
		Tempslist1=Tempslist1&ShowPage(CurrentPage,Pcount,totalrec,PageListNum,redcolor)
		Tempslist=Tempslist&Tempslist1
		Onlinemain=Tempslist
	End if
END function

'分页输出
Function ShowPage(CurrentPage,Pcount,totalrec,PageNum,redcolor)
Dim SearchStr
SearchStr="Boardid="&Dvbbs.boardid&"&reaction=onlineinfo"
ShowPage=template.html(6)
ShowPage=Replace(ShowPage,"{$CurrentPage}",CurrentPage)
ShowPage=Replace(ShowPage,"{$Pcount}",Pcount)
ShowPage=Replace(ShowPage,"{$PageNum}",PageNum)
ShowPage=Replace(ShowPage,"{$totalrec}",totalrec)
ShowPage=Replace(ShowPage,"{$SearchStr}",SearchStr)
ShowPage=Replace(ShowPage,"{$redcolor}",redcolor)
End Function

Function userinfo()
	Dim i,k
	Dim Tempslist,Tempslist1
	Sql="select UserGroupID,title,(select count(id) from [dv_online] where UserGroupID=dv_usergroups.UserGroupID) as Grouponline from [Dv_UserGroups] where ParentGID=0 order by UserGroupID"
	Set Yrs=Dvbbs.execute(sql)
	Sql=Yrs.Getrows(-1)
	Yrs.close:set Yrs=nothing
	k=11
	For i=0 to Ubound(sql,2)
		k=k+1
		Tempslist1=Tempslist1&template.html(3)
		Tempslist1=Replace(Tempslist1,"{$boardtype}",Sql(1,i))
		If Clng(Dvbbs.UserGroupID)=Clng(Sql(0,i)) then
			Tempslist1=Replace(Tempslist1,"{$selboard}",2)
		Else
			Tempslist1=Replace(Tempslist1,"{$selboard}",1)
		End if
		Tempslist1=Replace(Tempslist1,"{$statpic}",template.pic(k))
		If k>20 then k=11
		Tempslist1=Replace(Tempslist1,"{$statwidth}",Cint(replace(FormatPercent(Sql(2,i)/allnum),"%",""))*5.1)
		Tempslist1=Replace(Tempslist1,"{$statnum}",Sql(2,i))
	Next
	Tempslist=template.html(2)
	Tempslist=Replace(Tempslist,"{$Stattitle1}",template.Strings(17))
	Tempslist=Replace(Tempslist,"{$Stattitle2}",template.Strings(10))
	Tempslist=Replace(Tempslist,"{$Stattitle3}",template.Strings(4))
	Tempslist=Tempslist&Tempslist1
	userinfo=Tempslist
End function

Function LoadHideBBS()
	LoadHideBBS=True
	If CInt(BoardHide)=1  And CInt(Dvbbs.GroupSetting(37))<>1 Then
		LoadHideBBS=False
	End If
End Function

Function GetTimeStr(Str1,Str2)
	Dim GetTime
	GetTime=int(Datediff("n",str1,str2))
	If GetTime>59 Then 
		GetTimeStr=(GetTime \ 60)&"<font color="&redcolor&">h：</font>"
		GetTimeStr=GetTimeStr&(GetTime mod 60)
	Else
		GetTimeStr=GetTime
	End if
	GetTimeStr=GetTimeStr&"<font color="&redcolor&">m</font>"
	GetTime=int(Datediff("n",str2,Now()))
	If GetTime>59 Then 
		GetTimeStr=GetTimeStr&" | "& (GetTime \ 60)&"<font color="&redcolor&">h：</font>"
		GetTimeStr=GetTimeStr&(GetTime mod 60)
	Else
		GetTimeStr=GetTimeStr&" | "& GetTime
	End if
	GetTimeStr=GetTimeStr&"<font color="&redcolor&">m</font>"
End Function
%>