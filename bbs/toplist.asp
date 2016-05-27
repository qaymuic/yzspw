<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Dvbbs.LoadTemplates("paper_even_toplist")
Dvbbs.stats=template.Strings(6)

If Dvbbs.GroupSetting(1)="0" Then Dvbbs.AddErrCode(64)
Dvbbs.ShowErr()

Dim Page,Orders,ordername,Rs,SQL,keyword
Dim select1,select2,select3,select4,select5,select6,select7,select8
Dim TempStr,TempStr1,TempStr2,TempStr3,TempArray,TopTempArray
Dim TotalRec,i,Pcount

TotalRec=0
Page=request("page")
If Page="" Or Not IsNumerIc(Page) Then Page=1
Page=Clng(Page)
If Not IsNumerIc(request("orders")) Or request("orders")="" Then
	Orders=1
Else
	Orders=Cint(request("orders"))
End If
keyword=Request("keyword")
If keyword<>"" Then keyword = Dvbbs.CheckStr(keyword)
If Dvbbs.Forum_Setting(17)="0" Then keyword = ""

TempStr = template.html(7)
TopTempArray = Split(template.html(9),"||")
If Dvbbs.Forum_Setting(17)="1" Then
	TempStr = Replace(TempStr,"{$isusersearch}",TopTempArray(4))
	TempStr = Replace(TempStr,"{$keyword}",keyword)
Else
	TempStr = Replace(TempStr,"{$isusersearch}","")
End If
SQL="username,useremail,userclass,UserIM,UserPost,JoinDate,userwealth,userid"
Select Case orders
Case 1
	orders=1
	ordername=Replace(template.Strings(7),"{$toplistnum}",Dvbbs.Forum_Setting(68))
	select1="selected"
	If keyword<>"" Then keyword = " Where UserName='"&keyword&"'"
	SQL="select top "&Dvbbs.Forum_Setting(68)&" "&SQL&" from [dv_user] "&keyword&" order by UserPost desc"
	If Dvbbs.Forum_Setting(31)="0" Then Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(12)&"&action=OtherErr"
Case 2
	orders=2
	ordername=template.Strings(8)
	select2="selected"
	If keyword<>"" Then keyword = " Where UserName='"&keyword&"'"
	SQL="select top "&Dvbbs.Forum_Setting(68)&" "&SQL&" from [dv_user] "&keyword&" order by JoinDate desc"
	If Dvbbs.Forum_Setting(31)="0" Then Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(12)&"&action=OtherErr"
Case 3
	orders=3
	ordername=Replace(template.Strings(9),"{$toplistnum}",Dvbbs.Forum_Setting(68))
	select3="selected"
	If keyword<>"" Then keyword = " Where UserName='"&keyword&"'"
	SQL="select top "&Dvbbs.Forum_Setting(68)&" "&SQL&" from [dv_user] "&keyword&" order by userwealth desc"
	If Dvbbs.Forum_Setting(31)="0" Then Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(12)&"&action=OtherErr"
Case 7
	orders=7
	ordername=template.Strings(10)
	select7="selected"
	If keyword<>"" Then keyword = " Where UserName='"&keyword&"'"
	SQL="select "&SQL&" from [dv_user]  "&keyword&" order by userid desc"
	If Dvbbs.Forum_Setting(27)="0" Then Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(14)&"&action=OtherErr"
Case 8
	orders=8
	ordername=template.Strings(11)
	select8="selected"
	If keyword<>"" Then keyword = " And UserName='"&keyword&"'"
	SQL="select "&SQL&" from [dv_user] where usergroupid<=3 "&keyword&" order by usergroupid,UserPost desc"
	If Dvbbs.Forum_Setting(18)="0" Then Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(13)&"&action=OtherErr"
Case Else
	orders=1
	ordername=Replace(template.Strings(7),"{$toplistnum}",Dvbbs.Forum_Setting(68))
	select1="selected"
	If keyword<>"" Then keyword = " Where UserName='"&keyword&"'"
	SQL="select top "&Dvbbs.Forum_Setting(68)&" "&SQL&" from [dv_user] "&keyword&" order by UserPost desc"
	If Dvbbs.Forum_Setting(31)="0" Then Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(12)&"&action=OtherErr"
End Select

Dvbbs.Stats = ordername
Dvbbs.Nav()
Dvbbs.ShowErr()
Dvbbs.Head_var 0,0,template.Strings(6),"toplist.asp"

Set Rs=Dvbbs.Execute("Select Forum_PostNum,Forum_UserNum From Dv_Setup")
TempStr = Replace(TempStr,"{$postnum}",Rs(0))
TempStr = Replace(TempStr,"{$usernum}",Rs(1))

If Orders=7 and keyword="" Then
	TotalRec=Rs(1)
	If IsSqlDataBase=1 And IsBuss=1 Then
		Dim Cmd
		Set cmd = Server.CreateObject("ADODB.Command")
		Set cmd.ActiveConnection=conn
		cmd.CommandText="dv_toplist"
		cmd.CommandType=4
		cmd.Parameters.Append cmd.CreateParameter("@pagenow",3)
		cmd.Parameters.Append cmd.CreateParameter("@pagesize",3)
		cmd.Parameters.Append cmd.CreateParameter("@reture_value",3,2)
		cmd.Parameters.Append cmd.CreateParameter("@intUserRecordCount",3,2)
		cmd("@pagenow")=Page
		cmd("@pagesize")=Cint(Dvbbs.Forum_Setting(11))
		If Not IsObject(Conn) Then ConnectionDatabase
		Set Rs=Cmd.Execute
	Else
		Set Rs=Server.CreateObject("ADODB.RecordSet")
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs.Open SQL,Conn,1,1
		If Not Rs.Eof Then TotalRec=Rs.RecordCount
	End If
Else
	Set Rs=Server.CreateObject("ADODB.RecordSet")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open SQL,Conn,1,1
	If Not Rs.Eof Then TotalRec=Rs.RecordCount
End If
Dvbbs.SqlQueryNum = Dvbbs.SqlQueryNum + 1
If Rs.Eof And Rs.Bof Then
	TempStr = Replace(TempStr,"{$toplistloop}",TopTempArray(0))
	TempStr = Replace(TempStr,"{$pagelist}","")
Else
	If TotalRec Mod Dvbbs.Forum_Setting(11)=0 Then
		Pcount= TotalRec \ Dvbbs.Forum_Setting(11)
	Else
		Pcount= TotalRec \ Dvbbs.Forum_Setting(11)+1
	End If

	If Not (IsSqlDataBase=1 And Orders=7 And IsBuss=1) Then
		RS.MoveFirst
		if Page > Pcount then Page = Pcount
   		if Page < 1 then Page=1
		RS.Move (Page-1) * Dvbbs.Forum_Setting(11)
		SQL=Rs.GetRows(Dvbbs.Forum_Setting(11))
	Else
		SQL=Rs.GetRows(-1)
	End If
	Set Rs=Nothing
	'username=0,useremail=1,userclass=2,UserIM=3,UserPost=4,JoinDate=5,userwealth=6,userid=7
	TempStr1 = template.html(8)
	For i = 0 To Ubound(SQL,2)
		TempStr2 = TempStr1
		TempArray = Split(Dvbbs.HtmlEncode(Replace(SQL(3,i)&"","'","\'")),"|||")
		TempStr2 = Replace(TempStr2,"{$userid}",SQL(7,i))
		TempStr2 = Replace(TempStr2,"{$email}",SQL(1,i)&"")
		TempStr2 = Replace(TempStr2,"{$username}",Dvbbs.HtmlEncode(SQL(0,i)))
		TempStr2 = Replace(TempStr2,"{$adddate}",SQL(5,i)&"")
		TempStr2 = Replace(TempStr2,"{$userclass}",SQL(2,i)&"")
		REM 修正文章数NULL值出错问题 2004-5-21 Dv.Yz
		TempStr2 = Replace(TempStr2,"{$article}",SQL(4,i)&"")
		TempStr2 = Replace(TempStr2,"{$wealth}",SQL(6,i))
		If Ubound(TempArray)>1 Then
		TempStr2 = Replace(TempStr2,"{$homepage}",TempArray(0))
		TempStr2 = Replace(TempStr2,"{$oicq}",TempArray(1))
		Else
		TempStr2 = Replace(TempStr2,"{$homepage}","")
		TempStr2 = Replace(TempStr2,"{$oicq}","")
		End If
		TempStr3 = TempStr3 & TempStr2
	Next

	If IsSqlDataBase=1 And Orders=7 And keyword="" And IsBuss=1 Then
		TotalRec=cmd("@intUserRecordCount")
		If TotalRec Mod Dvbbs.Forum_Setting(11)=0 Then
			Pcount= TotalRec \ Dvbbs.Forum_Setting(11)
		Else
			Pcount= TotalRec \ Dvbbs.Forum_Setting(11)+1
		End If
		Set Cmd = Nothing
	End If

	TempStr = Replace(TempStr,"{$toplistloop}",TempStr3)
	TempStr = Replace(TempStr,"{$pagelist}",template.html(3))
	TempStr = Replace(TempStr,"{$page}",page)
	TempStr = Replace(TempStr,"{$Pcount}",Pcount)
	TempStr = Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	TempStr = Replace(TempStr,"{$alertcolor}",Dvbbs.mainsetting(1))
	TempStr = Replace(TempStr,"{$pagelimited}",Dvbbs.Forum_Setting(11))
	TempStr = Replace(TempStr,"{$listnum}",totalrec)
	TempStr = Replace(TempStr,"{$boardid}","0&orders="&orders)
	TempStr = Replace(TempStr,"{$emailpic}",template.pic(0))
	TempStr = Replace(TempStr,"{$oicqpic}",template.pic(1))
	TempStr = Replace(TempStr,"{$homepagepic}",template.pic(2))
	TempStr = Replace(TempStr,"{$msgpic}",template.pic(3))

	'管理团队
	If Dvbbs.Forum_Setting(18)<>"0" Then
		TempStr = Replace(TempStr,"{$myselect3}",TopTempArray(3))
	Else
		TempStr = Replace(TempStr,"{$myselect3}","")
	End If
	'用户排行
	If Dvbbs.Forum_Setting(31)<>"0" Then
		TempStr = Replace(TempStr,"{$myselect1}",TopTempArray(1))
	Else
		TempStr = Replace(TempStr,"{$myselect1}","")
	End If
	'所有用户
	If Dvbbs.Forum_Setting(27)<>"0" Then
		TempStr = Replace(TempStr,"{$myselect2}",TopTempArray(2))
	Else
		TempStr = Replace(TempStr,"{$myselect2}","")
	End If

	TempStr = Replace(TempStr,"{$ordername}",ordername)
	TempStr = Replace(TempStr,"{$pagelistnum}",Dvbbs.Forum_Setting(11))
	TempStr = Replace(TempStr,"{$select1}",select1)
	TempStr = Replace(TempStr,"{$select2}",select2)
	TempStr = Replace(TempStr,"{$select3}",select3)
	TempStr = Replace(TempStr,"{$select7}",select7)
	TempStr = Replace(TempStr,"{$select8}",select8)
	Response.Write TempStr
End If

Dvbbs.ActiveOnline()
Dvbbs.Footer()
%>
