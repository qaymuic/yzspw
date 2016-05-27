<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/md5.asp"-->
<%
'2003-12-3 Edit by YangZheng
Dvbbs.Loadtemplates("paper_even_toplist")
Dim cansmallpaper
cansmallpaper=false
Dvbbs.stats=Template.Strings(16)
GetBoardPermission
Dvbbs.Nav
Dvbbs.ShowErr()
If Cint(Dvbbs.GroupSetting(17))=0 then
	Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(18)&"&action=OtherErr"
Else
	If Dvbbs.userid=0 then
		Dvbb.membername=Template.Strings(19)
	End If
	cansmallpaper=True 
End If
Dvbbs.ShowErr()

If request("action")="savepaper" then
	Call savepaper()
Else
	call main()
End If
Dvbbs.ActiveOnline
Dvbbs.Footer()

Sub main()
	Dim redcolor,ispass1,ispass2
	Dim Tempwrite
	redcolor=Dvbbs.Mainsetting(1)
	If Dvbbs.Forum_Setting(35) then
		ispass1=Template.Strings(21)
	Else
		ispass1=Template.Strings(20)
	End if
	If Dvbbs.Forum_Setting(34) then
		ispass2=Template.Strings(21)
	Else
		ispass2=Template.Strings(20)
	End if
	If IsSqlDataBase=1 Then
		Dvbbs.execute("delete from Dv_smallpaper where datediff(d,s_addtime,"&SqlNowString&")>1")
	Else
		Dvbbs.execute("delete from Dv_smallpaper where datediff('d',s_addtime,"&SqlNowString&")>1")
	End If
	Dvbbs.Name = "BoardInfo_" & Dvbbs.BoardID
	Dvbbs.LoadBoardNews_Paper(Dvbbs.BoardID)
	Dvbbs.head_var 1,Dvbbs.Board_Data(4,0),"",""
	Tempwrite=Template.html(10)
	Tempwrite=Replace(Tempwrite,"{$username}",Dvbbs.HtmlEnCode(Dvbbs.Membername))
	Tempwrite=Replace(Tempwrite,"{$password}",Dvbbs.Memberword)
	Tempwrite=Replace(Tempwrite,"{$redcolor}",redcolor)
	Tempwrite=Replace(Tempwrite,"{$paymoney}",Dvbbs.GroupSetting(46))
	Tempwrite=Replace(Tempwrite,"{$ispass1}",ispass1)
	Tempwrite=Replace(Tempwrite,"{$ispass2}",ispass2)
	Tempwrite=Replace(Tempwrite,"{$boardid}",Dvbbs.Boardid)
	Response.Write Tempwrite
End Sub

Sub savepaper()
	Dim username
	Dim password
	Dim title
	Dim content
	userName=Dvbbs.Checkstr(trim(request.form("username")))
	PassWord=Dvbbs.Checkstr(trim(request.form("password")))
	title=Dvbbs.Checkstr(trim(request.form("title")))
	Content=Dvbbs.Checkstr(request.form("Content"))
	If Dvbbs.chkpost=False Then
		Dvbbs.AddErrCode(16)
	End If
	If UserName="" Or Dvbbs.strLength(userName)>Cint(Dvbbs.Forum_setting(41)) Or Dvbbs.strLength(userName) < Cint(Dvbbs.Forum_setting(40)) then
		Dvbbs.AddErrCode(66)
	End If
	If title="" Then
		Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(22)&"&action=OtherErr"
	ElseIf Dvbbs.strLength(title)>80 then
		Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(23)&"&action=OtherErr"
	End If
	If content="" Then
		Dvbbs.AddErrCode(80)
	ElseIf Dvbbs.strLength(content)>500 then
		Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(24)&"&action=OtherErr"
	End If
	Dvbbs.ShowErr()
	'客人不允许发，验证用户
	If cansmallpaper Then
		If Not ChkUserLogin(password,username) Then
			Dvbbs.AddErrCode(12)
			Dvbbs.Showerr()
		End If
		Dim Rs,SQL
		Set Rs=server.createobject("adodb.recordset")
		sql="Select userWealth From [Dv_User] Where UserName='"&UserName&"'"
		Rs.open sql,conn,1,3
		If Not(rs.eof and rs.bof) Then
			If CLng(rs("UserWealth"))<Clng(Dvbbs.GroupSetting(46)) Then
				Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(25)&"&action=OtherErr"
			Else
				rs("UserWealth")=rs("UserWealth")-Cint(Dvbbs.GroupSetting(46))
				rs.update
			End If
		Else
			If Dvbbs.userid<>0 or username<>Template.Strings(19) Then
				Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(26)&"&action=OtherErr"
			End If
		End If
		Rs.close:Set Rs=Nothing
	End If
	Dvbbs.ShowErr()
	sql="insert into Dv_smallpaper (s_boardid,s_username,s_title,s_content) values "&_
		"("&_
		Dvbbs.boardid&",'"&_
		username&"','"&_
		title&"','"&_
		content&"')"
		'response.write sql
	Dvbbs.execute(sql)
	'发表小字报成功后RELOAD缓存
	Dvbbs.Name = "BoardInfo_" & Dvbbs.BoardID
	Dvbbs.LoadBoardNews_Paper(Dvbbs.BoardID)
	Dvbbs.head_var 1,Dvbbs.Board_Data(4,0),"",""
	Dvbbs.Dvbbs_suc("<li>"&Template.Strings(27))
End Sub

'检查用户身份
Public Function ChkUserLogin(password,username)
	Dim SQL,Rs
	ChkUserLogin=False
	If PassWord<>Dvbbs.MemberWord Then PassWord=md5(PassWord,16)
	'校验用户名和密码是否合法
	If Not IstrueName(UserName) Then Dvbbs.AddErrCode(18)
	If Len(PassWord)<>16 AND Len(PassWord)<>32 Then Dvbbs.AddErrCode(18)
	If UserName=Dvbbs.MemberName Then PassWord=Dvbbs.MemberWord
	Dvbbs.ShowErr()
	SQL = "Select UserGroupID,userpassword,lockuser,TruePassWord From [Dv_User] Where UserName='"&UserName&"' "
	Set Rs=Dvbbs.Execute(SQL)
	If Not Rs.EOF Then
		If PassWord<>rs(1) And PassWord<>rs(3) Then
			ChkUserLogin=False
		ElseIf rs(2)=1 or rs(0)=5 Then
			ChkUserLogin=False
		Else
			ChkUserLogin=True
		End If
	Else
		Exit Function 
	End If:Set Rs = Nothing 
End Function
'通用函数
Function IstrueName(uName)
	IstrueName=False
	If InStr(uName,"=")>0 Then Exit Function
	If InStr(uName,"%")>0 Then Exit Function 
	If InStr(uName,Chr(32))>0 Then Exit Function 
	If InStr(uName,"?")>0 Then Exit Function 
	If InStr(uName,"&")>0 Then Exit Function 
	If InStr(uName,";")>0 Then Exit Function 
	If InStr(uName,",")>0 Then Exit Function 
	If InStr(uName,"'")>0 Then Exit Function 
	If InStr(uName,Chr(34))>0 Then Exit Function 
	If InStr(uName,chr(9))>0 Then Exit Function 
	If InStr(uName,"")>0 Then Exit Function 
	If InStr(uName,"$")>0 Then Exit Function
	IstrueName=True 	
End Function
%>