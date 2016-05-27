<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%	
If IsArray(Session(Dvbbs.CacheName & "UserID")) Then
	If Not IsObject(Conn) Then ConnectionDatabase
	Dim activeuser,TempNum
	If Session(Dvbbs.CacheName & "UserID")(0)="Dvbbs" Then
		activeuser="delete from Dv_online where username='"&Session(Dvbbs.CacheName & "UserID")(5)&"'"
		Conn.Execute activeuser,TempNum
		'更新缓存总用户在线数据
		MyBoardOnline.Forum_UserOnline = MyBoardOnline.Forum_UserOnline - TempNum
		Dvbbs.Name="Forum_UserOnline"
		Dvbbs.value=MyBoardOnline.Forum_UserOnline
	Else
		If IsNumeric(Session(Dvbbs.CacheName & "UserID")(0)) Then 
			activeuser="delete from Dv_online where id="&Session(Dvbbs.CacheName & "UserID")(0)
			Conn.Execute activeuser,TempNum
			'更新缓存总用户在线数据
			MyBoardOnline.Forum_GuestOnline = MyBoardOnline.Forum_GuestOnline - TempNum
			Dvbbs.Name="Forum_GuestOnline"
			Dvbbs.value=MyBoardOnline.Forum_GuestOnline
		End If 
	End If
	MyBoardOnline.Forum_Online = MyBoardOnline.Forum_Online - TempNum
	Dvbbs.Name="Forum_Online"
	Dvbbs.value=MyBoardOnline.Forum_Online
	Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
	Response.Cookies(Dvbbs.Forum_sn)("username")=""
	Response.Cookies(Dvbbs.Forum_sn)("password")=""
	Response.Cookies(Dvbbs.Forum_sn)("userclass")=""
	Response.Cookies(Dvbbs.Forum_sn)("userid")=""
	Response.Cookies(Dvbbs.Forum_sn)("userhidden")=""
	Response.Cookies(Dvbbs.Forum_sn)("usercookies")=""
	Session(Dvbbs.CacheName & "UserID")=Empty
	Session("flag")=Empty
	Response.Redirect Dvbbs.Forum_Info(11)
End If 

%>
