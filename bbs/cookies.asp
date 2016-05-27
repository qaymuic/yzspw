<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/Dv_ClsOther.asp" -->
<%
Dim action
action=Request("action")
Select Case action
	Case "hidden"
		Call hidden()
	Case "online"
		Call online()
	Case "stylemod"
		Call stylemod()
	Case "setlistmod"
		Call SetListmod
	Case "setlistmoda"
		Call SetListmoda		
	Case Else
End Select
If IsNull(Request.ServerVariables("HTTP_REFERER")) or Request.ServerVariables("HTTP_REFERER")="" Then
	response.redirect "index.asp"
Else
	response.redirect Request.ServerVariables("HTTP_REFERER")
End If

Sub hidden()
	If Not Dvbbs.founduser Then
		Dvbbs.AddErrCode "34":Dvbbs.Showerr()
	End If
	Dvbbs.execute("update [Dv_online] set userhidden=1 where userid="&Dvbbs.userid)
	Dvbbs.execute("update [Dv_user] set userhidden=1 where userid="&Dvbbs.userid)
	Dim usercookies
	usercookies=request.cookies(Dvbbs.Forum_sn)("usercookies")
	If IsNull(usercookies) or usercookies="" then usercookies="0"
	Select Case usercookies
		Case "0"
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 1
   			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 2
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 3
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	End Select 
	Response.Cookies(Dvbbs.Forum_sn)("userhidden") = 1
	Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
End Sub

Sub online()
	If Not Dvbbs.founduser Then
 		Dvbbs.AddErrCode "34":Dvbbs.Showerr()
	End If
	Dvbbs.execute("update [dv_online] set userhidden=2 where userid="&Dvbbs.userid)
	Dvbbs.execute("update [Dv_user] set userhidden=2 where userid="&Dvbbs.userid)
	Dim  usercookies
	usercookies=request.cookies(Dvbbs.Forum_sn)("usercookies")
	If IsNull(usercookies) or usercookies="" Then usercookies="0"
	Select Case usercookies
		Case "0"
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 1
   			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 2
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		Case 3
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	End select
	Response.Cookies(Dvbbs.Forum_sn)("userhidden") = 2
	Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
End Sub

Sub stylemod()
	Response.Cookies("skin").expires= date+7
	Response.Cookies("skin").path=Dvbbs.cookiepath
	If Not isnumeric(request("skinid")) Then
		Dvbbs.AddErrCode "35":Dvbbs.Showerr()
	End If
	Dim cssid,skinid
	cssid=Request("cssid")
	If Not isnumeric(cssid) Then
		cssid=0
	End If
	skinid=Request("skinid")
	If CInt(skinid)<>0 Then
		Response.Cookies("skin")("skinid_"&Dvbbs.boardid)=skinid
		Response.Cookies("skin")("cssid_"&Dvbbs.boardid)=cssid
	Else
		Response.Cookies("skin")("skinid_"&Dvbbs.boardid)=""
		Response.Cookies("skin")("cssid_"&Dvbbs.boardid)=""
	End If
End Sub
Sub SetListmod()
	Response.Write "<script language=""javascript"">"
	Response.Write "parent.ReShowList("&Request("id")&");"
	Response.Write "</script>"
	Response.Cookies("List").path=Dvbbs.cookiepath
	Response.Cookies("List").expires= date+7
	Response.Cookies("List")("list"&Request("id"))=request("thisvalue")
	Response.End 
End Sub
%>
