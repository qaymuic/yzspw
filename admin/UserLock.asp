<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '²Ù×÷È¨ÏÞ
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<%
dim UserID,Action,sql
UserID=trim(Request("UserID"))
Action=trim(request("Action"))
if UserID<>"" then
	if Action="Lock" then
		sql="Update Userinfo set LockUser=true where UserID=" & CLng(UserID)
	else
		sql="Update Userinfo set LockUser=false where UserID=" & CLng(UserID)
	end if
	conn.Execute sql
    call CloseConn()      
end if
response.redirect "UserManage.asp"
%>


