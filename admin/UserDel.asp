<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '²Ù×÷È¨ÏÞ
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<%
dim UserID,sql,rs
UserID=trim(Request("UserID"))
if UserID<>"" then
	sql="delete from Userinfo where UserID=" & Clng(UserID)
	conn.Execute sql
end if
call CloseConn()      
response.redirect "UserManage.asp"
%>


