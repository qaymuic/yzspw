<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '²Ù×÷È¨ÏÞ
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<%
dim id,sql,rs
id=trim(Request("id"))
if UserID<>"" then
	sql="delete from jgpg where id=" & Clng(id)
	conn.Execute sql
end if
call CloseConn()      
response.redirect "fangjgpg.asp"
%>


