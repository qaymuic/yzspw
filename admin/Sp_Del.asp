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
if id<>"" then
	sql="delete from spw where id=" & Clng(id)
	conn.Execute sql
end if
call CloseConn()      
response.redirect "sp_Manage.asp"
%>


