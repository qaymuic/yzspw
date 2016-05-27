<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '²Ù×÷È¨ÏÞ
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<%
dim BigClassName,sql
BigClassName=trim(Request("BigClassName"))
if BigClassName<>"" then
	sql="delete from a1 where BigClassName='" & BigClassName & "'"
	conn.Execute sql
	sql="delete from a2 where BigClassName='" & BigClassName & "'"
	conn.Execute sql
end if
call CloseConn()      
response.redirect "ClassManage1.asp"
%>


