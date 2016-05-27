<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=1    '²Ù×÷È¨ÏÞ
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<%
dim ID,sql
ID=trim(Request("ID"))
if ID<>"" then
	sql="Delete From Admin Where ID=" & CLng(ID)
	conn.execute sql
end if
call CloseConn()      
response.redirect "AdminManage.asp"
%>


