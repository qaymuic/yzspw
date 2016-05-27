<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '²Ù×÷È¨ÏÞ
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<%
dim SmallClassID,sql
SmallClassID=trim(Request("SmallClassID"))
if SmallClassID<>"" then
	sql="delete from SmallClass where SmallClassID="&Cint(SmallClassID)&""
	conn.Execute sql
end if
call CloseConn()      
response.redirect "ClassManage.asp"
%>


