
<%@language=vbscript codepage=936 %>
<%
if session("admin")=empty then 
response.redirect "admin.asp"
end if
%>
<!--#include file="conn.asp"-->
<%
dim SpecialID,sql
SpecialID=trim(Request("SpecialID"))
if SpecialID<>"" then
	sql="delete from Special where SpecialID="&Cint(SpecialID)&""
	conn.Execute sql
    call CloseConn()      
end if
response.redirect "adminSpecialManage.asp"
%>
