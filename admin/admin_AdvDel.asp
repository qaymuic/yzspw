<!--#include file="conn.asp"-->
<%
Const PurviewLevel=2
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim id,sql,rs
id=replace(request("id"),"'","")
if request("action")="É¾³ý¹ã¸æ" then
if id<>"" then
	sql="delete from adv where id in ("&id&")"
	'response.write sql
	'response.end
	conn.Execute sql
end if
'else
	'sql="delete ry_article"
	'conn.Execute sql
end if

call CloseConn()      
response.redirect "admin_advManage.asp?page="&request("page")
%>


