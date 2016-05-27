<%@language=vbscript codepage=936 %>
<!--#include file="conn.asp"-->
<%
dim id,sql
id=trim(Request("id"))
if id<>"" then
	sql="delete from ytiinews where id="& clng(id)
	conn.Execute sql
	sql="delete from article_pin where article_id="& clng(id)
end if
call CloseConn()      
response.redirect "ArticleManage.asp"
%>