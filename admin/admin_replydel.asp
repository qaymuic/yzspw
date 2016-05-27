<!--#include file="conn.asp"-->
<%
dim id,sql,rs,article_id
id=replace(request("id"),"'","")
article_id=replace(request("article_id"),"'","")
'arr1=split(id,",")
'Response.write Ubound(arr1)+1
'Response.end

if id<>"" then
	sql="delete * from article_pin where id =" & clng(id)
	'response.write sql
	'response.end
	conn.Execute sql
	'sql="update ry_article set isreplay=isreplay-1 where newsid="&clng(article_id)
	'conn.execute sql
end if

set conn=nothing    
response.redirect "admin_replylist.asp?id="&article_id
%>


