<!--#include file="inc/conn.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim id,title,atitle,AuthorName,content,logip,sql,rs
logip=Request.ServerVariables("REMOTE_HOST")
AuthorName=dvHTMLEncode(trim(request.form("AuthorName")))
title=dvHTMLEncode(trim(request.form("title")))
atitle=dvHTMLEncode(trim(request.form("atitle")))
content=dvHTMLEncode(request.form("content"))
id=trim(request.form("id"))
if request("id")="" or not isnumeric(request("id")) then
response.write "<script>alert('参数无效，请返回！');history.back(-1);</script>"
response.end
else
	id=clng(request("id"))
end if
if AuthorName="" or content="" then
response.write "<script>alert('您所有的信息不能为空，请返回！');history.back(-1);</script>"
response.end
end if
sql="insert into article_pin (article_id,article_title,name,title,content,user_ip) values ("&id&",'"&atitle&"','"&AuthorName&"','"&title&"','"&content&"','"&logip&"')"
set rs=conn.execute(sql)
conn.close
set conn=nothing
%>
<script>
alert('您的回复已提交成功，非常感谢您的参与，祝您有个好心情！');location.href='list.asp?id=<%=id%>';
</script>