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
response.write "<script>alert('������Ч���뷵�أ�');history.back(-1);</script>"
response.end
else
	id=clng(request("id"))
end if
if AuthorName="" or content="" then
response.write "<script>alert('�����е���Ϣ����Ϊ�գ��뷵�أ�');history.back(-1);</script>"
response.end
end if
sql="insert into article_pin (article_id,article_title,name,title,content,user_ip) values ("&id&",'"&atitle&"','"&AuthorName&"','"&title&"','"&content&"','"&logip&"')"
set rs=conn.execute(sql)
conn.close
set conn=nothing
%>
<script>
alert('���Ļظ����ύ�ɹ����ǳ���л���Ĳ��룬ף���и������飡');location.href='list.asp?id=<%=id%>';
</script>