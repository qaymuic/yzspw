<!--#include file=inc/conn.asp-->
<!--#include file=inc/md5.asp-->
<%
dim sql
dim rs
dim username
dim password
username=replace(trim(request("username")),"'","")
password=replace(trim(Request("password")),"'","")
password=md5(password)
set rs=server.createobject("adodb.recordset")
sql="select * from Userinfo where password='"&password&"' and username='"&username&"' and LockUser=false"
rs.open sql,conn,1,1
if not(rs.bof and rs.eof) then
	if password=rs("password") then
		session("username")=rs("username")
		session("company")=rs("company")
		session("contact")=rs("contact")
		session("address")=rs("address")
		session("tel")=rs("tel")
		session("fax")=rs("fax")
		session("pc")=rs("pc")
		session("email")=rs("email")
		rs.close
		set rs=nothing
		call CloseConn()
		Response.Redirect "sppost.asp"
	end if
end if
rs.close
set rs=nothing
call CloseConn()
%>
<script>
alert("用户名或密码错误！");history.back(-1);

</script>