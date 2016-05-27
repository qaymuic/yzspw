<!--#include file=conn.asp-->
<!--#include file=../inc/md5.asp-->
<%
dim sql
dim rs
dim username
dim password
username=replace(trim(request("username")),"'","")
password=replace(trim(Request("password")),"'","")
password=md5(password)
set rs=server.createobject("adodb.recordset")
sql="select * from admin where password='"&password&"' and username='"&username&"'"
rs.open sql,conn,1,1
if not(rs.bof and rs.eof) then
	if password=rs("password") then
		session("admin")=rs("username")
		session("realname")=rs("realname")
		session("purview")=rs("purview")
		rs.close
		set rs=nothing
		call CloseConn()
		Response.Redirect "manage.asp"
	end if
end if
rs.close
set rs=nothing
call CloseConn()
%>
<html>
<head>
<link rel='stylesheet' href='style.css'>
</head>
<body>
<br><br><br>
<table align='center' width='300' border='0' cellpadding='4' cellspacing='0' class='border'>
  <tr >
    <td height='15' colspan='2' align="center" class='title'>操作: 确认身份失败!</td>
  </tr>
  <tr>
    <td height='23' colspan='2' align="center" class='tdbg'> <br>
      <br>
      用户名或密码错误！！！<br>
      <br> <a href='javascript:onclick=history.go(-1)'>【返回】</a> <br>
      <br></td>
  </tr>
</table>
</body>
</html>

