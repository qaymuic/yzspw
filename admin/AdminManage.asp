<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=1    '操作权限
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<%
dim rs, sql, strPurview
Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from admin order by id"
rs.Open sql,conn,1,1
%>
<html>
<head>
<title>管理员管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function ConfirmDel()
{
   if(confirm("确定要删除此管理员吗？"))
     return true;
   else
     return false;
}
</script>
</head>
<body>
  <p align="center"><font size="6">管 理 员 管 理</font><br>
<br>
    <strong>管理选项：</strong><a href="AdminAdd.asp">新增管理员</a></p>
  
<table border="0" align="center" cellpadding="2" cellspacing="2" class="border">
  <tr align="center" class="title"> 
    <td  width="30" height="20"><strong> 序号</strong></td>
    <td  width="76" height="20"><strong>真实姓名</strong></td>
    <td  width="63"><strong>用户名</strong></td>
    <td  width="94" height="20"><strong> 密码</strong></td>
    <td  width="85" height="20"><strong> 权限</strong></td>
    <td  width="100" height="20"><strong> 操作</strong></td>
  </tr>
  <%while not rs.EOF %>
  <tr align="center" class="tdbg"> 
    <td width="30" height="25"><%=rs("ID")%></td>
    <td width="76" height="25"><%=rs("realname")%></td>
    <td width="63" height="25"><%=rs("username")%></td>
    <td width="94" height="25">********</td>
    <td width="85" height="25"> 
      <%
		  select case rs("purview")
		    case 1
              strPurview="管理员"
            case 2
              strpurview="部门主管"
            case 3
			  strpurview="一般用户"
		  end select
		  response.write(strPurview)
         %>
    </td>
    <td width="100" height="25"><a href="AdminModify.asp?ID=<%=rs("ID")%>">修改</a>&nbsp;&nbsp;<a href="AdminDel.asp?ID=<%=rs("ID")%>" onClick="return ConfirmDel();">删除</a></td>
  </tr>
  <%
     rs.MoveNext
   Wend
  %>
</table>
</body>
</html>
<%
rs.Close
set rs=Nothing
call CloseConn()
%>