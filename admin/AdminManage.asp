<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=1    '����Ȩ��
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
<title>����Ա����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ���˹���Ա��"))
     return true;
   else
     return false;
}
</script>
</head>
<body>
  <p align="center"><font size="6">�� �� Ա �� ��</font><br>
<br>
    <strong>����ѡ�</strong><a href="AdminAdd.asp">��������Ա</a></p>
  
<table border="0" align="center" cellpadding="2" cellspacing="2" class="border">
  <tr align="center" class="title"> 
    <td  width="30" height="20"><strong> ���</strong></td>
    <td  width="76" height="20"><strong>��ʵ����</strong></td>
    <td  width="63"><strong>�û���</strong></td>
    <td  width="94" height="20"><strong> ����</strong></td>
    <td  width="85" height="20"><strong> Ȩ��</strong></td>
    <td  width="100" height="20"><strong> ����</strong></td>
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
              strPurview="����Ա"
            case 2
              strpurview="��������"
            case 3
			  strpurview="һ���û�"
		  end select
		  response.write(strPurview)
         %>
    </td>
    <td width="100" height="25"><a href="AdminModify.asp?ID=<%=rs("ID")%>">�޸�</a>&nbsp;&nbsp;<a href="AdminDel.asp?ID=<%=rs("ID")%>" onClick="return ConfirmDel();">ɾ��</a></td>
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