<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '����Ȩ��
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim strFileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim rs, sql
strFileName="UserManage.asp"

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from Userinfo where mtype=0 order by UserID desc"
rs.Open sql,conn,1,1
%>
<html>
<head>
<title>�û�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ�����û���"))
     return true;
   else
     return false;
}
</script>
</head>
<body>
<p align="center"><font size="6">�� �� �� ��</font></p>
<p align="center"><font size="2"><a href="UserAdd1.asp">�����û�</a></font></p>
<%
  	if rs.eof and rs.bof then
		response.write "Ŀǰ���� 0 ��ע���û�"
	else
    	totalPut=rs.recordcount
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if

    	end if
	    if currentPage=1 then
        	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
        		showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
        	else
	        	currentPage=1
        		showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"���û�"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%>
  
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="2" class="border">
  <tr class="title"> 
      
    <td width="10%" height="20" align="center"><strong> �û���</strong></td>
      
    <td width="25%" height="20" align="center"><strong> ������</strong></td>
    <td width="10%" height="20" align="center"><strong> ��ϵ��</strong></td>
      
    <td width="11%" height="20" align="center"><strong> Email</strong></td>
      
    <td width="7%" height="20" align="center"><strong> �绰</strong></td>
      
    <td width="8%" height="20" align="center"><strong> ����</strong></td>
      
    <td width="9%" height="20" align="center"><strong> ״̬</strong></td>
      
    <td width="20%" height="20" align="center"><strong> ����</strong></td>
    </tr>
<%do while not rs.EOF %>
    <tr class="tdbg"> 
      <td align="center"><%=rs("username")%></td>
      <td align="center"><%=rs("company")%></td>
      <td align="center"><%=rs("contact")%></td>
      <td><a href='mailto:<%=rs("Email")%>'><%=rs("Email")%></a></td>
      <td><%=rs("tel")%></td>
      <td align="center"><%=rs("fax")%></td>
      <td align="center"><%
	  if rs("LockUser")=true then
	  	response.write "������"
	  else
	  	response.write "����"
	  end if
	  %></td>
      <td align="center"><a href="UserModify.asp?UserID=<%=rs("UserID")%>">�޸�</a>&nbsp; 
        <%if rs("LockUser")=False then %> <a href="UserLock.asp?Action=Lock&UserID=<%=rs("UserID")%>">����</a> 
        <%else%> <a href="UserLock.asp?Action=CancelLock&UserID=<%=rs("UserID")%>">����</a> 
        <%end if%>
        &nbsp;<a href="UserDel.asp?UserID=<%=rs("UserID")%>" onClick="return ConfirmDel();">ɾ��</a></td>
    </tr>
    <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
  </table>  
<%
end sub 
%>
</body>
</html>
<%
rs.Close
set rs=Nothing
call CloseConn()  
%>