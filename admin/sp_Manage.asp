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
strFileName="sp_manage.asp"

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from spw order by ID desc"
rs.Open sql,conn,1,1
%>
<html>
<head>
<title>����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ������"))
     return true;
   else
     return false;
}
</script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</head>
<body>
<br>
<p align="center"><font size="4"> �� �� �� Ϣ �� ��</font></p>
<center><a href='sp_add.asp'>��������</a></center>
<p align="center">
  <%
  	if rs.eof and rs.bof then
		response.write "Ŀǰ���� 0 ������<br><br>"
		response.write "<a href='sp_add.asp'>��������</a>"
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
        	showpage strFileName,totalput,MaxPerPage,true,true,"��"
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"��"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
        		showpage strFileName,totalput,MaxPerPage,true,true,"��"
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"��"
        	else
	        	currentPage=1
        		showpage strFileName,totalput,MaxPerPage,true,true,"��"
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"��"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%>
</p>
  <%do while not rs.EOF %>
<table width="97%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" class="border">
  <tr bgcolor="#FFFFFF" class="title"> 
    <td width="11%" height="20" align="center"> �������</td>
    <td width="12%" align="center"> �������� </td>
    <td width="11%" height="20" align="center"> ��������</td>
    <td width="20%" height="20" align="center"> ����λ��  </td>
    <td width="11%" height="20" align="center"> �� �� </td>
    <td width="10%" height="20" align="center">�� �� </td>
    <td width="5%" height="20" align="center"> ���</td>
    <td width="10%" align="center">ʱ��</td>
    <td width="10%" height="20" align="center">����</td>
  </tr>

  <tr bgcolor="#FFFFFF" class="tdbg">
    <td align="center"><%=rs("splb")%></td>
    <td align="center"><%=rs("spname")%></td>
    <td align="center"><%=rs("spgqlb")%></td>
    <td align="center"><%=rs("spaddress")%></td>
    <td align="center"><%=rs("spmj")%></td>
    <td align="center"><%=rs("spjg")%></td>
    <td align="center"><%=rs("sphits")%></td>
    <td align="center"><%=rs("spaddtime")%></td>
    <td align="center"><a href="sp_modify.asp?id=<%=rs("id")%>">�޸�</a> <a href="sp_del.asp?id=<%=rs("id")%>" onClick="return ConfirmDel();">ɾ��</a></td>
  </tr></table>
<br>
  <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>

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