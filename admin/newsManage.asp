<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
%>
<%
if session("admin")="" then
	response.write "���¼��"
	response.end
end if
%>
<!--#include file="conn.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim strFileName
const MaxPerPage=25
dim totalPut,CurrentPage,TotalPages,UserName
dim i,j
strFileName="newsmanage.asp"
dim rs, sql, strPurview
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
dim Title
Title=trim(request("username"))
Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from ytiinews where BigClassName='���̶�̬'"
if Title<>"" then
sql=sql & " and Title like '%" & Title & "%' "
end if
sql=sql & " order by id desc"
rs.Open sql,conn,1,1
%>
<html>
<head>
<title>���ݹ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ����������"))
     return true;
   else
     return false;
}
</script>
</head>
<body leftmargin="0" topmargin="0">
<br>
<p align="center">
  <strong>����ѡ�</strong><a href="newsAdd.asp">�������̶�̬</a></p>
<table width="95%" class="border" align="center">
  <tr class="tdbg" align="center"><form name="form3" method="get" action="newsManage.asp">
            <td height="30"> <strong>�������ݣ�</strong> 
              <input name="username" type="text" class=smallInput id="username" size="28">
<input name="Query" type="submit" id="Query" value="�� ѯ">
        &nbsp;&nbsp;���������ؼ��֣�������Ӧ���ݡ�</td>
          </form></tr></table>
  <br>
  <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr >
      <td width="95%" align=right><%
  	if rs.eof and rs.bof then
		response.write "���ҵ� 0 ������</td></tr></table>"
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
		response.Write "���ҵ� <font color=red>" & totalPut & " </font>������"
%></td>
    </tr>
  </table>
  
  <%		
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,false,"������"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage

            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,false,"������"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,false,"������"
	    	end if
		end if
	end if
%>
  <br>
  
<%  
sub showContent
   	dim i
    i=0
%> 

<table width="95%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
  <tr class="title"> 
    <td  width="67" align="center"><strong>״̬</strong></td>
    <td  width="294" height="22" align="center"><strong>���ݱ���</strong></td>
    <td  width="58" height="22" align="center"><strong> �����</strong></td>
    <td  width="68" height="22" align="center"><strong> ����ʱ��</strong></td>
    <td  width="140" height="22" align="center"><strong> ����</strong></td>
  </tr>
  <%do while not rs.EOF %>
  <tr class="tdbg"> 
    <td width="67" height="25"><%if rs("istop") then%>����ʾ<%else%>����ʾ<%end if%></td>
    <td width="294"><a href="/list.asp?id=<%=rs("id")%>" target=_blank><%=rs("Title")%></a></td>
    <td width="58" height="25" align="center"><%=rs("hits")%></td>
    <td width="68" height="25" align="center"> <%=rs("updatetime")%> </td>
    <td width="140" height="25" align="center"><a href="newsModify.asp?ID=<%=rs("ID")%>">�޸�</a>&nbsp;&nbsp;<a href="ArticleDel.asp?ID=<%=rs("ID")%>" onClick="return ConfirmDel();">ɾ��</a></td>
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