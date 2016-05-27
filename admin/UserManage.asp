<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '操作权限
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
<title>用户管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function ConfirmDel()
{
   if(confirm("确定要删除此用户吗？"))
     return true;
   else
     return false;
}
</script>
</head>
<body>
<p align="center"><font size="6">用 户 管 理</font></p>
<p align="center"><font size="2"><a href="UserAdd1.asp">增加用户</a></font></p>
<%
  	if rs.eof and rs.bof then
		response.write "目前共有 0 个注册用户"
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
        	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
        		showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
        	else
	        	currentPage=1
        		showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"个用户"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%>
  
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="2" class="border">
  <tr class="title"> 
      
    <td width="10%" height="20" align="center"><strong> 用户名</strong></td>
      
    <td width="25%" height="20" align="center"><strong> 申请商</strong></td>
    <td width="10%" height="20" align="center"><strong> 联系人</strong></td>
      
    <td width="11%" height="20" align="center"><strong> Email</strong></td>
      
    <td width="7%" height="20" align="center"><strong> 电话</strong></td>
      
    <td width="8%" height="20" align="center"><strong> 传真</strong></td>
      
    <td width="9%" height="20" align="center"><strong> 状态</strong></td>
      
    <td width="20%" height="20" align="center"><strong> 操作</strong></td>
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
	  	response.write "已锁定"
	  else
	  	response.write "正常"
	  end if
	  %></td>
      <td align="center"><a href="UserModify.asp?UserID=<%=rs("UserID")%>">修改</a>&nbsp; 
        <%if rs("LockUser")=False then %> <a href="UserLock.asp?Action=Lock&UserID=<%=rs("UserID")%>">锁定</a> 
        <%else%> <a href="UserLock.asp?Action=CancelLock&UserID=<%=rs("UserID")%>">解锁</a> 
        <%end if%>
        &nbsp;<a href="UserDel.asp?UserID=<%=rs("UserID")%>" onClick="return ConfirmDel();">删除</a></td>
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