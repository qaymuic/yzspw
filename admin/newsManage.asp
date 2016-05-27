<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
%>
<%
if session("admin")="" then
	response.write "请登录！"
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
sql="select * from ytiinews where BigClassName='商铺动态'"
if Title<>"" then
sql=sql & " and Title like '%" & Title & "%' "
end if
sql=sql & " order by id desc"
rs.Open sql,conn,1,1
%>
<html>
<head>
<title>内容管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function ConfirmDel()
{
   if(confirm("确定要删除此内容吗？"))
     return true;
   else
     return false;
}
</script>
</head>
<body leftmargin="0" topmargin="0">
<br>
<p align="center">
  <strong>管理选项：</strong><a href="newsAdd.asp">新增商铺动态</a></p>
<table width="95%" class="border" align="center">
  <tr class="tdbg" align="center"><form name="form3" method="get" action="newsManage.asp">
            <td height="30"> <strong>查找内容：</strong> 
              <input name="username" type="text" class=smallInput id="username" size="28">
<input name="Query" type="submit" id="Query" value="查 询">
        &nbsp;&nbsp;请输入标题关键字，查找相应内容。</td>
          </form></tr></table>
  <br>
  <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr >
      <td width="95%" align=right><%
  	if rs.eof and rs.bof then
		response.write "共找到 0 个内容</td></tr></table>"
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
		response.Write "共找到 <font color=red>" & totalPut & " </font>个内容"
%></td>
    </tr>
  </table>
  
  <%		
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,false,"个内容"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage

            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,false,"个内容"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,false,"个内容"
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
    <td  width="67" align="center"><strong>状态</strong></td>
    <td  width="294" height="22" align="center"><strong>内容标题</strong></td>
    <td  width="58" height="22" align="center"><strong> 点击数</strong></td>
    <td  width="68" height="22" align="center"><strong> 发布时间</strong></td>
    <td  width="140" height="22" align="center"><strong> 操作</strong></td>
  </tr>
  <%do while not rs.EOF %>
  <tr class="tdbg"> 
    <td width="67" height="25"><%if rs("istop") then%>不显示<%else%>已显示<%end if%></td>
    <td width="294"><a href="/list.asp?id=<%=rs("id")%>" target=_blank><%=rs("Title")%></a></td>
    <td width="58" height="25" align="center"><%=rs("hits")%></td>
    <td width="68" height="25" align="center"> <%=rs("updatetime")%> </td>
    <td width="140" height="25" align="center"><a href="newsModify.asp?ID=<%=rs("ID")%>">修改</a>&nbsp;&nbsp;<a href="ArticleDel.asp?ID=<%=rs("ID")%>" onClick="return ConfirmDel();">删除</a></td>
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