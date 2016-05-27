<!--#include file=inc/conn.asp -->
<!--#include file=inc/function.asp -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>扬州商铺网</title>
<link href="css/text.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.style2 {
	font-size: 14px;
	font-weight: bold;
	color: #FFFFFF;
}
.style3 {font-size: 14px}
.style4 {color: #000000}
.style5 {font-size: 14px; font-weight: bold; color: #000000; }
-->
</style>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" class="table-body">
  <tr>
    <td><!--#include file=top.asp --><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table width="776"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="180" valign="top"><!--#include file=wzleft.asp --></td>
              <td width="596" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="5">
                <tr>
                      <td>&nbsp;<a href="index.asp">返回首页</a>&nbsp;&nbsp;&gt;&gt;&gt;&nbsp;&nbsp;搜索结果</td>
                </tr>
              </table>
                
                <table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
                 	<%
dim strFileName,title,splb
title=ReplaceBadChar(request("title"))
splb=ReplaceBadChar(request("splb"))
const MaxPerPage=25
dim totalPut,CurrentPage,TotalPages,UserName
dim i,j
strFileName="searchlist.asp?title="&title&"&splb="&splb
dim strPurview
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from ytiinews where 1=1"
if title="" and splb="" then
sql=sql
else
 if title<>"" then
 sql=sql&" and title like '%"&title&"%'"
 end if
 if splb<>"" then
 sql=sql&" and BigClassName = '"&splb&"'"
 end if
end if
sql=sql & " order by id desc"

rs.Open sql,conn,1,1
%>
				  <tr>
                    <td width="100%" height="6" valign="top" bgcolor="#FFFFFF">
                      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="25" class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                  <td bgcolor="#eeeeee" class="td-tianchong-4px"><strong>搜索结果列表： <%if title<>"" then Response.write "关键字："&title end if%><%if splb<>"" then Response.write "，分类："&splb end if%>
                                    </strong></td>
                            </tr>
                          </table></td>
                        </tr>
                      </table>
                      <%		
	    totalPut=rs.recordcount
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"条"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage

            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"条"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"条"
	    	end if
		end if
%>
                      <br>
                      <%  
sub showContent
   	dim i
    i=0
%>
                      <TABLE width=95% border=0 align="center" cellPadding=0 cellSpacing=0 style="FONT-SIZE: 12px">
                        <TBODY>
                          <%do while not rs.eof%>
                          <TR vAlign=center>
                            <TD colSpan=3 height=6></TD>
                          </TR>
                          <TR>
                            <TD width=5%><DIV align=center><IMG height=7 src="images/jt.gif" width=7></DIV></TD>
                            <TD vAlign=center width=76%><A  href="list.asp?id=<%=rs("id")%>" target=_blank><%=gotTopic(rs("title"),70)%></A>　 </TD>
                            <TD vAlign=center width=19%><font color=#808080><%=rs("UpdateTime")%></font></TD>
                          </TR>
                          <TR vAlign=center>
                            <TD colSpan=3 height=6></TD>
                          </TR>
                          <TR vAlign=center>
                            <TD colSpan=3><IMG height=2 src="images/point2.gif" width=545></TD>
                          </TR>
                          <% 
		i=i+1
	    if i>=MaxPerPage then exit do
	rs.movenext   
	loop
%>
                        </TBODY>
                      </TABLE>
                      <%
   end sub 
%></td>
                  </tr>
                </table></td>
            </tr>
          </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
      <tr>
        <td align="center"><!--#include file=foot.asp --></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<%
rs.close
closeconn
%>