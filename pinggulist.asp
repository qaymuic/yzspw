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
                  <td>&nbsp;<a href="index.asp">返回首页</a>&nbsp;&gt;&gt;&gt;&nbsp;商铺评估</td><td align="right" width="150">>> <a href="ADDPINGGU.ASP" target="_blank"><font color="#FF0000"><b>我要评估</b></font></a> <<&nbsp;&nbsp;</td>
                </tr>
              </table>
                <table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
                	<%
dim strFileName
const MaxPerPage=25
dim totalPut,CurrentPage,TotalPages,UserName
dim i,j
strFileName="pinggulist.asp"
dim strPurview
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from ytiinews where BigClassName='商铺评估'"
sql=sql & " order by id desc"
rs.Open sql,conn,1,1
%>
			 
				<tr>
                  <td width="100%" valign="top" bgcolor="#FFFFFF"> <%
		   dim rs2,sql2
	       Set rs2=Server.CreateObject("Adodb.RecordSet")
           sql2 = "select top 1 * from ytiinews where BigClassName='商铺评估' and istop=true order by id desc"
           rs2.open sql2,conn,1,1
		   if not rs2.eof and not rs2.bof then
		   %>
                    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td align="left" bgcolor="#F5F5F5" class="td-tianchong-4px"><%=rs2("title")%></td>
                          </tr>
                        </table></td>
                      </tr>
                    </table>                    
                    <table width="576"  border="0" cellspacing="0" cellpadding="5">
                    <tr>
                      <td width="191" align="center" valign="top"><%if rs2("DefaultPicUrl")<>"" then%><a href="list.asp?id=<%=rs2("id")%>" target="_blank"><img src="<%=rs2("DefaultPicUrl")%>" width=191 class="img-border-1px"></a><%else%>暂无图片<%end if%></td>
                      <td width="365" height="160" align="left" valign="top" class="TD-MENU"><%=gottopic(nohtml(rs2("content")),500)%></td>
                    </tr>
                    <tr>
                      <td bgcolor="#F5F5F5"> </td>
                      <td height="40" align="right" bgcolor="#F5F5F5" class="td-tianchong-4px"><a href="list.asp?id=<%=rs2("id")%>" target="_blank">详细内容&gt;&gt;&gt;</a>&nbsp;&nbsp;&nbsp;&nbsp;</td>
                    </tr>
                  </table>
                    <%
				  end if
				  set rs2=nothing
				  %></td>
                </tr>
              </table>
                <br>
                <table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
                  <tr>
                    <td width="100%" height="6" valign="top" bgcolor="#FFFFFF">
                      <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="25" class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td bgcolor="#eeeeee" class="td-tianchong-4px"><strong>商铺评估列表： </strong></td>
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