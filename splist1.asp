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
.style1 {color: #990000}
.style2 {
	font-size: 14px;
	font-weight: bold;
	color: #FFFFFF;
}
.style3 {font-size: 14px}
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
              <td width="180" valign="top"><table width="89%"  border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                <tr>
                  <td bgcolor="#0066CC"><div align="center" class="style2">商铺类型</div></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF"><div align="center"><a href="splist1.asp?spgqlb=&#36716;&#35753;">商铺转让</a></div></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF"><div align="center"><a href="splist1.asp?spgqlb=&#20986;&#31199;">商铺出租</a></div></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF"> <div align="center"><a href="splist1.asp?spgqlb=&#27714;&#31199;">商铺求租</a></div></td>
                </tr>
              </table>
                <br>
                <table width="89%"  border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                  <tr>
                    <td bgcolor="#0066CC"><div align="center" class="style2">商铺类型</div></td>
                  </tr>
				  <%
		   dim rs1,sql1
	       Set rs1=Server.CreateObject("Adodb.RecordSet")
           sql1 = "select * from Special order by Specialid"
           rs1.open sql1,conn,1,1
		   do while not rs1.eof
		   %>
            <tr>
                    <td bgcolor="#FFFFFF"><div align="center"><a href="splist1.asp?splb=<%=rs1("SpecialName")%>"><%=trim(rs1("SpecialName"))%></a></div></td>
            </tr><%
		     rs1.movenext
    	     loop
             rs1.close:set rs1=nothing
			%>
 <%
					dim strFileName,splb,spgqlb
					splb=ReplaceBadChar(request("splb"))
					spgqlb=ReplaceBadChar(request("spgqlb"))
					const MaxPerPage=18
					dim totalPut,CurrentPage,TotalPages,UserName
					dim i,j
					strFileName="splist1.asp?splb="&splb&"&spgqlb="&spgqlb
					dim strPurview
					if request("page")<>"" then
						currentPage=cint(request("page"))
					else
						currentPage=1
					end if
					Set rs=Server.CreateObject("Adodb.RecordSet")
					sql="select * from spw where sptop1=false and spgqlb in ('求租','出租') "
					if splb="" and spgqlb="" then
					 sql=sql
					else
					  if splb<>"" then
					   sql=sql&" and splb='"&splb&"'"
					  end if
					  if spgqlb<>"" then
					   sql=sql&" and spgqlb='"&spgqlb&"'"
					  end if
					end if
					sql=sql & " order by id desc"
					rs.Open sql,conn,1,1
					  %>
					  
                </table></td>
              <td width="596" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="5">
                <tr>
                  <td>&nbsp;<a href="index.asp">返回首页</a>&nbsp;&gt;&gt;&gt;&nbsp;商铺租赁<%if splb<>"" then response.write "&nbsp;&gt;&gt;&gt;&nbsp;"&splb%></td>
                </tr>
              </table>
                <table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
                <tr>
                  <td width="100%" height="6" valign="top" bgcolor="#FFFFFF"><table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
					 <tr >
                        <td width="95%" align=right><%
  	if rs.eof and rs.bof then
		response.write "共找到 0 个商铺</td></tr></table>"
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
		
%>
                        </td>
                      </tr>
                    </table>
                      <%		
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,false,"个商铺"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage

            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,false,"个商铺"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,false,"个商铺"
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
                      <TABLE width=100% border=0 align="center" cellPadding=0 cellSpacing=0 style="FONT-SIZE: 12px">
                        <TBODY>
                          <%for j=0 to 5 %>
                          <TR vAlign=center>
                            <%for i=0 to 2%>
                            <TD><table width="100%" border="0" align="center" cellpadding="3" cellspacing="2">
                          <tr>
                                    <td width="75" valign="top"><a href="sqdetails.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("spphoto")%>" border="1" class="img-75_85"></a></td>
                            <td height="80" valign="top" bgcolor="#F5F5F5" class="td-tianchong-4px"><b><%=rs("spname")%></b><br>
                              <span class="style1">类型</span>:<%=rs("spgqlb")%><br>
                              <span class="style1">价格</span>:<%=rs("spjg")%>万/<%=rs("spmj")%>M<sup>2</sup><br>
                              <span class="style1">位置</span>:<%=rs("SmallClassName")%><br>
                                    &nbsp;&nbsp;&nbsp;&nbsp; <a href="sqdetails.asp?id=<%=rs("id")%>" target=_blank class=a_color_001>详细>>></a></td>
                          </tr>
                        </table></TD>
                            <%
				rs.movenext
				if rs.eof then
				exit for
				end if
				next
				%>
                          </tr>
                          <%
				if rs.eof then
				exit for
				end if
				next
				%>
                        </TBODY>
                      </TABLE>
                      <%
   end sub 
%>
                  </td>
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