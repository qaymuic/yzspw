<!--
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/text.css" rel="stylesheet" type="text/css">
//-->
<table width="95%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <form name="ssform1" method="post" action="searchlist.asp">
    <tr>
      <td bgcolor="#336699"><img src="images/wzsousuo.gif" width="159" height="26"></td>
    </tr>
    <tr>
      <td bgcolor="#FFFBE8" class="td-tianchong-4px">关键字
          <input name="title" type="text" size="10">
      </td>
    </tr>
    <tr>
      <td align="center" bgcolor="#FFFBE8" class="td-tianchong-4px"><%
	   Set rs=Server.CreateObject("Adodb.RecordSet")
           sql = "select * from bigclass order by BigClassID"
           rs.open sql,conn,1,1
		%>
          <select name="splb" size="1">
            <option value="" selected>全部</option>
            <%do while not rs.eof%>
            <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
            <%
		     rs.movenext
    	     loop
             rs.close:set rs=nothing
			%>
      </select></td>
    </tr>
    <tr>
      <td align="right" bgcolor="#FFFBE8" class="td-tianchong-4px"><input type="submit" name="Submit" value="搜 索">
      </td>
    </tr>
  </form>
</table>
<br>
<table width="95%"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td bgcolor="#CC6600"><img src="images/redian.gif" width="159" height="26"></td>
  </tr>
  <%
		   dim rs1,sql1
	       Set rs1=Server.CreateObject("Adodb.RecordSet")
           sql1 = "select top 10 * from ytiinews order by hits desc"
           rs1.open sql1,conn,1,1
		   do while not rs1.eof
		   %>
  <tr>
    <td bgcolor="#DFEBF0" class="td_text_001">·<a href="list.asp?id=<%=rs1("id")%>" target="_blank"><%=trim(rs1("title"))%></a></td>
  </tr>
  <%
		     rs1.movenext
    	     loop
             rs1.close:set rs1=nothing
			%>
</table>

