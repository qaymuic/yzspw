<!--#include file=inc/conn.asp -->
<!--#include file=inc/function.asp -->
<%
	if session("username")="" then
		response.write  "<script>alert('会员请先登陆，或免费注册再发布！');window.close()</script>"
		response.end
	end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>扬州商铺网</title>
<link href="css/text.css" rel="stylesheet" type="text/css">
<%
dim rs2
dim sql2
dim count
set rs2=server.createobject("adodb.recordset")
sql2 = "select * from a2 order by SmallClassID asc"
rs2.open sql2,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs2.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs2("SmallClassName"))%>","<%= trim(rs2("BigClassName"))%>","<%= trim(rs2("SmallClassName"))%>");
        <%
        count = count + 1
        rs2.movenext
        loop
        rs2.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassName.options[document.myform.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    


</script>

<style type="text/css">
<!--
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
              <td width="180" valign="top"><!--#include file=wzleft.asp --></td>
              <td width="596" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="5">
                <tr>
                  <td>&nbsp;<a href="index.asp">返回首页</a>&nbsp;&gt;&gt;&gt;&nbsp;发布商铺</td>
                </tr>
              </table>
                <table width=98% border=0 align="center" cellpadding=2 cellspacing=1 bordercolor="#FFFFFF" bgcolor="#CCCCCC" class="border" style="border-collapse: collapse">
                  <FORM name='myform' action='spsave.asp' method='post'>
				  <TR align=right bgcolor="#FFFFFF" class='title'>
                    <TD height=20 colSpan=2>请认真填写下列表单，以便您的商铺及时反馈，带<font color="#FF0000">*</font>为必填项</TD>
                  </TR>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center"><span class="style3">商铺类别<BR>                    
                      </span></TD>
                    <TD width="70%"><%
	   Set rs1=Server.CreateObject("Adodb.RecordSet")
       sql1 = "select * from Special order by Specialid"
        rs1.open sql1,conn,1,1
		%>
                        <select name="splb" size="1">
                          <%do while not rs1.eof%>
                          <option value="<%=trim(rs1("SpecialName"))%>"><%=trim(rs1("SpecialName"))%></option>
                          <%
		     rs1.movenext
    	     loop
             rs1.close:set rs1=nothing
			%>
                        </select>
                    </TD>
                  </TR>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center"><span class="style3"> 商铺名称<BR>                    
                      </span></TD>
                    <TD width="70%"><INPUT name=spname   type=text id="spname" size=30 maxLength=12>
                        <font color="#FF0000">*</font> </TD>
                  </TR>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center"><span class="style3"> 交易类型<BR>                    
                      </span></TD>
                    <TD width="70%"><select name="spgqlb" id="spgqlb">
                        <option value="出租" selected>出租</option>
                        <option value="求租">求租</option>
                        <option value="转让">转让</option>
                        <option value="求购">求购</option>
                        <option value="出售">出售</option>
                      </select>
                    </TD>
                  </TR>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center"><span class="style3">地理位置<BR>                    
                      </span></TD>
                    <TD width="70%"><font color="#FF0000">
              <%
        set rs=server.createobject("adodb.recordset")
        sql = "select * from a1"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加地区。"
		else
		%> <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
			dim selclass
		    selclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
              </select> <select name="SmallClassName">
               <option value="" selected>不选择地区</option>
                <%
			sql="select * from a2 where BigClassName='" & selclass & "'"
			rs.open sql,conn,1,1
			if not(rs.eof and rs.bof) then
			%>
                <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <% rs.movenext
				do while not rs.eof%>
                <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
              </select></font></TD>
                  </TR>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center"><span class="style3">面 积<BR>                    
                      </span></TD>
                    <TD width="70%"><INPUT name="spmj"   type=text id="spmj" size=10 maxLength=20>
      平方米 <font color="#FF0000">*</font> </TD>
                  </TR>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center"><span class="style3">价 格<BR>                    
                      </span></TD>
                    <TD width="70%"><INPUT name="spjg"   type=text id="spjg" size=10 maxLength=20>
      万<font color="#FF0000">*</font></TD>
                  </TR>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center"><span class="style3">联系方式<BR>                    
                      </span></TD>
                    <TD width="70%"><INPUT name=spcontact id="spcontact" value="<%=session("tel")%>" size=30   maxLength=50>
                        <font color="#FF0000">*填联系电话、手机等</font></TD>
                  </TR>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center"><span class="style3"> 具体位置<BR>                    
                      </span></TD>
                    <TD width="70%"><INPUT name=spaddress id="spaddress" value="<%=session("address")%>" size=40   maxLength=100>
                        <font color="#FF0000">*</font></TD>
                  </TR>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center"><p class="style3"> 联 系 人<br>
                      </p></TD>
                    <TD width="70%"><INPUT name=spren id="spren" value="<%=session("contact")%>" size=20 maxLength=20>
                    </TD>
                  </TR>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center"><span class="style3"> 详细说明<br>                    
                      </span></TD>
                    <TD width="70%"><textarea name="spcontent" cols="40" rows="5" id="textarea"></textarea></TD>
                  </TR>
                  <tr bgcolor="#FFFFFF" class="tdbg" >
                    <td align="center"><span class="style3"> 商铺图片 </span></td>
                    <td><iframe name="ad1" frameborder=0 width=100% height=20 scrolling=no src=uploadface1.asp></iframe>
                        <br>
                        <input type="text" name="Document1" size="48" class="Inpt">
                    </td>
                  </tr>
                  <tr bgcolor="#FFFFFF" class="tdbg" >
                    <td height="22" align="center"><span class="style3">信息有效期</span></td>
                    <td><input name="spendtime" type="text" id="spendtime" value="2004-12-31" size="10"></td>
                  </tr>
                  <TR bgcolor="#FFFFFF" class="tdbg" >
                    <TD width="30%" align="center">&nbsp;</TD>
                    <TD><input   type=submit value=" 保 存 " name=Submit>
&nbsp;
      <input name=Reset   type="button" id="Reset2" value=" 返 回 " onclick='javascript:history.back(-1)'>
                    </TD>
                  </TR></form>
                </TABLE></td>
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
closeconn
%>