<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '操作权限
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
	dim sql2,rs2,id
	id=request("id")
	sql2="select * from spw where id="&id
	set rs2=server.createobject("adodb.recordset")
	rs2.open sql2,conn,1,1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language = "JavaScript">
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from a2 order by SmallClassID asc"
rs.open sql,conn,1,1
%>
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
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
<title>商铺增加</title>
</head>

<body>
<FORM name='myform' action='sp_modisave.asp' method='post'>
		
  <table width=88% border=0 align="center" cellpadding=2 cellspacing=1 bordercolor="#FFFFFF" style="border-collapse: collapse" class="border">
    <TR align=center class='title'> 
      <TD height=20 colSpan=2><font class=en><b>商铺增加</b></font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><b>商铺类别</b><BR> </TD>
      <TD width="70%"> <%
	   dim sql1,rs1
	   Set rs1=Server.CreateObject("Adodb.RecordSet")
       sql1 = "select * from Special order by Specialid desc"
        rs1.open sql1,conn,1,1
		%> <select name="splb" size="1">
		    <%do while not rs1.eof%>
                <option value="<%=trim(rs1("SpecialName"))%>" <%if rs2("splb")=rs1("SpecialName") then Response.write "selected"%>><%=trim(rs1("SpecialName"))%></option>
            <%
		     rs1.movenext
    	     loop
             rs1.close:set rs1=nothing
			%></select> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><B> 商铺名称</B><BR> </TD>
      <TD width="70%"> <INPUT name=spname   type=text id="spname" value="<%=rs2("spname")%>" size=30 maxLength=12>
      <font color="#FF0000">*</font>      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong> 交易类型</strong><BR> </TD>
      <TD width="70%">        <select name="spgqlb" id="spgqlb">
        <option value="出租" <% if rs2("spgqlb")="出租" then Response.write "selected"%>>出租</option>
        <option value="求租" <% if rs2("spgqlb")="求租" then Response.write "selected"%>>求租</option>
        <option value="转让" <% if rs2("spgqlb")="转让" then Response.write "selected"%>>转让</option>
        <option value="求购" <% if rs2("spgqlb")="求购" then Response.write "selected"%>>求购</option>
        <option value="出售" <% if rs2("spgqlb")="出售" then Response.write "selected"%>>出售</option>
      </select>      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>地理位置</strong><BR> </TD>
      <TD width="70%"><font color="#FF0000"> 
              <%
	if session("purview")=3 or session("purview")=4 then
		response.write rs2("BigClassName") & "<input name='BigClassName' type='hidden' value='" & rs2("BigClassName") & "'>&gt;&gt;"
	else		
        sql = "select * from a1"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
		else
		%>
              <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
                <%
		    do while not rs.eof
			%>
                <option <% if rs("BigClassName")=rs2("BigClassName") then response.Write("selected") end if%> value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
              </select>
              <%
	end if
	if session("purview")=4 then
		response.write rs2("SmallClassName") & "<input name='SmallClassName' type='hidden' value='" & rs2("SmallClassName") & "'>"
	else
	%>
              <select name="SmallClassName">
                <option value="" <%if rs2("SmallClassName")="" then response.write "selected"%>>不选择地区</option>
                <%
			sql="select * from a2 where BigClassName='" & rs2("BigClassName") & "'"
			rs.open sql,conn,1,1
			if not(rs.eof and rs.bof) then
				do while not rs.eof%>
                <option <% if rs("SmallClassName")=rs2("SmallClassName") then response.Write("selected") end if%> value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
              </select>
              <%
	end if
	%>
              </font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>面 积</strong><BR> </TD>
      <TD width="70%"> <INPUT name="spmj"   type=text id="spmj" value="<%=rs2("spmj")%>" size=10 maxLength=20>
      平方米 <font color="#FF0000">*</font>      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>价 格</strong><BR> </TD>
      <TD width="70%"><INPUT name="spjg"   type=text id="spjg" value="<%=rs2("spjg")%>" size=10 maxLength=20>
      万<font color="#FF0000">*</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>联系方式</strong><BR> </TD>
      <TD width="70%"> <INPUT name=spcontact id="spcontact" value="<%=rs2("spcontact")%>" size=30   maxLength=50> 
      <font color="#FF0000">*填联系电话、手机等</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong> 具体位置</strong><BR> </TD>
      <TD width="70%"> <INPUT name=spaddress id="spaddress" value="<%=rs2("spaddress")%>" size=40   maxLength=100>
      <font color="#FF0000">*</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><p><strong> 联 系 人</strong><br> 
        </p>      </TD>
      <TD width="70%"> <INPUT name=spren id="spren" value="<%=rs2("spren")%>" size=20 maxLength=20>
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong> 详细说明</strong><br> </TD>
      <TD width="70%"> <textarea name="spcontent" cols="40" rows="5" id="textarea"><%=rs2("spcontent")%></textarea></TD>
    </TR>
    <tr class="tdbg" > 
      <td><strong> 商铺图片 </strong></td>
      <td> <iframe name="ad1" frameborder=0 width=100% height=20 scrolling=no src=../uploadface1.asp></iframe> 
        <br> <input name="Document1" type="text" class="Inpt" value="<%=rs2("spphoto")%>" size="48">        </td>
    </tr>
    <tr class="tdbg" >
      <td> <strong>是否为商铺展示</strong></td>
      <td><input name="sptop1" type="checkbox" id="sptop1" value="yes" <% if rs2("sptop1")=true then Response.write "checked"%>></td>
    </tr>
    <tr class="tdbg" >
      <td height="30"><strong>是否滚动展示</strong></td>
      <td><input name="sptop2" type="checkbox" id="sptop2" value="yes" <% if rs2("sptop2")=true then Response.write "checked"%>></td>
    </tr>
    <tr class="tdbg" >
      <td height="22"><strong>信息有效期</strong></td>
      <td><input name="spendtime" type="text" id="spendtime" value="<%=rs2("spendtime")%>" size="10">
      <input name="id" type="hidden" id="id" value="<%=id%>"></td>
    </tr>
    <TR class="tdbg" > 
      <TD width="30%">&nbsp;</TD>
      <TD><input   type=submit value=" 保 存 " name=Submit> &nbsp; <input name=Reset   type="button" id="Reset2" value=" 返 回 " onclick='javascript:history.back(-1)'> 
      </TD>
    </TR>
  </TABLE>
	
  <div align="center"> </div>
</form>
</body>
</html>
<%
set rs2=nothing
closeconn
%>