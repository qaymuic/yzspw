<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '操作权限
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/md5.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from a2 order by SmallClassID asc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
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
<FORM name='myform' action='sp_save.asp' method='post'>
		
  <table width=88% border=0 align="center" cellpadding=2 cellspacing=1 bordercolor="#FFFFFF" style="border-collapse: collapse" class="border">
    <TR align=center class='title'> 
      <TD height=20 colSpan=2><font class=en><b>商铺增加</b></font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><b>商铺类别</b><BR> </TD>
      <TD width="70%"> <%
	   dim sql1,rs1
	   Set rs1=Server.CreateObject("Adodb.RecordSet")
       sql1 = "select * from Special order by Specialid"
        rs1.open sql1,conn,1,1
		%> <select name="splb" size="1">
		    <%do while not rs1.eof%>
                <option value="<%=trim(rs1("SpecialName"))%>"><%=trim(rs1("SpecialName"))%></option>
            <%
		     rs1.movenext
    	     loop
             rs1.close:set rs1=nothing
			%></select> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><B> 商铺名称</B><BR> </TD>
      <TD width="70%"> <INPUT name=spname   type=text id="spname" size=30 maxLength=12>
      <font color="#FF0000">*</font>      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong> 交易类型</strong><BR> </TD>
      <TD width="70%">        <select name="spgqlb" id="spgqlb">
        <option value="出租" selected>出租</option>
        <option value="求租">求租</option>
        <option value="转让">转让</option>
        <option value="求购">求购</option>
        <option value="出售">出售</option>
      </select>      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>地理位置</strong><BR> </TD>
      <TD width="70%"><font color="#FF0000">
              <%
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
    <TR class="tdbg" > 
      <TD width="30%"><strong>面 积</strong><BR> </TD>
      <TD width="70%"> <INPUT name="spmj"   type=text id="spmj" size=10 maxLength=20>
      平方米 <font color="#FF0000">*</font>      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>价 格</strong><BR> </TD>
      <TD width="70%"><INPUT name="spjg"   type=text id="spjg" size=10 maxLength=20>
      万<font color="#FF0000">*</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>联系方式</strong><BR> </TD>
      <TD width="70%"> <INPUT name=spcontact id="spcontact" size=30   maxLength=50> 
      <font color="#FF0000">*填联系电话、手机等</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong> 具体位置</strong><BR> </TD>
      <TD width="70%"> <INPUT name=spaddress id="spaddress" size=40   maxLength=100>
      <font color="#FF0000">*</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><p><strong> 联 系 人</strong><br> 
        </p>      </TD>
      <TD width="70%"> <INPUT name=spren id="spren" size=20 maxLength=20>
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong> 详细说明</strong><br> </TD>
      <TD width="70%"> <textarea name="spcontent" cols="40" rows="5" id="textarea"></textarea></TD>
    </TR>
    <tr class="tdbg" > 
      <td><strong> 商铺图片 </strong></td>
      <td> <iframe name="ad1" frameborder=0 width=100% height=20 scrolling=no src=../uploadface1.asp></iframe> 
        <br> <input type="text" name="Document1" size="48" class="Inpt">        </td>
    </tr>
    <tr class="tdbg" >
      <td> <strong>是否为商铺展示</strong></td>
      <td><input name="sptop1" type="checkbox" id="sptop1" value="yes"></td>
    </tr>
    <tr class="tdbg" >
      <td height="30"><strong>是否滚动展示</strong></td>
      <td><input name="sptop2" type="checkbox" id="sptop2" value="yes"></td>
    </tr>
    <tr class="tdbg" >
      <td height="22"><strong>信息有效期</strong></td>
      <td><input name="spendtime" type="text" id="spendtime" value="2004-12-31" size="10"></td>
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
