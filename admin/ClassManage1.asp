<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '操作权限
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<%
dim sql,rsBigClass,rsSmallClass,ErrMsg
set rsBigClass=server.CreateObject("adodb.recordset")
rsBigClass.open "Select * From a1",conn,1,3
%>
<html>
<head>
<title>栏目管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("大类名称不能为空！");
    document.form1.BigClassName.focus();
    return false;
  }
}
function checkSmall()
{
  if (document.form2.BigClassName.value=="")
  {
    alert("请先添加大类名称！");
	document.form1.BigClassName.focus();
	return false;
  }

  if (document.form2.SmallClassName.value=="")
  {
    alert("小类名称不能为空！");
	document.form2.SmallClassName.focus();
	return false;
  }
}
function ConfirmDelBig()
{
   if(confirm("确定要删除此文章大类吗？删除此大类同时将删除所包含的小类，并且不能恢复！"))
     return true;
   else
     return false;
	 
}

function ConfirmDelSmall()
{
   if(confirm("确定要删除此文章小类吗？一旦删除将不能恢复！"))
     return true;
   else
     return false;
	 
}
</script>

</head>
<body>
<div align="center"> 
  <p><font size="6"> 栏 目 管 理</font><br><br>
    管理选项：<a href="ClassAddBig1.asp">添加管理大类</a></p>
  <table width="500" border="0" cellpadding="2" cellspacing="2" class="border">
    <tr class="title"> 
      <td width="150" height="20" align="center"><strong>栏目名称</strong></td>
      <td align="center"><strong>管理员</strong></td>
      <td height="20" align="center"><strong>操作选项</strong></td>
    </tr>
    <%
	do while not rsBigClass.eof
%>
    <tr class="tdbg"> 
      <td width="150"><img src="../Images/tree_folder4.gif" width="15" height="15"><%=rsBigClass("BigClassName")%></td>
      <td align="center">
	  <%
	  if rsBigClass("Admin")&""<>"" then
	  	response.write rsBigClass("Admin")
	  else
	  	response.write "&nbsp;"
	  end if
	  %>
	  </td>
      <td align="center"><a href="ClassAddSmall1.asp?BigClassName=<%=rsBigClass("BigClassName")%>">添加子栏目</a> | <a href="ClassModifyBig1.asp?BigClassID=<%=rsBigClass("BigClassID")%>">修改</a> 
        | <a href="ClassDelBig1.asp?BigClassName=<%=rsBigClass("BigClassName")%>" onClick="return ConfirmDelBig();">删除</a></td>
    </tr>
    <%
	  set rsSmallClass=server.CreateObject("adodb.recordset")
	  rsSmallClass.open "Select * From a2 Where BigClassName='" & rsBigClass("BigClassName") & "'",conn,1,3
	  if not(rsSmallClass.bof and rsSmallClass.eof) then
		do while not rsSmallClass.eof
	%>
    <tr class="tdbg"> 
      <td width="150">&nbsp;&nbsp;<img src="../Images/tree_folder3.gif" width="15" height="15"><%=rsSmallClass("SmallClassName")%></td>
      <td align="center">
	  <%
	  if rsSmallClass("Admin")<>"" then
	  	response.write rsSmallClass("Admin")
	  else
	  	response.write "&nbsp;"
	  end if
	  %>
	  </td>
      <td align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <a href="ClassModifySmall1.asp?SmallClassID=<%=rsSmallClass("SmallClassID")%>">修改</a> 
        | <a href="ClassDelSmall1.asp?SmallClassID=<%=rsSmallClass("SmallClassID")%>" onClick="return ConfirmDelSmall();">删除</a></td>
    </tr>
    <%
			rsSmallClass.movenext
		loop
	  end if
	  rsSmallClass.close
	  set rsSmallClass=nothing	
	  rsBigClass.movenext
	loop
%>
  </table>
</div>
</body>
</html>
<%
rsBigClass.close
set rsBigClass=nothing
call CloseConn()
%>
