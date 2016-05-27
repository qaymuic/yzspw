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
dim Action,BigClassName,Admin,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
Admin=trim(request("Admin"))
if Action="Add" then
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章大类名不能为空！</li>"
	end if
	if Admin="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>管理员不能为空！</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From a1 Where BigClassName='" & BigClassName & "'",conn,1,3
		if not (rs.bof and rs.EOF) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>文章大类“" & BigClassName & "”已经存在！</li>"
		else
    	 	rs.addnew
     		rs("BigClassName")=BigClassName
			rs("Admin")=Admin
    	 	rs.update
     		rs.Close
	     	set rs=Nothing
    	 	call CloseConn()
			Response.Redirect "ClassManage1.asp"  
		end if
	end if
end if
if FoundErr=True then
	call WriteErrMsg()
else
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
  if (document.form1.Admin.value=="")
  {
    alert("管理员不能为空！");
    document.form1.Admin.focus();
    return false;
  }
}
</script>
</head>
<body>
<form name="form1" method="post" action="ClassAddBig1.asp" onsubmit="return checkBig()">
  <table width="400" border="0" align="center" cellpadding="2" cellspacing="2" class="border">
    <tr class="title"> 
      <td height="20" colspan="2" align="center"><strong>添加大类</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="150"><strong>大类名称：</strong></td>
      <td><input name="BigClassName" type="text" size="30" maxlength="20"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="150"><strong>管理员：</strong><br>
      </td>
      <td><input name="Admin" type="text" id="Admin" value="admin" size="30" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="Add">
        <input name="Add" type="submit" value=" 添 加 "></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
end if
call CloseConn()
%>
