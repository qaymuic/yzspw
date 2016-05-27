<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=1    '操作权限
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/md5.asp"-->
<%
dim username, password,PwdConfirm,purview,realname
dim rs, sql
dim Action,FoundErr,ErrMsg
Action=trim(request("Action"))
username=trim(Request("username"))
realname=trim(Request("realname"))
password=trim(Request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
purview=trim(Request("purview"))
if Action="Add" then
	if username="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>用户名不能为空！</li>"
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>初始密码不能为空！</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>确认密码必须与初始密码相同！</li>"
	end if
	if purview="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>用户权限不能为空！</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.Open "Select * from Admin where username='"&username&"'",conn,1,3
		if not (rs.bof and rs.EOF) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>数据库中已经存在此管理员！</li>"
		else
			password=md5(password)
	     	rs.addnew
    	 	rs("username")=username
			rs("realname")=realname
	     	rs("password")=password
    	 	rs("purview")=Cint(purview)
	     	rs.update
    	 	rs.Close
	     	set rs=Nothing
			Call CloseConn()
			response.Redirect "AdminManage.asp"
		end if
	end if
end if
if FoundErr=True then
	Call WriteErrMsg()
else
%>
<html>
<head>
<title>新增管理员</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function check()
{
  if(document.form1.username.value=="")
    {
      alert("用户名不能为空！");
	  document.form1.username.focus();
      return false;
    }
    
  if(document.form1.Password.value=="")
    {
      alert("密码不能为空！");
	  document.form1.Password.focus();
      return false;
    }
    
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("初始密码与确认密码不同！");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
}
</script>
</head>
<body>
<form method="post" action="AdminAdd.asp" name="form1" onsubmit="javascript:return check();">
  <table width="300" border="0" align="center" cellpadding="2" cellspacing="2" class="border" >
    <tr class="title"> 
      <td height="20" colspan="2"> <div align="center"><strong>新增管理员</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" class="tdbg"> 用 户 名：</td>
      <td class="tdbg"><input name="username" type="text"> &nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td align="right" class="tdbg"> 真实姓名：</td>
      <td class="tdbg"><input name="realname" type="text" id="realname"></td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" class="tdbg"> 初始密码： </td>
      <td class="tdbg"><font size="2"> 
        <input type="password" name="Password">
        </font></td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" class="tdbg"> 确认密码：</td>
      <td class="tdbg"><font size="2"> 
        <input type="password" name="PwdConfirm">
        </font></td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" class="tdbg">权限设置： </td>
      <td class="tdbg"><select name="purview" id="purview">
          <option value="1" selected>管理员</option>
          <option value="2">部门主管</option>
          <option value="3">一般用户</option>
        </select></td>
    </tr>
    <tr align="center" class="tdbg"> 
      <td height="40" colspan="2" class="tdbg"><input name="Action" type="hidden" id="Action" value="Add"> 
        <input  type="submit" name="Submit" value=" 添 加 "></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
end if
Call CloseConn()
%>
