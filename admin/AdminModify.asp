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
dim Action,UserID, password,PwdConfirm, purview, rs, sql,FoundErr,ErrMsg,realname
Action=trim(request("Action"))
UserID=trim(Request("ID"))
realname=trim(Request("realname"))
password=trim(Request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
purview=trim(Request("purview"))
if UserID="" then
	response.Redirect("AdminManage.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from Admin where ID="&UserID&"",conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>不存在此用户！</li>"
else
	if Action="Modify" then
		if rs("UserName")="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>用户名不能为空！</li>"
		end if
		if rs("realname")="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>真实姓名不能为空！</li>"
		end if
		if PwdConfirm<>Password then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>确认密码必须与新密码相同！</li>"
		end if
		if Purview="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>用户权限不能为空！</li>"
		end if
		if FoundErr<>True then
			if Password<>"" then
				rs("password")=md5(password)
			end if
	     	rs("purview")=Cint(purview)  
			rs("realname")=realname
    	 	rs.update
			rs.Close
	     	set rs=Nothing
    	 	call CloseConn()
     		Response.Redirect "AdminManage.asp"
		end if
	end if
	purview=rs("purview")
end if
if FoundErr=True then
	Call WriteErrMsg()
else
%>
<html>
<head>
<title>修改管理员信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<body>
<form method="post" action="AdminModify.asp" name="form1">
    
  <table width="400" border="0" align="center" cellpadding="2" cellspacing="2" class="border" >
    <tr class="title"> 
      <td height="20" colspan="2"> <div align="center"><font size="2"><strong>修改管理员信息</strong></font></div></td>
    </tr>
    <tr> 
      <td align="right" class="tdbg"><strong>用 户 名：</strong></td>
      <td class="tdbg"><%=rs("UserName")%> <input name="ID" type="hidden" value="<%=rs("ID")%>"></td>
    </tr>
    <tr>
      <td align="right" class="tdbg"><strong>真实姓名：</strong></td>
      <td class="tdbg"><input name="realname" type="text" id="realname" value="<%=rs("realname")%>"></td>
    </tr>
    <tr> 
      <td align="right" class="tdbg"><strong>新 密 码：</strong></td>
      <td class="tdbg"><input type="password" name="Password"> <font color="#0000FF">如果不想修改，请保持为空</font></td>
    </tr>
    <tr> 
      <td align="right" class="tdbg"><strong>确认密码：</strong></td>
      <td class="tdbg"><input type="password" name="PwdConfirm"> <font color="#0000FF">如果不想修改，请保持为空</font></td>
    </tr>
    <tr> 
      <td align="right" class="tdbg"><strong>权限设置： </strong></td>
      <td class="tdbg"><select name="purview" id="purview">
          <option value="1" <%if purview=1 then %>selected<% end if %>>管理员</option>
          <option value="2" <%if purview=2 then %>selected<% end if %>>部门主管</option>
          <option value="3" <%if purview=3 then %>selected<% end if %>>一般用户</option>
        </select></td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="Modify"> 
        <input  type="submit" name="Submit" value=" 确 定 "></td>
    </tr>
  </table>
  </form>
</body>
</html>
<%
end if
rs.close
set rs=nothing
call CloseConn()
%>
