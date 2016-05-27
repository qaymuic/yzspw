<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '操作权限
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../Inc/md5.asp"-->
<%
dim UserName,Password,PwdConfirm,company,address,contact,Email,fax,tel,pc,founderr,content,errmsg
UserName=trim(request("UserName"))
Password=trim(request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
company=trim(request("company"))
address=trim(request("address"))
contact=trim(Request("contact"))
Email=trim(request("Email"))
fax=trim(request("fax"))
tel=trim(request("tel"))
pc=trim(request("pc"))
content=replace(replace(request.form("content")," ","&nbsp;"),chr(13),"<br>")
if UserName="" or strLength(UserName)>14 or strLength(UserName)<4 then
	founderr=true
	errmsg=errmsg & "<br><li>请输入用户名(不能大于14小于4)</li>"
else
  	if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"")>0 or Instr(UserName,"$")>0 then
		errmsg=errmsg+"<br><li>用户名中含有非法字符</li>"
		founderr=true
	end if
end if
if Password="" or strLength(Password)>12 or strLength(Password)<6 then
	founderr=true
	errmsg=errmsg & "<br><li>请输入密码(不能大于12小于6)</li>"
else
	if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
		errmsg=errmsg+"<br><li>密码中含有非法字符</li>"
		founderr=true
	end if
end if
if PwdConfirm="" then
	founderr=true
	errmsg=errmsg & "<br><li>请输入确认密码(不能大于12小于6)</li>"
else
	if Password<>PwdConfirm then
		founderr=true
		errmsg=errmsg & "<br><li>密码和确认密码不一致</li>"
	end if
end if
if company="" then
	founderr=true
	errmsg=errmsg & "<br><li>申请商不能为空</li>"
end if
if address="" then
	founderr=true
	errmsg=errmsg & "<br><li>地址不能为空</li>"
end if
if contact="" then
	founderr=true
	errmsg=errmsg & "<br><li>联系人不能为空</li>"
end if
if Email="" then
	founderr=true
	errmsg=errmsg & "<br><li>Email不能为空</li>"
else
	if IsValidEmail(Email)=false then
		errmsg=errmsg & "<br><li>您的Email有错误</li>"
   		founderr=true
	end if
end if
if tel<>"" then
	if not isnumeric(tel) then
		errmsg=errmsg & "<br><li>电话号码只能是数字，您可以选择不输入。</li>"
		founderr=true
	end if
end if
if fax<>"" then
	if not isnumeric(fax) then
		errmsg=errmsg & "<br><li>传真号码只能是数字，您可以选择不输入。</li>"
		founderr=true
	end if
end if

if founderr=false then
	dim sqlReg,rsReg
	sqlReg="select * from Userinfo where UserName='" & Username & "'"
	set rsReg=server.createobject("adodb.recordset")
	rsReg.open sqlReg,conn,1,3
	if not(rsReg.bof and rsReg.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>你加入的用户已经存在！请换一个用户名再试试！</li>"
	else
		rsReg.addnew
		rsReg("UserName")=UserName
		rsReg("Password")=md5(Password)
		rsReg("company")=company
		rsReg("address")=address
		rsReg("contact")=contact
		rsReg("Email")=Email
		rsReg("tel")=tel
		rsReg("fax")=fax
		rsReg("pc")=pc
		rsReg("content")=content
		rsReg.update
		founderr=false
	end if
	rsReg.close
	set rsReg=nothing
end if		
%>
<html>
<head>
<title>完成加入</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
</head>

<body>
<table width="600" height="300" border="0" align="center" cellpadding="0" cellspacing="0"  bgcolor="#FFFFFF" style="border-collapse: collapse">
  <tr> 
    <td width="600" align="center" valign=top> 
      <%
if founderr=false then
	call RegSuccess()
else
	call WriteErrmsg()
end if
%>
      <table border="1" align=center cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" height="20" bordercolorlight="#000000" bordercolordark="#FFFFFF" width="120">
        <tr> 
          <td height="20" width="119" bgcolor="#CCCCCC"><b>加入成功,<a href=usermanage.asp>请返回</a></b></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<%
call CloseConn

sub WriteErrMsg()
    response.write "<table align='center' width='300' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>由于以下的原因不能加入用户！</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>" & errmsg & "<p align='center'>【<a href='javascript:onclick=history.go(-1)'>返 回</a>】<br></p></td></tr>"
	response.write "</table>" 
end sub

sub RegSuccess()
    response.write "<table align='center' width='300' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>成功加入用户！</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>你加入的用户名：" & UserName & "<p align='center'>【<a href='javascript:onclick=window.close()'>关 闭</a>】<br></p></td></tr>"
	response.write "</table>" 
end sub
%>