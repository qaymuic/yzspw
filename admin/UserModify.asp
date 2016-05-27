<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '操作权限
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim UserID,Action,FoundErr,ErrMsg
dim rsUser,sqlUser
Action=trim(request("Action"))
UserID=trim(request("UserID"))
if UserID="" then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	call WriteErrMsg()
else
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from Userinfo where UserID=" & Clng(UserID)
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
	else
		if Action="Modify" then
			dim UserName,Password,company,address,contact,Email,tel,fax,pc,LockUser,content
			UserName=trim(request("UserName"))
			Password=trim(request("Password"))
			company=trim(request("company"))
			address=trim(request("address"))
			contact=trim(Request("contact"))
			Email=trim(request("Email"))
			tel=trim(request("tel"))
			fax=trim(request("fax"))
			pc=trim(request("pc"))
			content=replace(replace(request.form("content")," ","&nbsp;"),chr(13),"<br>")
			LockUser=trim(request("LockUser"))
			if UserName="" or strLength(UserName)>14 or strLength(UserName)<4 then
				founderr=true
				errmsg=errmsg & "<br><li>请输入用户名(不能大于14小于4)</li>"
			else
  				if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"")>0 or Instr(UserName,"$")>0 then
					errmsg=errmsg+"<br><li>用户名中含有非法字符</li>"
					founderr=true
				else
					dim sqlReg,rsReg
					sqlReg="select * from Userinfo where UserName='" & Username & "' and UserID<>" & UserID
					set rsReg=server.createobject("adodb.recordset")
					rsReg.open sqlReg,conn,1,1
					if not(rsReg.bof and rsReg.eof) then
						founderr=true
						errmsg=errmsg & "<br><li>用户名已经存在！请换一个用户名再试试！</li>"
					end if
					rsReg.Close
					set rsReg=nothing
				end if
			end if
			if Password<>"" then
				if strLength(Password)>12 or strLength(Password)<6 then
					founderr=true
					errmsg=errmsg & "<br><li>请输入密码(不能大于12小于6)。如不想修改，请留空！</li>"
				else
					if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
						errmsg=errmsg+"<br><li>密码中含有非法字符</li>"
						founderr=true
					end if
				end if
			end if
			if company="" then
				founderr=true
				errmsg=errmsg & "<br><li>申请商不能为空</li>"
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
			if fax<>"" then
				if not isnumeric(fax) then
					errmsg=errmsg & "<br><li>fax号码只能是数字，您可以选择不输入。</li>"
					founderr=true
				end if
			end if
			if pc<>"" then
				if not isnumeric(pc) or len(cstr(fax))>7 then
					errmsg=errmsg & "<br><li>fax号码只能是数字，您可以选择不输入。</li>"
					founderr=true
				end if
			end if
			if LockUser="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>用户状态不能为空！</li>"
			end if
			if FoundErr<>true then
				rsUser("UserName")=UserName
				if Password<>"" then
					rsUser("Password")=md5(Password)
				end if
				rsUser("company")=company
				rsUser("address")=address
				rsUser("contact")=contact
				rsUser("Email")=Email
				rsUser("tel")=tel
				rsUser("fax")=fax
				rsUser("pc")=pc
				rsUser("content")=content
				rsUser("LockUser")=LockUser
				rsUser.update
				rsUser.Close
				set rsUser=nothing
				call CloseConn()
				response.redirect "UserManage.asp"
			end if
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<html>
<head>
<title>修改用户信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body>
<FORM name="Form1" action="UserModify.asp" method="post">
  <table width=500 border=0 align="center" cellpadding=2 cellspacing=2 class='border'>
    <TR align=center class='title'> 
      <TD height=20 colSpan=2><font class=en><b>修改用户信息</b></font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><b>用 户 名：</b></TD>
      <TD> <INPUT name=UserName value="<%=rsUser("UserName")%>" size=10   maxLength=14> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><B>密码(至少6位)：</B></TD>
      <TD> <INPUT   type=password maxLength=16 size=10 name=Password> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>申请商：</strong></TD>
      <TD> <INPUT name="company"   type=text value="<%=rsUser("company")%>" size=40> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>地址：</strong></TD>
      <TD> <INPUT name="address"   type=text value="<%=rsUser("address")%>" size=40> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>联系人：</strong></TD>
      <TD><INPUT name="contact"   type=text id="contact" value="<%=rsUser("contact")%>" size=10> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>Email地址：</strong></TD>
      <TD> <INPUT name=Email value="<%=rsUser("Email")%>" size=30   maxLength=50> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>电话：</strong></TD>
      <TD> <INPUT   maxLength=100 size=10 name=tel value="<%=rsUser("tel")%>"></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>传真：</strong></TD>
      <TD> <INPUT name=fax value="<%=rsUser("fax")%>" size=10 maxLength=20></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>邮编：</strong></TD>
      <TD> <INPUT name=pc value="<%=rsUser("pc")%>" size=10 maxLength=50></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="120" align="right"><strong>简介：</strong></TD>
      <TD><textarea name="content" cols="40" rows="5" id="content"><%=rsUser("content")%></textarea></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>用户状态：</strong></TD>
      <TD><input type="radio" name="LockUser" value="False" <%if rsUser("LockUser")=False then response.write "checked"%>>
        正常&nbsp;&nbsp; <input type="radio" name="LockUser" value="True" <%if rsUser("LockUser")=True then response.write "checked"%>>
        锁定</TD>
    </TR>
    <TR align="center" class="tdbg" > 
      <TD height="40" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify"> 
        <input name=Submit   type=submit id="Submit" value="保存修改结果"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>">
        　　 
        <INPUT onclick="javascript:history.back(1)" style="FONT-SIZE: 9pt" type=button value=返回></TD>
    </TR>
  </TABLE>
</form>
</body>
</html>
<%
	end if
	rsUser.close
	set rsUser=nothing
end if
call CloseConn()
%>
