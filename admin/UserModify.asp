<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '����Ȩ��
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
	ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	call WriteErrMsg()
else
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from Userinfo where UserID=" & Clng(UserID)
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
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
				errmsg=errmsg & "<br><li>�������û���(���ܴ���14С��4)</li>"
			else
  				if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"��")>0 or Instr(UserName,"$")>0 then
					errmsg=errmsg+"<br><li>�û����к��зǷ��ַ�</li>"
					founderr=true
				else
					dim sqlReg,rsReg
					sqlReg="select * from Userinfo where UserName='" & Username & "' and UserID<>" & UserID
					set rsReg=server.createobject("adodb.recordset")
					rsReg.open sqlReg,conn,1,1
					if not(rsReg.bof and rsReg.eof) then
						founderr=true
						errmsg=errmsg & "<br><li>�û����Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
					end if
					rsReg.Close
					set rsReg=nothing
				end if
			end if
			if Password<>"" then
				if strLength(Password)>12 or strLength(Password)<6 then
					founderr=true
					errmsg=errmsg & "<br><li>����������(���ܴ���12С��6)���粻���޸ģ������գ�</li>"
				else
					if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
						errmsg=errmsg+"<br><li>�����к��зǷ��ַ�</li>"
						founderr=true
					end if
				end if
			end if
			if company="" then
				founderr=true
				errmsg=errmsg & "<br><li>�����̲���Ϊ��</li>"
			end if
			if contact="" then
				founderr=true
				errmsg=errmsg & "<br><li>��ϵ�˲���Ϊ��</li>"
			end if
			if Email="" then
				founderr=true
				errmsg=errmsg & "<br><li>Email����Ϊ��</li>"
			else
				if IsValidEmail(Email)=false then
					errmsg=errmsg & "<br><li>����Email�д���</li>"
			   		founderr=true
				end if
			end if
			if fax<>"" then
				if not isnumeric(fax) then
					errmsg=errmsg & "<br><li>fax����ֻ�������֣�������ѡ�����롣</li>"
					founderr=true
				end if
			end if
			if pc<>"" then
				if not isnumeric(pc) or len(cstr(fax))>7 then
					errmsg=errmsg & "<br><li>fax����ֻ�������֣�������ѡ�����롣</li>"
					founderr=true
				end if
			end if
			if LockUser="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�û�״̬����Ϊ�գ�</li>"
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
<title>�޸��û���Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body>
<FORM name="Form1" action="UserModify.asp" method="post">
  <table width=500 border=0 align="center" cellpadding=2 cellspacing=2 class='border'>
    <TR align=center class='title'> 
      <TD height=20 colSpan=2><font class=en><b>�޸��û���Ϣ</b></font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><b>�� �� ����</b></TD>
      <TD> <INPUT name=UserName value="<%=rsUser("UserName")%>" size=10   maxLength=14> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><B>����(����6λ)��</B></TD>
      <TD> <INPUT   type=password maxLength=16 size=10 name=Password> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>�����̣�</strong></TD>
      <TD> <INPUT name="company"   type=text value="<%=rsUser("company")%>" size=40> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>��ַ��</strong></TD>
      <TD> <INPUT name="address"   type=text value="<%=rsUser("address")%>" size=40> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>��ϵ�ˣ�</strong></TD>
      <TD><INPUT name="contact"   type=text id="contact" value="<%=rsUser("contact")%>" size=10> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>Email��ַ��</strong></TD>
      <TD> <INPUT name=Email value="<%=rsUser("Email")%>" size=30   maxLength=50> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>�绰��</strong></TD>
      <TD> <INPUT   maxLength=100 size=10 name=tel value="<%=rsUser("tel")%>"></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>���棺</strong></TD>
      <TD> <INPUT name=fax value="<%=rsUser("fax")%>" size=10 maxLength=20></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>�ʱࣺ</strong></TD>
      <TD> <INPUT name=pc value="<%=rsUser("pc")%>" size=10 maxLength=50></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="120" align="right"><strong>��飺</strong></TD>
      <TD><textarea name="content" cols="40" rows="5" id="content"><%=rsUser("content")%></textarea></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="120" align="right"><strong>�û�״̬��</strong></TD>
      <TD><input type="radio" name="LockUser" value="False" <%if rsUser("LockUser")=False then response.write "checked"%>>
        ����&nbsp;&nbsp; <input type="radio" name="LockUser" value="True" <%if rsUser("LockUser")=True then response.write "checked"%>>
        ����</TD>
    </TR>
    <TR align="center" class="tdbg" > 
      <TD height="40" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify"> 
        <input name=Submit   type=submit id="Submit" value="�����޸Ľ��"> <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>">
        ���� 
        <INPUT onclick="javascript:history.back(1)" style="FONT-SIZE: 9pt" type=button value=����></TD>
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
