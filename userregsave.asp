<!--#include file=inc/conn.asp -->
<!--#include file=inc/md5.asp -->
<!--#include file=inc/function.asp -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����������</title>
<link href="css/text.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.style12 {
	color: #FFFFFF;
	font-weight: bold;
	font-size: 14px;
}
-->
</style>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" class="table-body">
  <tr>
    <td><!--#include file=top.asp --><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
        <tr>
          <td bgcolor="#FFFFFF"><TABLE WIDTH=95% BORDER=0 align="center" CELLPADDING=0 CELLSPACING=0>
            <TR>
              <TD width="100%">&nbsp;</TD>
            </TR>
            <TR>
              <TD valign="top"><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="100%" height="6" valign="top"><FORM name='UserReg' action='userregsave.asp' method='post'>
                        <table width=88% height="86" border=0 align="center" cellpadding=0 cellspacing=1 bgcolor="#CCCCCC" class="border" style="border-collapse: collapse">
                          <TR align=center class='title'>
                            <TD width="100%" height=29 bgcolor="#0066CC"><span class="style12"><font>�û�ע�����д�ɹ���</font></span></TD>
                          </TR>
                          <TR align=center class='title'>
                            <TD height=22 bgcolor="#FFFFFF"><%
dim UserName,Password,PwdConfirm,Question,Answer,Email,contact,address,company,tel,fax,pc,content,founderr,errmsg
UserName=trim(request("UserName"))
Password=trim(request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
Question=trim(request("Question"))
Answer=trim(request("Answer"))
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
	errmsg=errmsg & "<li>�������û���(���ܴ���14С��4)</li>"
else
  	if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"��")>0 or Instr(UserName,"$")>0 then
		errmsg=errmsg+"<br><li>�û����к��зǷ��ַ�</li>"
		founderr=true
	end if
end if
if Password="" or strLength(Password)>12 or strLength(Password)<6 then
	founderr=true
	errmsg=errmsg & "<br><li>����������(���ܴ���12С��6)</li>"
else
	if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
		errmsg=errmsg+"<br><li>�����к��зǷ��ַ�</li>"
		founderr=true
	end if
end if
if PwdConfirm="" then
	founderr=true
	errmsg=errmsg & "<br><li>������ȷ������(���ܴ���12С��6)</li>"
else
	if Password<>PwdConfirm then
		founderr=true
		errmsg=errmsg & "<br><li>�����ȷ�����벻һ��</li>"
	end if
end if
if Question="" then
	founderr=true
	errmsg=errmsg & "<br><li>������ʾ���ⲻ��Ϊ��</li>"
end if
if Answer="" then
	founderr=true
	errmsg=errmsg & "<br><li>����𰸲���Ϊ��</li>"
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
if founderr=false then
	dim sqlReg,rsReg
	sqlReg="select * from Userinfo where UserName='" & Username & "'"
	set rsReg=server.createobject("adodb.recordset")
	rsReg.open sqlReg,conn,1,3
	if not(rsReg.bof and rsReg.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>��ע����û��Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
	else
		rsReg.addnew
		rsReg("UserName")=UserName
		rsReg("Password")=md5(Password)
		rsReg("Question")=Question
		rsReg("Answer")=md5(Answer)
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
                                <table width="320"  border="0" align="center" cellpadding="0" cellspacing="0"  bgcolor="#FFFFFF" style="border-collapse: collapse">
                                  <tr>
                                    <td  valign=top width="150"><%
if founderr=false then
	call RegSuccess()
else
	call WriteErrmsg()
end if
%>
                                    </td>
                                  </tr>
                                </table>
                                <%
call CloseConn

sub WriteErrMsg()
    response.write "<table align='center' width='300' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='left' height='15'><br><font color=#ff0000>�������µ�ԭ����ע���û���</font></td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'>" & errmsg & "<p align='center'>��<a href='javascript:onclick=history.go(-1)'>�� ��</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub

sub RegSuccess()
    response.write "<table align='center' width='300' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>�ɹ�ע���û���</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>��ע����û�����" & UserName & "<p align='center'>��<a href='sppost.asp'>������������</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub
%></TD>
                          </TR>
                        </TABLE>
                        <div align="center"> </div>
                    </form></td>
                  </tr>
              </table></TD>
            </TR>
            <TR>
              <TD>&nbsp;</TD>
            </TR>
          </TABLE></td>
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