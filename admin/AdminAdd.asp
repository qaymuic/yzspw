<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=1    '����Ȩ��
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
		ErrMsg=ErrMsg & "<br><li>�û�������Ϊ�գ�</li>"
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ʼ���벻��Ϊ�գ�</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>ȷ������������ʼ������ͬ��</li>"
	end if
	if purview="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�û�Ȩ�޲���Ϊ�գ�</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.Open "Select * from Admin where username='"&username&"'",conn,1,3
		if not (rs.bof and rs.EOF) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>���ݿ����Ѿ����ڴ˹���Ա��</li>"
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
<title>��������Ա</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function check()
{
  if(document.form1.username.value=="")
    {
      alert("�û�������Ϊ�գ�");
	  document.form1.username.focus();
      return false;
    }
    
  if(document.form1.Password.value=="")
    {
      alert("���벻��Ϊ�գ�");
	  document.form1.Password.focus();
      return false;
    }
    
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("��ʼ������ȷ�����벻ͬ��");
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
      <td height="20" colspan="2"> <div align="center"><strong>��������Ա</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" class="tdbg"> �� �� ����</td>
      <td class="tdbg"><input name="username" type="text"> &nbsp;</td>
    </tr>
    <tr class="tdbg">
      <td align="right" class="tdbg"> ��ʵ������</td>
      <td class="tdbg"><input name="realname" type="text" id="realname"></td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" class="tdbg"> ��ʼ���룺 </td>
      <td class="tdbg"><font size="2"> 
        <input type="password" name="Password">
        </font></td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" class="tdbg"> ȷ�����룺</td>
      <td class="tdbg"><font size="2"> 
        <input type="password" name="PwdConfirm">
        </font></td>
    </tr>
    <tr class="tdbg"> 
      <td align="right" class="tdbg">Ȩ�����ã� </td>
      <td class="tdbg"><select name="purview" id="purview">
          <option value="1" selected>����Ա</option>
          <option value="2">��������</option>
          <option value="3">һ���û�</option>
        </select></td>
    </tr>
    <tr align="center" class="tdbg"> 
      <td height="40" colspan="2" class="tdbg"><input name="Action" type="hidden" id="Action" value="Add"> 
        <input  type="submit" name="Submit" value=" �� �� "></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
end if
Call CloseConn()
%>
