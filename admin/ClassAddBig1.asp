<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2    '����Ȩ��
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
		ErrMsg=ErrMsg & "<br><li>���´���������Ϊ�գ�</li>"
	end if
	if Admin="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����Ա����Ϊ�գ�</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From a1 Where BigClassName='" & BigClassName & "'",conn,1,3
		if not (rs.bof and rs.EOF) then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>���´��ࡰ" & BigClassName & "���Ѿ����ڣ�</li>"
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
<title>��Ŀ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("�������Ʋ���Ϊ�գ�");
    document.form1.BigClassName.focus();
    return false;
  }
  if (document.form1.Admin.value=="")
  {
    alert("����Ա����Ϊ�գ�");
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
      <td height="20" colspan="2" align="center"><strong>��Ӵ���</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="150"><strong>�������ƣ�</strong></td>
      <td><input name="BigClassName" type="text" size="30" maxlength="20"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="150"><strong>����Ա��</strong><br>
      </td>
      <td><input name="Admin" type="text" id="Admin" value="admin" size="30" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="Add">
        <input name="Add" type="submit" value=" �� �� "></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
end if
call CloseConn()
%>
