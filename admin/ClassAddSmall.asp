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
dim Action,BigClassName,SmallClassName,Admin,rs,FoundErr,ErrMsg
Action=trim(Request("Action"))
BigClassName=trim(request("BigClassName"))
SmallClassName=trim(request("SmallClassName"))
Admin=trim(request("Admin"))
if Action="Add" then
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>���´���������Ϊ�գ�</li>"
	end if
	if SmallClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����С��������Ϊ�գ�</li>"
	end if
	if Admin="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����Ա����Ϊ�գ�</li>"
	end if
	if FoundErr<>True then
		Set rs=Server.CreateObject("Adodb.RecordSet")
		rs.open "Select * From SmallClass Where BigClassName='" & BigClassName & "' AND SmallClassName='" & SmallClassName & "'",conn,1,3
		if not rs.EOF then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��" & BigClassName & "�����Ѿ���������С�ࡰ" & SmallClassName & "����</li>"
		else
     		rs.addnew
			rs("BigClassName")=BigClassName
    	 	rs("SmallClassName")=SmallClassName
			rs("Admin")=Admin
     		rs.update
	     	rs.Close
    	 	set rs=Nothing
     		call CloseConn()
			Response.Redirect "ClassManage.asp"  
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
function checkSmall()
{
  if (document.form2.BigClassName.value=="")
  {
    alert("������Ӵ������ƣ�");
	document.form1.BigClassName.focus();
	return false;
  }

  if (document.form2.SmallClassName.value=="")
  {
    alert("С�����Ʋ���Ϊ�գ�");
	document.form2.SmallClassName.focus();
	return false;
  }
  if (document.form2.Admin.value=="")
  {
    alert("����Ա����Ϊ�գ�");
	document.form2.Admin.focus();
	return false;
  }
}
</script>

</head>
<body>
<form name="form2" method="post" action="ClassAddSmall.asp" onsubmit="return checkSmall()">
  <table width="400" border="0" align="center" cellpadding="2" cellspacing="2" class="border">
    <tr class="title"> 
      <td height="20" colspan="2" align="center"><strong>���С��</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="150"><strong>�������ࣺ</strong></td>
      <td> <select name="BigClassName">
          <%
	dim rsBigClass
	set rsBigClass=server.CreateObject("adodb.recordset")
	rsBigClass.open "Select * From BigClass",conn,1,1
	if rsBigClass.bof and rsBigClass.bof then
		response.write "<option>����������´���</option>"
	else
		do while not rsBigClass.eof
			if rsBigClass("BigClassName")=BigClassName then
				response.write "<option value='"& rsBigClass("BigClassName") & "' selected>" & rsBigClass("BigClassName") & "</option>"
			else
				response.write "<option value='"& rsBigClass("BigClassName") & "'>" & rsBigClass("BigClassName") & "</option>"
			end if
			rsBigClass.movenext
		loop
	end if
	rsBigClass.close
	set rsBigClass=nothing
	%>
        </select></td>
    </tr>
    <tr class="tdbg"> 
      <td width="150"><strong>С�����ƣ�</strong></td>
      <td><input name="SmallClassName" type="text" size="30" maxlength="20"></td>
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
