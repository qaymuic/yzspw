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
dim BigClassID,Action,rs,NewBigClassName,OldBigClassName,Admin,FoundErr,ErrMsg
BigClassID=trim(Request("BigClassID"))
Action=trim(Request("Action"))
NewBigClassName=trim(Request("NewBigClassName"))
OldBigClassName=trim(Request("OldBigClassName"))
Admin=trim(request("Admin"))

if BigClassID="" then
  response.Redirect("ClassManage1.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from a1 where BigClassID=" & CLng(BigClassID),conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�����´��಻���ڣ�</li>"
else
	if Action="Modify" then
		if NewBigClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>���´���������Ϊ�գ�</li>"
		end if
		if Admin="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����Ա����Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			rs("BigClassName")=NewBigClassName
			rs("Admin")=Admin
    	 	rs.update
			rs.Close
	     	set rs=Nothing
			if NewBigClassName<>OldBigClassName then
				conn.execute "Update a2 set BigClassName='" & NewBigClassName & "' where BigClassName='" & OldBigClassName & "'"
				conn.execute "Update spw set BigClassName='" & NewBigClassName & "' where BigClassName='" & OldBigClassName & "'"
     		end if		
			call CloseConn()
     		Response.Redirect "ClassManage1.asp"
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<html>
<head>
<title>�޸Ĵ�������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
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
<form name="form1" method="post" action="ClassModifyBig1.asp">
  <table width="400" border="0" align="center" cellpadding="5" cellspacing="0" class="border">
    <tr class="title"> 
      <td height="20" colspan="2" align="center"><strong>�޸Ĵ�������</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="150"><strong>����ID��</strong></td>
      <td><%=rs("BigClassID")%> <input name="BigClassID" type="hidden" id="BigClassID" value="<%=rs("BigClassID")%>"> 
        <input name="OldBigClassName" type="hidden" id="OldBigClassName" value="<%=rs("BigClassName")%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="150"><strong>�������ƣ�</strong></td>
      <td><input name="NewBigClassName" type="text" id="NewBigClassName" value="<%=rs("BigClassName")%>" size="30" maxlength="20"></td>
    </tr>
    <tr class="tdbg"> 
      <td width="150"><strong>����Ա��</strong></td>
      <td><input name="Admin" type="text" id="Admin" value="<%=rs("Admin")%>" size="30" maxlength="100"></td>
    </tr>
    <tr class="tdbg"> 
      <td colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="Modify">
        <input name="Save" type="submit" id="Save" value=" �� �� "></td>
    </tr>
  </table>  
  </form>
</body>
</html>
<%
	end if
end if
rs.close
set rs=nothing
call CloseConn()
%>
