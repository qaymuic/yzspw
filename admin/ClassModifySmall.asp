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
dim SmallClassID,Action,BigClassName, SmallClassName, OldSmallClassName,rs,Admin,FoundErr,ErrMsg
SmallClassID=trim(Request("SmallClassID"))
Action=trim(Request("Action"))
BigClassName=trim(Request.form("BigClassName"))
SmallClassName=trim(Request.form("SmallClassName"))
OldSmallClassName=trim(request.form("OldSmallClassName"))
Admin=trim(request("Admin"))
if SmallClassID="" then
	response.Redirect("ClassManage.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from SmallClass where SmallClassID="&SmallClassID&"",conn,1,3
if rs.Bof or rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>������С�಻���ڣ�</li>"
else
	if Action="Modify" then
		if SmallClassName="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����С��������Ϊ�գ�</li>"
		end if
		if Admin="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>����Ա����Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			rs("SmallClassName")=SmallClassName
			rs("Admin")=Admin
     		rs.update
			rs.Close
    	 	set rs=Nothing
			if SmallClassName<>OldSmallClassName then
				conn.execute "Update Article set SmallClassName='" & SmallClassName & "' where BigClassName='" & BigClassName & "' and SmallClassName='" & OldSmallClassName & "'"
	     	end if	
			call CloseConn()
    	 	Response.Redirect "ClassManage.asp"
		end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%>
<html>
<head>
<title>�޸�С������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
</head>

<body>
<form name="form1" method="post" action="ClassModifySmall.asp">
  <p>&nbsp;</p>
  <table width="400" border="0" align="center" cellpadding="5" cellspacing="0" class="border">
    <tr class="title"> 
      <td height="20" colspan="2" align="center"><strong>�޸�С������</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td><strong>�������ࣺ</strong></td>
      <td><%=rs("BigClassName")%> <input name="SmallClassID" type="hidden" id="SmallClassID" value="<%=rs("SmallClassID")%>"> 
        <input name="OldSmallClassName" type="hidden" id="OldSmallClassName" value="<%=rs("SmallClassName")%>"> 
        <input name="BigClassName" type="hidden" id="BigClassName" value="<%=rs("BigClassName")%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td><strong>С�����ƣ�</strong></td>
      <td><input name="SmallClassName" type="text" id="SmallClassName" value="<%=rs("SmallClassName")%>" size="30" maxlength="20"></td>
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