<%@language=vbscript codepage=936 %>
<%
if session("admin")=empty then 
response.redirect "admin.asp"
end if
%>
<!--#include file="conn.asp"-->
<%
dim SpecialID, SpecialName,OldSpecialName,rs
SpecialID=trim(Request("SpecialID"))
SpecialName=trim(Request.form("SpecialName"))
OldSpecialName=trim(request.form("OldSpecialName"))
if SpecialID="" then
  response.Redirect("SpecialManage.asp")
else
  Set rs=Server.CreateObject("Adodb.RecordSet")
  rs.Open "Select * from Special where SpecialID="&SpecialID&"",conn,1,3
  if rs.Bof or rs.EOF then
   	Response.Write "<font color=red><div align=center><br><br>�����̷��಻����</div></font>"
   	Response.End   
	rs.close
  else
	if SpecialName<>"" then
		rs("SpecialName")=SpecialName
     	rs.update
		rs.Close
     	set rs=Nothing
     	call CloseConn()
     	Response.Redirect "adminSpecialManage.asp"
	end if
%>
<html>
<head>
<title>�޸����̷�������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Style.css" rel="stylesheet" type="text/css">
<style>
A:link {
	COLOR: #000000; TEXT-DECORATION: none
}
A:visited {
	COLOR: #000000; TEXT-DECORATION: none
}
A:active {
	COLOR: #000000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #000000; TEXT-DECORATION: none
}
TD {
	FONT: 9pt/12pt "����"
}
.clkun {
	FONT-SIZE: 13px;
	LINE-HEIGHT: 24px;
	word-break:break-all;
	color: #ffffff;
}
.clkun1 {
	FONT-SIZE: 15px;
	LINE-HEIGHT: 32px;
	word-break:break-all;
	color: #ffffff;
	font-family: "����";
}
.clkun2 {
	FONT-SIZE: 15px;
	LINE-HEIGHT: 32px;
	word-break:break-all;
	color: #ffffff;
	font-family: "����";
}
</style>
<style>
body{
    scrollbar-face-color: #EAEAEA;
    scrollbar-shadow-color: #666666;
    scrollbar-highlight-color: #FFFFFF;
    scrollbar-3dlight-color: #666666;
    scrollbar-darkshadow-color: #DCE0E2;
    scrollbar-track-color: #FFFFFF;
    scrollbar-arrow-color: #ff0000;
.style1 {color: #FF0000}
</style>

</head>

<body>
<form name="form1" method="post" action="adminSpecialModify.asp">
  <p>&nbsp;</p>
  <table width="96%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#eeeeee" class="border">
    <tr class="title"> 
      <td colspan="2">���������̷������ -&gt; ��Ŀ�޸�</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="39%" align="right">���̷���ID��</td>
      <td width="61%"><%=rs("SpecialID")%> <input name="SpecialID" type="hidden" id="SpecialID" value="<%=rs("SpecialID")%>">
      <input name="OldSpecialName" type="hidden" id="OldSpecialName" value="<%=rs("SpecialName")%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td align="right">���̷������ƣ�</td>
      <td><input name="SpecialName" type="text" value="<%=rs("SpecialName")%>" size="20"></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" class="tdbg"> 
      <td colspan="2"><input name="Save" type="submit" id="Save" value=" �� �� ">��
        ��
        <input name="Submit2" type="button" class="btn" value=" �� �� " onclick='javascript:history.back(-1);'></td>
    </tr>
  </table>  
</form>
</body>
</html>
<%
  end if
  rs.close
  set rs=nothing
end if
conn.close
set conn=nothing
%>
