<%@language=vbscript codepage=936 %>
<%
if session("admin")=empty then 
response.redirect "admin.asp"
end if
%>
<!--#include file="conn.asp"-->
<%
dim rs, sql, strPurview,ErrMsg
Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from Special order by SpecialID"
rs.Open sql,conn,1,1
%>
<html>
<head>
<title>���̷������</title>
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

<script language=javascript>
function check()
{
  if(document.form1.SpecialName.value=="")
    {
      alert("���̷������Ʋ���Ϊ�գ�");
	  document.form1.SpecialName.focus();
      return false;
    }
}
</script>
</head>
<body>
<table width="96%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#eeeeee" class="border">
    <tr class="tdbg">
      <td colspan="3">���������̷������ -&gt; ��Ŀά��</td>
    </tr>
  <tr align="center" bgcolor="#FFFFFF" class="title"> 
      <td  width="8%" height="20"><strong> ���</strong></td>
      <td width="79%" height="20"><strong> ���̷�������</strong></td>
      <td  width="13%" height="20"><strong> ����</strong></td>
  </tr>
    <%while not rs.EOF %>
    <tr align="center" bgcolor="#FFFFFF" class="tdbg"> 
      <td width="8%"><%=rs("SpecialID")%></td>
      <td><%=rs("SpecialName")%></td>
      <td width="13%"> <a href="adminSpecialModify.asp?SpecialID=<%=rs("SpecialID")%>">�޸�</a>&nbsp;<a href="adminSpecialDel.asp?SpecialID=<%=rs("SpecialID")%>">ɾ��</a></td>
    </tr>

    <%
     rs.MoveNext
   Wend
  %>
</table>  
<br>
<form method="post" action="adminSpecialAdd.asp" name="form1" onsubmit="javascript:return check();">
  <table width="96%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#eeeeee" class="border" >
    <tr class="title"> 
      <td height="20" colspan="2"> ���������̷������ -&gt; ��Ŀ����</td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="45%" align="right" class="tdbg"> ���̷������ƣ�</td>
      <td width="55%" class="tdbg"><input name="SpecialName" type="text" id="SpecialName" maxlength="30"> 
      &nbsp;</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" class="tdbg"> 
      <td colspan="2" class="tdbg"><input  type="submit" name="Submit" value=" �� �� "></td>
    </tr>
  </table>
</form>
<%
if request("Err")<>"" then
  select case request("Err")
    case "SpecialExist"
	   ErrMsg="���̷��������Ѿ�����"
    case else
	   ErrMsg="δ֪����"
  end select  
  response.write "<script language='JavaScript' type='text/JavaScript'>alert('" & ErrMsg & "');</script>"
end if
%>
</body>
</html>
<%
  rs.Close
  set rs=Nothing
  
  conn.Close
  set conn=Nothing
%>