<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim a1,a2,a3,a4,a5,a6,a7,founderr,errmsg
founderr=false
a1=trim(request("a1"))
a3=trim(request("a3"))
a4=trim(request("a4"))
a5=trim(request("a5"))
a6=trim(Request("a6"))
a7=trim(request("a7"))
a2=replace(replace(request.form("a2")," ","&nbsp;"),chr(13),"<br>")
if a1="" then
	founderr=true
	errmsg=errmsg & "<br><li>����������Ŀ�ģ�</li>"
end if
if a2="" then
	founderr=true
	errmsg=errmsg & "<br><li>������������ݣ�</li>"
end if
if a4="" then
	founderr=true
	errmsg=errmsg & "<br><li>��������ϵ�ˣ�</li>"
end if
if a5="" then
	founderr=true
	errmsg=errmsg & "<br><li>��������ϵ�绰��</li>"
end if

if founderr=false then
	dim sqlReg,rsReg
	sqlReg="select * from jgpg "
	set rsReg=server.createobject("adodb.recordset")
	rsReg.open sqlReg,conn,1,3
		rsReg.addnew
		rsReg("a1")=a1
		rsReg("a2")=a2
		rsReg("a3")=a3
		rsReg("a4")=a4
		rsReg("a5")=a5
		rsReg("a6")=a6
		rsReg("a7")=a7
		rsReg.update
		founderr=false
	rsReg.close
	set rsReg=nothing
end if		
%>
<html>
<head>
<title>��ɼ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.CSS" rel="stylesheet" type="text/css">
</head>

<body>
<table width="98%" height="300" border="0" align="center" cellpadding="0" cellspacing="0"  bgcolor="#FFFFFF" style="border-collapse: collapse">
  <tr> 
    <td align="center" valign=top> 
      <%
if founderr=false then
	call RegSuccess()
else
	call WriteErrmsg()
end if
%>
</td>
  </tr>
</table>
</body>
</html>
<%
call CloseConn

sub WriteErrMsg()
    response.write "<table align='center' width='98%' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>�������µ�ԭ���ܼ��뷿Դ��</td></tr>"
    response.write "<tr class='tdbg'><td align='left' height='100'><br>" & errmsg & "<p align='center'>��<a href='javascript:onclick=history.go(-1)'>�� ��</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub

sub RegSuccess()
    response.write "<table align='center' width='98%' border='0' cellpadding='2' cellspacing='0' class='border'>"
    response.write "<tr class='title'><td align='center' height='15'>�ɹ�������������������Ϣ�����ǽ�����������ϵ��</td></tr>"
    response.write "<tr class='tdbg'><td align='center' height='100'><br><p align='center'>��<a href='pinggulist.asp'>������������</a>��<br></p></td></tr>"
	response.write "</table>" 
end sub
%>