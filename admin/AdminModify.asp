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
dim Action,UserID, password,PwdConfirm, purview, rs, sql,FoundErr,ErrMsg,realname
Action=trim(request("Action"))
UserID=trim(Request("ID"))
realname=trim(Request("realname"))
password=trim(Request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
purview=trim(Request("purview"))
if UserID="" then
	response.Redirect("AdminManage.asp")
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from Admin where ID="&UserID&"",conn,1,3
if rs.Bof and rs.EOF then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�����ڴ��û���</li>"
else
	if Action="Modify" then
		if rs("UserName")="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�û�������Ϊ�գ�</li>"
		end if
		if rs("realname")="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ʵ��������Ϊ�գ�</li>"
		end if
		if PwdConfirm<>Password then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>ȷ�������������������ͬ��</li>"
		end if
		if Purview="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>�û�Ȩ�޲���Ϊ�գ�</li>"
		end if
		if FoundErr<>True then
			if Password<>"" then
				rs("password")=md5(password)
			end if
	     	rs("purview")=Cint(purview)  
			rs("realname")=realname
    	 	rs.update
			rs.Close
	     	set rs=Nothing
    	 	call CloseConn()
     		Response.Redirect "AdminManage.asp"
		end if
	end if
	purview=rs("purview")
end if
if FoundErr=True then
	Call WriteErrMsg()
else
%>
<html>
<head>
<title>�޸Ĺ���Ա��Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<body>
<form method="post" action="AdminModify.asp" name="form1">
    
  <table width="400" border="0" align="center" cellpadding="2" cellspacing="2" class="border" >
    <tr class="title"> 
      <td height="20" colspan="2"> <div align="center"><font size="2"><strong>�޸Ĺ���Ա��Ϣ</strong></font></div></td>
    </tr>
    <tr> 
      <td align="right" class="tdbg"><strong>�� �� ����</strong></td>
      <td class="tdbg"><%=rs("UserName")%> <input name="ID" type="hidden" value="<%=rs("ID")%>"></td>
    </tr>
    <tr>
      <td align="right" class="tdbg"><strong>��ʵ������</strong></td>
      <td class="tdbg"><input name="realname" type="text" id="realname" value="<%=rs("realname")%>"></td>
    </tr>
    <tr> 
      <td align="right" class="tdbg"><strong>�� �� �룺</strong></td>
      <td class="tdbg"><input type="password" name="Password"> <font color="#0000FF">��������޸ģ��뱣��Ϊ��</font></td>
    </tr>
    <tr> 
      <td align="right" class="tdbg"><strong>ȷ�����룺</strong></td>
      <td class="tdbg"><input type="password" name="PwdConfirm"> <font color="#0000FF">��������޸ģ��뱣��Ϊ��</font></td>
    </tr>
    <tr> 
      <td align="right" class="tdbg"><strong>Ȩ�����ã� </strong></td>
      <td class="tdbg"><select name="purview" id="purview">
          <option value="1" <%if purview=1 then %>selected<% end if %>>����Ա</option>
          <option value="2" <%if purview=2 then %>selected<% end if %>>��������</option>
          <option value="3" <%if purview=3 then %>selected<% end if %>>һ���û�</option>
        </select></td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="Modify"> 
        <input  type="submit" name="Submit" value=" ȷ �� "></td>
    </tr>
  </table>
  </form>
</body>
</html>
<%
end if
rs.close
set rs=nothing
call CloseConn()
%>
