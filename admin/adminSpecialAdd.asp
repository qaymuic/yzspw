<%@language=vbscript codepage=936 %>
<%
if session("admin")=empty then 
response.redirect "admin.asp"
end if
%>
<!--#include file="conn.asp"-->
<%
dim SpecialName,rs
SpecialName=trim(request.Form("SpecialName"))
if SpecialName<>"" then
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.open "Select * From Special Where SpecialName='" & SpecialName & "'",conn,1,3
	if not rs.EOF then
		rs.close
	    set rs=Nothing
    	call CloseConn()
    	Response.Redirect "adminSpecialManage.asp?Err=SpecialExist"  
	else
     	rs.addnew
     	rs("SpecialName")=SpecialName
     	rs.update
     	rs.Close
     	set rs=Nothing
     	call CloseConn()
	end if
end if
Response.Redirect "adminSpecialManage.asp"  
%>