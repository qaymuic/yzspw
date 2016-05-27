<%
if session("admin")="" then
	response.redirect "login.asp"
else
	if session("purview")>PurviewLevel then
		response.write "<br><p align=center><font color='red'>您没有操作的权限</font></p>"
		response.end
	end if
end if
%>
