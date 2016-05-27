<!--#include file=connad.asp-->
<%
response.expires=0
response.buffer=true
response.clear

Dim rs
Dim filename

filename=request("fn")
if filename="" then
	Response.Write "错误的系统参数1。"
	Response.end
else
	filename=replace(filename,"'","")
end if

set rs=connad.execute("select * from dv_chanad where A_Adname='"&filename&"'")
if rs.eof and rs.bof then
	Response.Write "错误的系统参数2。"
	Response.end	
else
	if rs("a_adtype")="swf" then
	Response.ContentType = "application/x-shockwave-flash"
	else
	Response.ContentType = "image/" & replace(lcase(rs("a_adtype")),"jpg","jpeg")
	'Response.ContentType = "img/*"
	end if
	Response.BinaryWrite rs("A_data").GetChunk(7500000)
	'Response.Write rs("a_data")

end if
rs.close
set rs=nothing
connad.close
set connad=nothing
%>