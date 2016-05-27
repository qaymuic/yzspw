<%
if session("admin") = "" and session("Purview") = "" then
    Error()
    response.end
end if    
%>
<html>
<head>
</head>
<link rel="stylesheet" href="html.css">
<title>网站管理系统--后台管理</title>
<frameset framespacing="0" border="false" cols="160,*" frameborder="1">
<frame name="left"  scrolling="auto" marginwidth="0" marginheight="0" src="Left.asp">
<frame name="right" scrolling="auto" src="Main.asp">
  </frameset>
  <noframes>
  <body>
  <p>This page uses frames, but your browser doesn't support them.</p>
  </body>
  </noframes>
</html>

<%
	sub Error()
		response.write "   <html><head><link rel='stylesheet' href='style.css'></head><body>"
	    	response.write "   <br><br><br>"
	    	response.write "    <table align='center' width='300' border='0' cellpadding='4' cellspacing='0' class='border'>"
	    	response.write "      <tr > "
	    	response.write "        <td class='title' height='20'> "
	    	response.write "          <div align='center'>操作: 确认身份失败!</div>"
	    	response.write "        </td>"
	    	response.write "      </tr>"
	    	response.write "      <tr>"
	    	response.write "        <td class='tdbg' height='23'> "
	    	response.write "          <div align='center'><br><br>"
	    	response.write "      非法登陆,您的操作已经被记录!!! <br><br>"
	    	response.write "        <a href='javascript:onclick=history.go(-1)'>返回</a>"        
	    	response.write "        <br><br></div></td>"
	    	response.write "      </tr></table></body></html>" 
	end sub
%>