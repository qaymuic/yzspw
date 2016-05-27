<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dvbbs.LoadTemplates("usermanager")
Dvbbs.Stats=Dvbbs.MemberName&template.Strings(1)
Dvbbs.Head()
%>
	<table border="0"  cellspacing="0" cellpadding="0" width=100%>
	<tr>
	<td class=tablebody1>
	<form name="form" method="post" action="upfile.asp" enctype="multipart/form-data" >
	<input type="hidden" name="filepath" value="uploadFace">
	<input type="hidden" name="act" value="upload">
	<input type="file" name="file1">
	<input type="hidden" name="fname">
	<input type="submit" name="Submit" value="ÉÏ´«" onclick="fname.value=file1.value,parent.document.theForm.Submit.disabled=true,parent.document.theForm.Submit2.disabled=true;">
	</form>
</body>
</html>