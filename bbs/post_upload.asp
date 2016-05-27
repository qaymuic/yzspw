<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dvbbs.Loadtemplates("")
Dvbbs.Head()
Dvbbs.ShowErr
dim dateupnum
dim upset
upset=CInt(Dvbbs.Groupsetting(66))
dateupnum=Cint(Dvbbs.UserToday(2))
%>
<script>
if (top.location==self.location){
	top.location="index.asp"
}
function uploadframe(num)
{
	var obj=parent.document.getElementById("ad");
	if (parseInt(obj.height)+num>=24) {
		obj.height = parseInt(obj.height) + num;
	document.getElementById("Setupload").style.display="";
	document.getElementById("allupload").style.display="none";
	}
}
function setid()
{
	str='';
	if(!window.form.upcount.value)
	window.form.upcount.value=1;
	if(window.form.upcount.value><%=upset%>){
	alert("您最多只能同时上传<%=upset%>个文件!");
	window.form.upcount.value = <%=upset%>;
	setid();
	}
	else{
	for(i=1;i<=window.form.upcount.value;i++)
	str+='文件'+i+':<input type="file" name="file'+i+'" style="width:200"><br>';
	window.upid.innerHTML=str+'<br>';

	var num=i*16
	var obj=parent.document.getElementById("ad");
	if (parseInt(obj.height)+num>=24) {
		obj.height = 24 + num;	
	}
	}
}
</script>
<form name="form" method="post" action="post_upfile.asp?boardid=<%=request("boardid")%>" enctype="multipart/form-data">
<table border=0 cellspacing=0 cellpadding=0 style="width:100%;height:100%">
<tr>
<%if Cint(Dvbbs.Groupsetting(7))=0 then%>
您没有在本论坛上传文件的权限
<%else%>
<Input type="hidden" name="act" value="upload">
<input type="hidden" name="upnum" value="upload">
<TD id="upid" class=tablebody2 valign=top>
<input type="file" name="file1" width=200 value="" size="40"></TD>
<td class=tablebody2 valign=top width=1>
<input type="submit" name="Submit" value="上传" onclick="parent.document.Dvform.Submit.disabled=true,
parent.document.Dvform.Submit2.disabled=true;">
</td>
<td id=allupload class=tablebody2 valign=top>
<% if upset > 1 then %>
<input type="button" name="setload" onClick="uploadframe(25);" value="批量上传">
<% end if %>
</td><td class=tablebody2 valign=top>
<DIV id=Setupload style="display:none">
<% if upset > 1 then %>
设置上传的个数
<input type="text" value="1" name="upcount" style="width:40">
<input type="button" name="Button" onClick="setid();" value="设定"><br>(每次可以设置同时上传<font color="#FF0000"><%=upset%></font>个文件)
<% end if %>
</div>
<font color=<%=Dvbbs.mainsetting(1) %> >今天还可上传<%=Dvbbs.Groupsetting(50)-dateupnum%>个</font>；
  <a style="CURSOR: help" title="论坛限制:一次<%=Dvbbs.Groupsetting(40)%>个，一天<%=Dvbbs.Groupsetting(50)%>个,每个<%=Dvbbs.Groupsetting(44)%>K">(查看论坛限制)</a>
</TD>
<%end if%>
</tr>
</table>
</form>
</body>
</html>
