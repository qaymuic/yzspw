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
	alert("�����ֻ��ͬʱ�ϴ�<%=upset%>���ļ�!");
	window.form.upcount.value = <%=upset%>;
	setid();
	}
	else{
	for(i=1;i<=window.form.upcount.value;i++)
	str+='�ļ�'+i+':<input type="file" name="file'+i+'" style="width:200"><br>';
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
��û���ڱ���̳�ϴ��ļ���Ȩ��
<%else%>
<Input type="hidden" name="act" value="upload">
<input type="hidden" name="upnum" value="upload">
<TD id="upid" class=tablebody2 valign=top>
<input type="file" name="file1" width=200 value="" size="40"></TD>
<td class=tablebody2 valign=top width=1>
<input type="submit" name="Submit" value="�ϴ�" onclick="parent.document.Dvform.Submit.disabled=true,
parent.document.Dvform.Submit2.disabled=true;">
</td>
<td id=allupload class=tablebody2 valign=top>
<% if upset > 1 then %>
<input type="button" name="setload" onClick="uploadframe(25);" value="�����ϴ�">
<% end if %>
</td><td class=tablebody2 valign=top>
<DIV id=Setupload style="display:none">
<% if upset > 1 then %>
�����ϴ��ĸ���
<input type="text" value="1" name="upcount" style="width:40">
<input type="button" name="Button" onClick="setid();" value="�趨"><br>(ÿ�ο�������ͬʱ�ϴ�<font color="#FF0000"><%=upset%></font>���ļ�)
<% end if %>
</div>
<font color=<%=Dvbbs.mainsetting(1) %> >���컹���ϴ�<%=Dvbbs.Groupsetting(50)-dateupnum%>��</font>��
  <a style="CURSOR: help" title="��̳����:һ��<%=Dvbbs.Groupsetting(40)%>����һ��<%=Dvbbs.Groupsetting(50)%>��,ÿ��<%=Dvbbs.Groupsetting(44)%>K">(�鿴��̳����)</a>
</TD>
<%end if%>
</tr>
</table>
</form>
</body>
</html>
