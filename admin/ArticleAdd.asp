<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=5    '操作权限
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/config.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>添加内容</title>
<link rel="stylesheet" type="text/css" href="style.css">
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass order by SmallClassID asc"
rs.open sql,conn,1,1
%>
<script language = "JavaScript">
var onecount;
subcat = new Array();
        <%
        count = 0
        do while not rs.eof 
        %>
subcat[<%=count%>] = new Array("<%= trim(rs("SmallClassName"))%>","<%= trim(rs("BigClassName"))%>","<%= trim(rs("SmallClassName"))%>");
        <%
        count = count + 1
        rs.movenext
        loop
        rs.close
        %>
onecount=<%=count%>;

function changelocation(locationid)
    {
    document.myform.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassName.options[document.myform.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    

function CheckForm()
{
  if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
  else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 

  if (document.myform.Title.value=="")
  {
    alert("标题不能为空！");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("内容不能为空！");
	editor.HtmlEdit.focus();
	return false;
  }
  return true;  
}
function loadForm()
{
  editor.HtmlEdit.document.body.innerHTML=document.myform.Content.value;
  return true
}
</script>

</head>

<body leftmargin="5" topmargin="10" onLoad="javascipt:setTimeout('loadForm()',1000);">
<form method="POST" name="myform" onSubmit="return CheckForm();" action="ArticleSave.asp?action=add" target="_self">
  <table width="615" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr>
      <td height="20" align="center" class="title"><b>添 加 内 容</b></td>
    </tr>
    <tr align="center">
      <td class="tdbg">
	<table width="100%" border="0" cellpadding="2" cellspacing="0">
          <tr> 
            <td width="90" align="right">所属栏目：</td>
            <td width="500"> <font color="#FF0000">
              <%
        sql = "select * from BigClass where trim(BigClassName)<>'商铺动态'"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "请先添加栏目。"
		else
		%> <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
                <option selected value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
			dim selclass
		    selclass=rs("BigClassName")
        	rs.movenext
		    do while not rs.eof
			%>
                <option value="<%=trim(rs("BigClassName"))%>"><%=trim(rs("BigClassName"))%></option>
                <%
		        rs.movenext
    	    loop
		end if
        rs.close
			%>
              </select> <select name="SmallClassName">
                <option value="" selected>不指定小类</option>
                <%
			sql="select * from SmallClass where BigClassName='" & selclass & "'"
			rs.open sql,conn,1,1
			if not(rs.eof and rs.bof) then
			%>
                <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <% rs.movenext
				do while not rs.eof%>
                <option value="<%=rs("SmallClassName")%>"><%=rs("SmallClassName")%></option>
                <%
			    	rs.movenext
				loop
			end if
	        rs.close
			%>
              </select></font></td>
          </tr>
          <tr> 
            <td align="right">标　　题：</td>
            <td width="500"><input name="Title" type="text"
           id="Title" size="70" maxlength="255"></td>
          </tr>
          <tr> 
            <td width="90" align="right" valign="middle">内　　容：</td>
            <td><textarea name="Content" style="display:none"></textarea> 
              <iframe ID="editor" src="../editor.asp" frameborder=1 scrolling=no width="530" height="405">内容提要内容提要</iframe> 
            </td>
          </tr>
          <tr> 
            <td width="90" align="right">包含图片：</td>
            <td><input name="IncludePic" type="checkbox" id="IncludePic" value="yes">
              是<font color="#0000FF">（选中在标题前面显示[图文]）</font></td>
          </tr>
          <tr> 
            <td align="right">首页图片：</td>
            <td><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="40" maxlength="200"> 
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option selected>选择图片</option>
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles"> 
            </td>
          </tr>
          <% if session("purview")<>"" and session("purview")<=2 then %>
          <tr> 
            <td width="90" align="right">发布人：</td>
            <td><%=session("admin")%>　是否推荐 
            <input name="istop" type="checkbox" id="istop" value="yes">
是</td>
          </tr>
          <% end if %>
        </table>
      </td>
    </tr>
  </table>
  <div align="center">
    <p> 
      <input
  name="Add" type="submit"  id="Add" value=" 添 加 " onClick="document.myform.action='ArticleSave.asp?action=add';document.myform.target='_self';">
      　&nbsp; 
      <input type="reset" name="Submit" value=" 取 消 ">
    </p>
  </div>
</form>
</body>
</html>
