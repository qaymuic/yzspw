<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=5    '����Ȩ��
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/config.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>�������</title>
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
    alert("���ⲻ��Ϊ�գ�");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("���ݲ���Ϊ�գ�");
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
      <td height="20" align="center" class="title"><b>�� �� �� ��</b></td>
    </tr>
    <tr align="center">
      <td class="tdbg">
	<table width="100%" border="0" cellpadding="2" cellspacing="0">
          <tr> 
            <td width="90" align="right">������Ŀ��</td>
            <td width="500"> <font color="#FF0000">
              <%
        sql = "select * from BigClass where trim(BigClassName)<>'���̶�̬'"
        rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "���������Ŀ��"
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
                <option value="" selected>��ָ��С��</option>
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
            <td align="right">�ꡡ���⣺</td>
            <td width="500"><input name="Title" type="text"
           id="Title" size="70" maxlength="255"></td>
          </tr>
          <tr> 
            <td width="90" align="right" valign="middle">�ڡ����ݣ�</td>
            <td><textarea name="Content" style="display:none"></textarea> 
              <iframe ID="editor" src="../editor.asp" frameborder=1 scrolling=no width="530" height="405">������Ҫ������Ҫ</iframe> 
            </td>
          </tr>
          <tr> 
            <td width="90" align="right">����ͼƬ��</td>
            <td><input name="IncludePic" type="checkbox" id="IncludePic" value="yes">
              ��<font color="#0000FF">��ѡ���ڱ���ǰ����ʾ[ͼ��]��</font></td>
          </tr>
          <tr> 
            <td align="right">��ҳͼƬ��</td>
            <td><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" size="40" maxlength="200"> 
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option selected>ѡ��ͼƬ</option>
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles"> 
            </td>
          </tr>
          <% if session("purview")<>"" and session("purview")<=2 then %>
          <tr> 
            <td width="90" align="right">�����ˣ�</td>
            <td><%=session("admin")%>���Ƿ��Ƽ� 
            <input name="istop" type="checkbox" id="istop" value="yes">
��</td>
          </tr>
          <% end if %>
        </table>
      </td>
    </tr>
  </table>
  <div align="center">
    <p> 
      <input
  name="Add" type="submit"  id="Add" value=" �� �� " onClick="document.myform.action='ArticleSave.asp?action=add';document.myform.target='_self';">
      ��&nbsp; 
      <input type="reset" name="Submit" value=" ȡ �� ">
    </p>
  </div>
</form>
</body>
</html>
