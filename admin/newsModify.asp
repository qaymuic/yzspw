<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=5    '����Ȩ��
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/config.asp"-->
<%
dim id,rsArticle,FoundErr,ErrMsg,PurviewChecked,sqlnews
id=trim(request("id"))
FoundErr=False
PurviewChecked=False
if id="" then 
	response.Redirect("newsManage.asp")
end if
sqlnews="select * from ytiinews where id=" & id & ""
Set rsArticle= Server.CreateObject("ADODB.Recordset")
rsArticle.open sqlnews,conn,1,1
if rsArticle.bof and rsArticle.eof then
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>�Ҳ�������</li>"
else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸�����</title>
<link rel="stylesheet" type="text/css" href="style.css">
<script language = "JavaScript">
<%
dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from SmallClass order by SmallClassID asc"
rs.open sql,conn,1,1
%>
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
    alert("���ݱ��ⲻ��Ϊ�գ�");
	document.myform.Title.focus();
	return false;
  }
  if (document.myform.Content.value=="")
  {
    alert("�������ݲ���Ϊ�գ�");
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
<form method="POST" name="myform" onSubmit="return CheckForm();" action="newsSave.asp?action=Modify">
  <table width="615" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr>
      <td height="25" align="center" class="title"><b>�� �� �� ��</b></td>
    </tr>
    <tr align="center">
      <td class="tdbg">
	<table width="100%" border="0" cellpadding="2" cellspacing="0">
          <tr>
            <td align="right">������Ŀ��</td>
            <td>���̶�̬
                <input type="hidden" name="BigClassName" value="���̶�̬">
            </td>
          </tr>
          <tr> 
            <td align="right">���ݱ��⣺</td>
            <td width="500"><input name="Title" type="text"
           id="Title" value="<%=rsArticle("Title")%>" size="70" maxlength="255"></td>
          </tr>
          <tr> 
            <td width="90" align="right" valign="middle">�������ݣ�</td>
            <td><textarea name="Content" style="display:none"><%=rsArticle("Content")%></textarea> 
              <iframe ID="editor" src="../editor.asp" frameborder=1 scrolling=no width="530" height="405"></iframe> 
            </td>
          </tr>
          <tr> 
            <td width="90" align="right">����ͼƬ��</td>
            <td><input name="IncludePic" type="checkbox" id="IncludePic" value="yes" <% if rsArticle("IncludePic")=true then response.Write("checked") end if%>>
              ��<font color="#0000FF">�����ѡ�еĻ����ڱ���ǰ����ʾ[ͼ��]��</font></td>
          </tr>
          <tr> 
            <td width="90" align="right">��ҳͼƬ��</td>
            <td><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="<%=rsArticle("DefaultPicUrl")%>" size="50" maxlength="200"> 
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option value=""<% if rsArticle("DefaultPicUrl")="" then response.write "selected" %>>��ָ����ҳͼƬ</option>
                <%
				if rsArticle("UploadFiles")<>"" then
					dim IsOtherUrl
					IsOtherUrl=True
					if instr(rsArticle("UploadFiles"),"|")>1 then
						dim arrUploadFiles,intTemp
						arrUploadFiles=split(rsArticle("UploadFiles"),"|")						
						for intTemp=0 to ubound(arrUploadFiles)
							if rsArticle("DefaultPicUrl")=arrUploadFiles(intTemp) then
								response.write "<option value='" & arrUploadFiles(intTemp) & "' selected>" & arrUploadFiles(intTemp) & "</option>"
								IsOtherUrl=False
							else
								response.write "<option value='" & arrUploadFiles(intTemp) & "'>" & arrUploadFiles(intTemp) & "</option>"
							end if
						next
					else
						if rsArticle("UploadFiles")=rsArticle("DefaultPicUrl") then
							response.write "<option value='" & rsArticle("UploadFiles") & "' selected>" & rsArticle("UploadFiles") & "</option>"
							IsOtherUrl=False
						else
							response.write "<option value='" & rsArticle("UploadFiles") & "'>" & rsArticle("UploadFiles") & "</option>"		
						end if
					end If
					if IsOtherUrl=True then
						response.write "<option value='" & rsArticle("DefaultPicUrl") & "' selected>" & rsArticle("DefaultPicUrl") & "</option>"
					end if
				end if
				 %>
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles" value="<%=rsArticle("UploadFiles")%>"> 
            </td>
          </tr>
          <% if session("purview")<>"" and session("purview")<=2 then %>
          <tr> 
            <td width="90" align="right">�����ˣ�</td>
            <td><input name="Author" type="text" id="Author" value="<%=rsArticle("Author")%>" maxlength="50">
            �Ƿ���ҳ����ʾ 
              <input name="istop" type="checkbox" id="istop" value="yes" <% if rsArticle("istop")=true then response.Write("checked") end if%>>
              ��
            </td>
          </tr>
          <% end if %>
        </table>
      </td>
    </tr>
  </table>
  <div align="center"> 
    <p> 
      <input name="id" type="hidden" id="id" value="<%=rsArticle("id")%>">
      <input
  name="Save" type="submit"  id="Save" value="�����޸Ľ��">
      ���� 
      <INPUT name="button" type=button style="FONT-SIZE: 9pt" onclick="javascript:history.back(1)" value=����>
    </p>
  </div>
</form>
</body>
</html>
<%
end if
rsArticle.close
set rsArticle=nothing
call CloseConn()
%>