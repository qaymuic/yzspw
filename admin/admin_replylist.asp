<!--#include file="Conn.asp"-->
<!-- #include file="../inc/function.asp" -->
<!--< %
dim rs,sql
Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from article_pin"
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<script>alert('�����ۣ��뷵�أ�');history.back(-1);</script>"
response.end
end if
%>-->
<%
dim strFileName,md,txt,searchmd,lasttxt
const MaxPerPage=25
dim totalPut,CurrentPage,TotalPages,UserName
dim i,j
strFileName="admin_replylist.asp"
md=trim(request("md"))
if md="" then
	if trim(request("lasttxt"))<>"" then lasttxt=request("lasttxt")
else
	searchmd=cint(request("xlist"))
	txt="%"&trim(request("txt"))&"%"
	lasttxt=" where "
	select case searchmd
		case 0
			lasttxt=lasttxt&"article_title like '"&txt&"' or Name like '"&txt&"' or title like '" &txt&"' or content like '"&txt&"' "
		case 1
			lasttxt=lasttxt&"article_title like '"&txt&"' "
		case 2
			lasttxt=lasttxt&"title like '" &txt&"' "
		case 3
			lasttxt=lasttxt&"content like '"&txt&"' "
	end select
end if
if lasttxt<>"" then
	md=server.URLEncode(lasttxt)
	strFileName=strFileName&"?lasttxt="&md
end if
dim strPurview
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from article_pin"
sql=sql&lasttxt& " order by addtime desc"
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "<script>alert('��û�����ۣ��뷵�أ�');history.back(-1);</script>"
	response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�鿴</title>
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ����������"))
     return true;
   else
     return false;
}
function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=no' );
}
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
<link href="../css/text.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E6E6E6">
  <tr>
    <td class="TD-MENU"><b>=&gt;&nbsp;��������</b></td>
  </tr>
  <tr>
    <td bgcolor="#f5f5f5" class="TD-MENU">
	<table width="100%"  border="0" cellspacing="0" cellpadding="0">
	<form name="form1" method="post" action="<%=strFileName%><%if instr(strFileName,"?")>0 then%>&md=saerch<%else%>?md=saerch<%end if%>">
      <tr>
        <td width="17%" align="right" class="td-tianchong-4px">�ؼ���</td>
        <td width="40%" align="left" class="td-tianchong-4px">
          <input type="hidden" name="lasttxt" value="<%=lasttxt%>"><input type="text" name="Txt">
          <select name="xlist">
            <option value="0" selected>ģ������</option>
            <option value="1">���±���</option>
            <option value="2">���۱���</option>
            <option value="3">��������</option>
          </select>
        </td>
        <td width="43%" align="left" class="td-tianchong-4px"><input type="submit" name="Submit" value="�� ��"></td>
      </tr></form>
    </table></td>
  </tr>
</table>
<%
		totalPut=rs.recordcount
	    if currentPage=1 then
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"��"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage

            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"��"
        	else
	        	currentPage=1
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"��"
	    	end if
		end if
sub showContent
   	dim i
    i=0
	%>
<table width="100%" height="116" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#B0ECD9"> 
    <td height="22" colspan="2" class="TD-MENU"><b>�����������£�</b> 
      <div align="center"></div></td>
  </tr>
<%  
do while not rs.eof %>
  <tr> 
    <td width="88%" height="22" bgcolor="#eeeeee" class="TD-MENU"><strong>���±���</strong>��<%=rs("article_title")%>&nbsp;&nbsp;&nbsp;<strong>���۱���</strong>��<%=rs("title")%><br><strong>������</strong>��<%=rs("Name")%></a> 
    <strong>������ʱ��</strong>��<%=rs("addtime")%>��<strong>IP:<%=rs("user_ip")%>��<a href="../list.asp?id=<%=rs("article_id")%>" target="_blank">ԭ��</a></strong></td>
    <td width="12%" height="22" bgcolor="#eeeeee">
<div align="center"><a href=admin_replydel.asp?id=<%=rs("id")%>&article_id=<%=rs("article_id")%> onClick="return ConfirmDel();"><font color="#FF0000">ɾ��</font></a></div></td>
  </tr>
  <tr> 
    <td colspan="2" bgcolor="#FFFFFF" class=text-p><%=rs("Content")%></td>
  </tr>
  <%
  i=i+1
  if i>=MaxPerPage then exit do
  rs.movenext
  loop
%>
</table>
<%  end sub%>
</body>
</html>
<%
set rs=nothing
closeconn
%>