<!--#include file="conn.asp"-->
<%
Const PurviewLevel=2
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim SiteName,SiteUrl,SiteIntro,ImgUrl,ImgWidth,ImgHeight,IsFlash,sql,advid,endtime1,ImgUrl1
SiteName=trim(request("SiteName"))
endtime1=trim(request("endtime1"))
SiteUrl=trim(request("SiteUrl"))
SiteIntro=trim(request("SiteIntro"))
ImgUrl=trim(request("Document1"))
ImgUrl1=trim(request("Document2"))
if ImgUrl="" then Imgurl="��ͼƬ"
ImgWidth=trim(request("ImgWidth"))
ImgHeight=Trim(request("ImgHeight"))
IsFlash=trim(request("IsFlash"))
advid=trim(request("advid"))
if SiteName<>"" and ImgUrl<>"" then
	if SiteUrl="http://" then SiteUrl="http://www.yzspw.com"
	if ImgWidth="" then 
		ImgWidth=0
	else
		ImgWidth=Cint(ImgWidth)
	end if
	if ImgHeight="" then
		ImgHeight=0
	else
		ImgHeight=Cint(ImgHeight)
	end if
	if IsFlash="True" then IsFlash=1
	if Isflash="False" then isflash=0
	sql="Insert Into adv (SiteName,SiteUrl,SiteIntro,ImgUrl,ImgUrl1,ImgWidth,ImgHeight,IsFlash,advid,endtime) values ('" & SiteName & "','" & SiteUrl & "','" & SiteIntro & "','" & ImgUrl  & "','" & ImgUrl1  & "'," & ImgWidth & "," & ImgHeight & "," & IsFlash & "," & advid & ",'" & endtime1 & "')"
	conn.execute sql
	call CloseConn()
	response.write  "<script>alert('�����ɹ���');location.href='admin_advmanage.asp'</script>"
end if
%>

<html>
<head>
<title>��ӹ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>
<body>
<form name="myform" method="post" action="admin_advAdd.asp">
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="999999" >
    <tr bgcolor="#cccccc"> 
      <td height="30" colspan="8"><span class="masterTitle">������������ -&gt; ���ӹ��</span></td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td width="21%"  align="right"><span class="masterTitle">���</span>���⣺</td>
      <td width="79%"  bgcolor="eeeeee"> <input name="SiteName" type="text" class="Inpt" id="SiteName" value="" size="50" maxlength="255">
        <font color="#FF0000">*</font></td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td  align="right"><span class="masterTitle">���</span>��ַ��</td>
      <td > <input name="SiteUrl" type="text" class="Inpt" id="SiteUrl" value="http://" size="50" maxlength="255"></td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td  align="right"><span class="masterTitle">���</span>˵����</td>
      <td > <textarea name="SiteIntro" cols="50" rows="4" class="Inpt" id="SiteIntro"></textarea> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td align="right">СͼƬ(Flash)��ַ��</td>
      <td> <iframe name="ad1" frameborder=0 width=100% height=20 scrolling=no src=/uploadface1.asp></iframe> 
        <br> <input type="text" name="Document1" size="48" class="Inpt">
        <font color="#FF0000">*</font></td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td align="right">СͼƬ(Flash)��С��</td>
      <td>�� 
        <input name="ImgWidth" type="text" id="ImgWidth" value="160" size="3" maxlength="5">
        ����&nbsp;&nbsp;&nbsp;&nbsp;�ߣ� 
        <input name="ImgHeight" type="text" id="ImgHeight" value="40" size="3" maxlength="5">
        ����&nbsp;&nbsp;&lt;1-6��С������Ч&gt;</td>
    </tr>
	<!--
    <tr bgcolor="#FFFFFF" >
      <td height="29" align="right">��ͼƬ(Flash)��ַ��</td>
      <td><iframe name="ad1" frameborder=0 width=100% height=20 scrolling=no src=../uploadface2.asp></iframe> 
        <br> <input type="text" name="Document2" size="48" class="Inpt"></td>
    </tr>
	-->
    <tr bgcolor="#FFFFFF" > 
      <td align="right">�Ƿ�FLASH��</td>
      <td><input type="radio" name="IsFlash" value="True">
        ��&nbsp;&nbsp;&nbsp;&nbsp; <input name="IsFlash" type="radio" value="False" checked>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td align="right">����λ�ã�</td>
      <td><input name="advid" type="text" class="Inpt" id="advid" value="1" size="4">
        &nbsp;(�����֣�1��������,2,3,4,5,6,7�������)</td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td align="right"><span class="style1">������Ч�ڣ�</span></td>
      <td><input name="endtime1" type="text" class="Inpt" id="endtime1" value="2004-12-31" size="10">
        (��д��ʽ��2004-12-1)</td>
    </tr>
    <tr align="center" bgcolor="eeeeee" > 
      <td height="40" colspan="2"> <input name="Submit" type="submit" class="btn" value=" �� �� ">
        �� 
        <input name="Submit2" type="button" class="btn" onclick="javascript:history.back(-1)" value=" �� �� "></td>
    </tr>
  </table>
</form>
</body>
</html>
