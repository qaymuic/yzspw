<!--#include file="conn.asp"-->
<%
Const PurviewLevel=2
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim ID,SiteName,SiteUrl,SiteIntro,ImgUrl,ImgWidth,ImgHeight,IsFlash,sql,rs,advid,endtime1,ImgUrl1
ID=trim(request("ID"))
endtime1=trim(request("endtime1"))
advid=trim(request("advid"))
SiteName=trim(request("SiteName"))
SiteUrl=trim(request("SiteUrl"))
SiteIntro=trim(request("SiteIntro"))
ImgUrl=trim(request("document1"))
ImgUrl1=trim(request("Document2"))
if ImgUrl="" then Imgurl="无图片"
ImgWidth=trim(request("ImgWidth"))
ImgHeight=Trim(request("ImgHeight"))
IsFlash=trim(request("IsFlash"))
if ID="" then
	response.Redirect "admin_advManage.asp"
end if
sql="select * from adv where ID=" & clng(ID)
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
	rs.close
	set rs=nothing
	call CloseConn()	
	response.redirect "admin_advManage.asp"
end if

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
	if IsFlash="False" then IsFlash=0
	if IsFlash="True" then IsFlash=1
	rs("SiteName")=SiteName
	rs("SiteUrl")=SiteUrl
	rs("SiteIntro")=SiteIntro
	rs("ImgUrl")=ImgUrl
	rs("ImgUrl1")=ImgUrl1
	rs("ImgWidth")=ImgWidth
	rs("ImgHeight")=ImgHeight
	rs("IsFlash")=IsFlash
	rs("endtime")=endtime1
	rs("advid")=advid
	rs.update
	rs.close
	set rs=nothing
	call CloseConn()
	response.redirect "admin_advManage.asp"
end if
%>

<html>
<head>
<title>修改广告</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="STYLE.CSS" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>
<body>
<form name="myform" method="post" action="admin_advModify.asp">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="2">
  <table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="999999" >
    <tr bgcolor="#eeeeee">
      <td height="30" colspan="8"><span class="masterTitle">　　　广告管理 -&gt; 修改广告</span></td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td width="214" align="right">广告名称：</td>
      <td width="748"> 
<input name="SiteName" type="text" class="Inpt" id="SiteName" value="<%=rs("SiteName")%>" size="50" maxlength="255">
        <font color="#FF0000">*</font></td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td align="right">广告地址：</td>
      <td width="748"> <input name="SiteUrl" type="text" class="Inpt" id="SiteUrl" value="<%=rs("SiteUrl")%>" size="50" maxlength="255">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td align="right">广告简介：</td>
      <td width="748"> <textarea name="SiteIntro" cols="50" rows="4" class="Inpt" id="SiteIntro"><%=rs("SiteIntro")%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td align="right">图片地址：</td>
      <td>      <iframe name="ad1" frameborder=0 width=100% height=20 scrolling=no src=/uploadface1.asp></iframe><br>
                  <input type="text" name="Document1" size="48" class="Inpt" value="<%=rs("ImgUrl")%>">
        <font color="#FF0000">*</font></td>
    </tr>
    <tr bgcolor="#FFFFFF" > 
      <td align="right">图片大小：</td>
      <td>宽： 
        <input name="ImgWidth" type="text" id="ImgWidth" value="<%=rs("ImgWidth")%>" size="6" maxlength="5">
        像素&nbsp;&nbsp;&nbsp;&nbsp;高： 
        <input name="ImgHeight" type="text" id="ImgHeight" value="<%=rs("ImgHeight")%>" size="6" maxlength="5">
      像素&nbsp;&nbsp;&lt;1-6大小设置无效&gt;</td>
    </tr>
	<!--
	<tr bgcolor="#FFFFFF" >
      <td height="29" align="right">大图片(Flash)地址：</td>
      <td><iframe name="ad1" frameborder=0 width=100% height=20 scrolling=no src=../uploadface2.asp></iframe> 
        <br> 
        <input name="Document2" type="text" class="Inpt" value="< %=rs("ImgUrl1")% >" size="48"></td>
    </tr>
	-->
    <tr bgcolor="#FFFFFF" > 
      <td align="right">是否FLASH：</td>
      <td><input type="radio" name="IsFlash" value="True" <% if rs("IsFlash")=true then response.write "checked"%>>
        是&nbsp;&nbsp;&nbsp;&nbsp; <input name="IsFlash" type="radio" value="False" <% if rs("IsFlash")=False then response.write "checked"%>>
      否</td>
    </tr>
    <tr bgcolor="#FFFFFF" >
      <td align="right">广告位置：</td>
      <td><input name="advid" type="text" class="Inpt" id="advid" value="<%=rs("advid")%>" size="4">
        &nbsp;(填数字，1友情链接,2,3,4,5,6,7浮动广告)</td>
    </tr>
	    <tr bgcolor="#FFFFFF" >
      <td align="right"><span class="style1">广告有效期：</span></td>
      <td><input name="endtime1" type="text" class="Inpt" id="endtime1" value="<%=rs("endtime")%>" size="10">
        (填写格式：2004-12-1)</td>
    </tr>
    <tr align="center" bgcolor="eeeeee" > 
      <td height="40" colspan="2"><input name="ID" type="hidden" id="ID" value="<%=rs("ID")%>">
      <input name="Submit" type="submit" class="btn" value=" 保 存 ">　
      <input name="Submit2" type="button" class="btn" onclick="javascript:history.back(-1)" value="返 回"></td>
    </tr>
  </table>
</form>
</body>
</html>
