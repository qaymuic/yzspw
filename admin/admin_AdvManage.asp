<!--#include file="conn.asp"-->
<%
Const PurviewLevel=2
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim strFileName,delart,SiteName,advid
delart="删除广告"
SiteName=trim(request("SiteName"))
advid=trim(request("advid"))
const MaxPerPage=50
dim totalPut,CurrentPage,TotalPages
dim rs, sql,ID,Action
Action=Trim(Request("Action"))
ID=Trim(Request("ID"))
strFileName="admin_advManage.asp?SiteName="&SiteName&"&advid="&advid
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from adv where 1=1"
if advid="" and SiteName="" then
sql=sql
else
	 if advid<>"" then
	  sql=sql& " and advid="&advid
	 end if
	if SiteName<>"" then
	sql=sql& " and SiteName like '%"&SiteName&"%' "
	end if
sql=sql& " order by id desc"
end if
rs.Open sql,conn,1,1
%>
<html>
<head>
<title>广告管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="style.css">
<script language=javascript>
function ConfirmDel()
{
   if(confirm("确定要删除此广告吗？"))
     return true;
   else
     return false;
}
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>
<body>
<%
  	if rs.eof and rs.bof then
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;目前共有 0 个广告&nbsp;&nbsp;&nbsp;&nbsp;<a href='admin_advadd.asp'>添加广告</a>"
	else
    	totalPut=rs.recordcount
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if

    	end if
	    if currentPage=1 then
        	'showpage strFileName,totalput,MaxPerPage,true,true,"个广告"
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"个广告"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
        		'showpage strFileName,totalput,MaxPerPage,true,true,"个广告"
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"个广告"
        	else
	        	currentPage=1
        		'showpage strFileName,totalput,MaxPerPage,true,true,"个广告"
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"个广告"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%>

  <table width="95%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="999999">
  <form name="Selform" method="post" action="admin_advdel.asp">
    <tr bgcolor="#cccccc" class="title">
      <td height="30" colspan="8"><span class="masterTitle">　　　广告管理 -&gt; <a href="admin_advmanage.asp">广告列表</a>　　　<a href="admin_advadd.asp">增加广告</a>　　<a href="admin_advmanage.asp?advid=1">1友情链接</a> 
        <a href="admin_advmanage.asp?advid=2">2广告</a> <a href="admin_advmanage.asp?advid=3">3广告</a> 
        <a href="admin_advmanage.asp?advid=4">4广告</a> <a href="admin_advmanage.asp?advid=5">5广告</a>&nbsp;<a href="admin_advmanage.asp?advid=6">6广告</a></span>&nbsp;<span class="masterTitle"><a href="admin_advmanage.asp?advid=7">7浮动广告</a></span></td>
    </tr>
    <tr bgcolor="cccccc"> 
      <td width="5%"  height="20" align="center" bgcolor="cccccc">ID</td>
      <td  height="20" align="center">广告名称</td>
      <td width="18%"  height="20" align="center">广告图片</td>
      <td width="19%"  height="20" align="center" bgcolor="cccccc">广告位置</td>
      <td width="15%" align="center" bgcolor="cccccc"><span class="style1">有效时间</span></td>
      <td width="18%"  height="20" align="center"><strong> 操作</strong></td>
    </tr>
<%do while not rs.EOF %>
    <tr bgcolor="#FFFFFF" > 
       <td  align="center"><input name="id" type="checkbox" id="id" value="<%=rs("id")%>"></td>
      <td  align="center"><a href="<%=rs("SiteUrl")%>" target='blank' title="广告地址：<%=rs("SiteUrl") & vbcrlf %>广告简介：<%=vbcrlf & rs("SiteIntro")%>"><%=rs("SiteName")%></a></td>
      <td  align="center">
<%
if rs("ImgUrl")<>"无图片" then
 if rs("isflash")=true then
	Response.Write "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'"
	if rs("ImgWidth")>0 then 
		if rs("ImgWidth")<300 then
			response.write " width='" & rs("ImgWidth") & "'"
			if rs("ImgHeight")>0 then response.write " height='" & rs("ImgHeight") & "'"
		else
			response.write " width='300'"
			if rs("ImgHeight")>0 then response.write " height='" & Cint(300/rs("ImgWidth")*rs("ImgHeight")) & "'"
		end if
	end if
	response.write "><param name='movie' value='/" & rs("ImgUrl") & "'><param name='quality' value='high'><embed src='/" & rs("ImgUrl") & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'"
	if rs("ImgWidth")>0 then 
		if rs("ImgWidth")<300 then
			response.write " width='" & rs("ImgWidth") & "'"
			if rs("ImgHeight")>0 then response.write " height='" & rs("ImgHeight") & "'"
		else
			response.write " width='300'"
			if rs("ImgHeight")>0 then response.write " height='" & Cint(300/rs("ImgWidth")*rs("ImgHeight")) & "'"
		end if
	end if
	response.write "></embed></object>"
else
	response.write "<a href='" & rs("SiteUrl") & "' target='_blank' title='图片地址：" & rs("ImgUrl") & vbcrlf & "图片宽度：" & rs("ImgWidth") & "像素" & vbcrlf & "图片高度：" & rs("ImgHeight") & "像素'><img src=../" & rs("ImgUrl") & ""
	if rs("ImgWidth")>0 then 
		if rs("ImgWidth")<300 then
			response.write " width='" & rs("ImgWidth") & "'"
			if rs("ImgHeight")>0 then response.write " height='" & rs("ImgHeight") & "'"
		else
			response.write " width='300'"
			if rs("ImgHeight")>0 then response.write " height='" & Cint(300/rs("ImgWidth")*rs("ImgHeight")) & "'"
		end if
	end if
	response.write " border='0'></a>"
 end if
else
Response.write "无图片"
end if
%>
	  </td>
      <td  align="center">
<%
select case rs("advid")
case 1
	response.write "友情链接"
case 2
	response.write "广告二&nbsp;778*100"
case 3
	response.write "广告三&nbsp;125*60"
case 4
	response.write "广告四&nbsp;125*60"
case 5
	response.write "广告五&nbsp;778*100"
case 6
	response.write "广告六&nbsp;778*100"
case 7
	response.write "广告七&nbsp;浮动"
end select
%>
</td>
      <td  align="center"><%=rs("endtime")%></td>
      <td align="center"><a href="admin_advModify.asp?ID=<%=rs("ID")%>">修改</a>&nbsp;<a href="admin_advDel.asp?id=<%=rs("id")%>&page=<%=request("page")%>&action=<%=delart%>" onClick="return ConfirmDel();">删除</a> </td>
    </tr>
    <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
    <tr align="right" bgcolor="cccccc" class="tdbg"> 
      <td colspan="6">&nbsp;
        <input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">
        选中所有显示记录&nbsp;
        <input type=submit name=action onclick="{if(confirm('确定删除选定的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="删除广告" class="btn">
&nbsp;      </td>
    </tr></form>
</table>  

<%
end sub 
%>
<table width="95%" height="76" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#999999">
  <tr bgcolor="#CCCCCC">
    <td height="31" colspan="2"><span class="masterTitle">　　　广告管理 -&gt; 站点名搜索 （模糊查询） </span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td width="33%" height="35" align="center" bgcolor="#FFFFFF"><strong><span class="style1">站点名</span>关键字搜索：</strong> </td>
    <form name="form2" method="post" action="admin_advmanage.asp">
      <td width="67%"> 　　 　　　
          <input name="SiteName" type="text" class="Inpt" id="SiteName">
          <input name="Submit3" type="submit" class="btn" value="搜 索">
      </td>
    </form>
  </tr>
  <tr bgcolor="#CCCCCC">
    <td height="6" colspan="2"></td>
  </tr>
</table>
</body>
</html>
<%
rs.Close
set rs=Nothing
call CloseConn()  
%>