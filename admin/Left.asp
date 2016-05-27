<!--<meta http-equiv="Content-Type" content="text/html; charset=gb2312">-->
<%
session.timeout=100
if session("admin")="" then
	response.write "请重新登录！"
else
%>
<!--#include file="conn.asp"-->
<!--#include file="clock.asp"-->
<html>
<head>
<title>管理导航</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
body  { background:#F5FEED; margin:0px; font:9pt 宋体; }
table  { border:0px; }
td  { font:normal 12px 宋体; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px 宋体; color:#000000; text-decoration:none; line-height: 18px;  }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#F5FEED; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
.tableBody1
{
width:97%;
border: 1px; 
background-color: #ffffff;
}
</style>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
</head>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=100% cellpadding=0 cellspacing=0 border=0 align=left>
    <tr><td valign=top>
<table width=158 border="0" align=center cellpadding=0 cellspacing=0>
  <tr>
        <td height=38 valign=bottom background="images/title1.gif"></td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td style="display:" id='submenu0'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20>操作人：<%=session("realname")%></td>
              </tr>
<tr>
                <td height=20>权　限：<font color=red><%
		  select case session("purview")
		    case 1
              strPurview="管理员"
            case 2
              strpurview="部门主管"
            case 3
			  strpurview="一般用户"
		  end select
		  response.write(strPurview)
         %>
                  </font></td>
              </tr>
</table>
</div>
	</td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=37 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/a1.gif" id=menuTitle3 onclick="showsubmenu(3)" style="cursor:hand;"> 
          <span>&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;</td>
  </tr>
  <tr>
    <td style="display:" id='submenu3'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr> 
                <td height=20 align=center> 
                <br><a href="ArticleAdd.asp" target=right>添加信息</a><br>
                  <a href="ArticleManage.asp" target=right>信息管理</a><br>
				  <a href="ClassManage.asp" target=right>信息管理</a><br>
                  <hr width="100%" size="1">
				  <a href="newsManage.asp" target=right>商铺动态管理</a><br>
                  <a href="sp_add.asp" target="right"><font color=red>商铺增加</font></a> <br>
                  <a href="sp_Manage.asp" target="right"><font color=red>商铺管理</font></a> <br>
                  <a href="ClassManage1.asp" target="right">商铺地区分类</a> <br>
                  <hr width="100%" size="1">
                  <a href="fangjgpg.asp" target="right">商铺评估</a> <br>
                  <hr width="100%" size="1">
                   <a href="admin_AdvManage.asp" target="right">广告管理</a> <br>
                   <hr width="100%" size="1">
                  <a href="admin_replylist.asp" target="right">查看评论</a> <br>
                  <br>
                </td>
              </tr>
            </table>
</div>
	</td>
  </tr>
</table>

<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
        <td height=37 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/a2.gif" id=menuTitle2 onclick="showsubmenu(2)" style="cursor:hand;"> 
          <span>&nbsp;&nbsp;&nbsp;&nbsp;</span> </td>
  </tr>
  <tr>
    <td style="display:" id='submenu2'>
<div class=sec_menu style="width:158"><table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20 align=center><span><a href="UserManage.asp" target="right">用户管理</a></span><br>
                  <%if session("purview")=1 then%>
                  <a href="AdminManage.asp" target="right">管理员管理</a><br>
                  <%end if%>
                  <span><a href=logout.asp target=_top><b><font color="#990000">退出系统</font></b></a></span></td>
              </tr>
</table> </div>
</td>
  </tr>
</table>

    <br>
    <table width="158" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><TABLE border=0 width="110" align=center cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
            <tr>
              <td bgcolor="#CCCCCC" width="107"> <TABLE border=0 width="100%" height="100%" cellpadding="0" cellspacing="1" bgcolor="#0099FF">
                  <tr> 
                    <td> <TABLE border=0 width="100%" height="100%" cellpadding="3" cellspacing="0">
                        <TR> 
                          <td class=tableBody1 align="center"><img src="images/clock/cb.gif" name="a"><img src="images/clock/cb.gif" name="b"><img src="images/clock/colon.gif" name="c"><img src="images/clock/cb.gif" name="d"><img src="images/clock/cb.gif" name="e"><img src="images/clock/colon.gif" name="f"><img src="images/clock/cb.gif" name="g"><img src="images/clock/cb.gif" name="h"></td>
                        </TR>
                        <TR> 
                          <td class=tableBody1 align="center" id="tdyear"></td>
                        </TR>
                        <TR> 
                          <td class=tableBody1 align="center" id="tdweek"></td>
                        </TR>
                        <TR> 
                          <td class=tableBody1 align="center" id="tdcyear"></td>
                        </TR>
                      </TABLE></td>
                  </tr>
                </TABLE></td>
            </tr>

          </table>
          <div align="center">
            <script language='javascript'>Chen_CAL();</script>
          </div></td>
      </tr>
    </table></body>
</html>
<%
end if
%>