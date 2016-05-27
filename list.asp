<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/function.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>扬州商铺网</title>
<link href="css/text.css" rel="stylesheet" type="text/css">
<SCRIPT language=Javascript>
function checkform(){
if (ListForm.Title.value=="") {
document.ListForm.Title.value = "无标题";
}
if (ListForm.AuthorName.value=="") {
document.ListForm.AuthorName.value = "GUEST";
}
if (ListForm.Content.value=="") {
alert("请填写评论内容");
document.ListForm.Content.focus();
return false;
}

}

</SCRIPT>
<style type="text/css">
<!--
.style5 {font-size: 14px}
-->
</style>
<script language="javascript">
<!--
function GetImgWH()
{
  var OriginImage=new Image();
  var oImg = document.all("ShowImg");
  if(OriginImage.src!=oImg.src)OriginImage.src=oImg.src;
  var Wth=OriginImage.width;
  var Hgh=OriginImage.height;
  var BaiFB;
  var i=100;
  while(Wth>250){
  		i=i-1;
  		BaiFB=i/100;
		Wth=Wth*BaiFB;
		Hgh=Hgh*BaiFB;
  }  
  //if(Wth>330)Wth=330;
  //if(Hgh>345)Hgh=345;
  oImg.width= Wth;
  oImg.height= Hgh;
}
//-->
</script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" class="table-body">
  <tr>
    <td><!--#include file=top.asp --><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
		<%
		dim id
		id=ReplaceBadChar(request("id"))
		if id<>"" and isnumeric(id) then
		sql="update ytiinews set hits=hits+1 where id=" & id & ""
		conn.execute sql
		sql="select * from ytiinews where id=" & id & ""
		Set rs= Server.CreateObject("ADODB.Recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			response.Write("<p>找不到内容</p>")
		else
		%>
          <td><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
            <tr bgcolor="#FFFFFF" class=title>
              <td height="18"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr><td>&nbsp;<a href="index.asp">返回首页</a>&nbsp;>>&nbsp;<%=rs("bigclassname")%></td><td align="right" width="150"><%if instr(rs("bigclassname"),"商铺评估")>0 then%>>> <a href="ADDPINGGU.ASP" target="_blank"><font color="#FF0000"><b>我要评估</b></font></a> <<<%end if%>&nbsp;&nbsp;</td></tr></table></td>
              </tr>
            <tr valign="middle" bgcolor="#FFFFFF">
              <td height="40" class=tdbg><div align="center">
                  <p class="TD"><strong><font color="#006600" size="3"> <%=rs("Title")%> </font></strong></p>
              </div></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td align="center" bgcolor="#F5F5F5" class=td-tianchong-4px> <font color="#666666">发布时间：<%=rs("UpdateTime")%></font> &nbsp;<font color="#666666">&nbsp;责任编辑：
                      <%if rs("Author")<>"" then
			   response.write rs("Author")
			  else
			   response.write "本站编辑"
			  end if%>
                  </font>                    </td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td valign="top" class="td-tianchong-4px"><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td valign="top" class="text-p">
					<% if rs("DefaultPicUrl")<>"" then 
					dim FileType,Str
					Str="jpg,gif,png,bmp"
					FileType=lcase(right(rs("DefaultPicUrl"),3))
					if instr(Str,FileType)>0 then
					%>					
					<table width="20"  border="0" align="left" cellpadding="10" cellspacing="0">
                      <tr>
                        <td><a href="<%=rs("DefaultPicUrl")%>" target="_blank"><img src="<%=rs("DefaultPicUrl")%>" border="0" class="img-border-1px" id="ShowImg"></a></td>
                      </tr>
                    </table><script language="javascript">GetImgWH()</script>
					<% 
					end if
					end if%>
                      <%=rs("content")%> </td>
                  </tr>
                  <tr>
                    <td height="23" align="right" bgcolor="#F5F5F5" class="td-tianchong-4px">【<a href="javascript:window.close();">关闭窗口</a>】</td>
                  </tr>
                  <tr>
                    <td align="left" class="td-tianchong-4px"><TABLE width=100% border=0 cellPadding=0 cellSpacing=0>
                      <TBODY>
                        <TR>
                          <TD background="images/art_pl3.gif"><SPAN class=HeadCell><IMG 
                  src="images/art_pl2.gif" 
                  width=458 height=43 border=0></SPAN></TD>
                          <TD width="20" align="right"><img src="images/art_pl4.gif" width="20" height="43"></TD>
                        </TR>
                      </TBODY>
                    </TABLE>
                      <TABLE width=100% border=0 cellPadding=0 
            cellSpacing=0 bgcolor="#FFFFFF" 
            style="BORDER-RIGHT: #c4c4c4 1px solid; BORDER-LEFT: #c4c4c4 1px solid; BORDER-BOTTOM: #c4c4c4 1px solid">
                        <FORM name=ListForm onsubmit="return checkform();" action="savepin.asp" method=post>
                          <TR>
                              <TD height="10" colspan="2"></TD>
                            </TR>
                            <TR bgcolor="#f5f5f5">
                              <TD width="15%" align=right>标题：</TD>
                              <TD width="85%" height=25 align=left><INPUT id=Title style="BORDER-RIGHT: #888888 1px solid; BORDER-TOP: #888888 1px solid; BORDER-LEFT: #888888 1px solid; BORDER-BOTTOM: #888888 1px solid" maxLength=40 size=30 name=Title>
                                  <input name="id" type="hidden" id="id" value="<%=id%>">
                                  <input name="atitle" type="hidden" id="atitle" value="<%=rs("title")%>"></TD>
                            </TR>
                            <TR bgcolor="#f5f5f5">
                              <TD align=right>姓名：</TD>
                              <TD width="85%" height=25 align=left><INPUT id=AuthorName style="BORDER-RIGHT: #888888 1px solid; BORDER-TOP: #888888 1px solid; BORDER-LEFT: #888888 1px solid; BORDER-BOTTOM: #888888 1px solid" maxLength=40 size=30 name=AuthorName ></TD>
                            </TR>
                            <TR bgcolor="#f5f5f5">
                              <TD align=right valign="top">内容：</TD>
                              <TD width="85%" align="left"><TEXTAREA id=Content style="BORDER-RIGHT: #888888 1px solid; BORDER-TOP: #888888 1px solid; BORDER-LEFT: #888888 1px solid; BORDER-BOTTOM: #888888 1px solid" name=Content rows=3 cols=45></TEXTAREA>                              <br>
                              &nbsp;</TD>
                            </TR>
                            <TR>
                              <TD align=middle colSpan=2><TABLE width="100%" height=25 
                  border=0 cellPadding=0 cellSpacing=0 bgcolor="#E6E6E6">
                                <TR>
                                    <TD width="39%" height="35" align=right><!--<INPUT type=image height=21 width=57  src="images/form_submit.gif" border=0 name=imageField>-->
									<input style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BACKGROUND-IMAGE: url(images/form_submit.gif); BORDER-LEFT: 0px; WIDTH: 57px; BORDER-BOTTOM: 0px; HEIGHT: 21px" type=Submit value="　 " name=imageField ;>
									</TD>
                                    <TD width="61%" align=left>&nbsp;<input style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BACKGROUND-IMAGE: url(images/form_reset.gif); BORDER-LEFT: 0px; WIDTH: 57px; BORDER-BOTTOM: 0px; HEIGHT: 21px" type=reset value="　 " name=Submit2 ;></TD>
                                </TR>
                              </TABLE></TD>
                            </TR>
                          </FORM>
                      </TABLE>
                      </td>
                  </tr>
				  <tr>
				    <td class="td-tianchong-4px"><br>&nbsp;相关评论：
				      <hr size="1" noshade></td>
				  </tr>
				  <tr>
				    <td class="td-tianchong-4px"><%
	    Set rs=Server.CreateObject("Adodb.RecordSet")
		sql="select * from article_pin where article_id="&id
		rs.Open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "&nbsp;&nbsp;暂未有评论!"
        else
      %>
                      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
                        <% do while not rs.eof %>
                        <tr>
                          <td height="22" bgcolor="#F5F5F5" class="TD-MENU"><strong>　标题</strong>：<%=rs("title")%>&nbsp;&nbsp;&nbsp;<strong>评论者</strong>：<%=rs("Name")%><strong>　评论时间</strong>：<%=rs("addtime")%></td>
                        </tr>
                        <tr>
                          <td bgcolor="#FFFFFF" class=newstitle><table width="95%" border="0" align="center" cellpadding="0" cellspacing="10">
                              <tr>
                                <td><%=rs("Content")%></td>
                              </tr>
                          </table></td>
                        </tr>
                        <%
  rs.movenext
  loop
  %>
                      </table>
                      <%end if%></td>
				    </tr>
				  <tr>
				    <td class="td-tianchong-4px">&nbsp;</td>
				    </tr>
              </table></td>
            </tr>
          </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
      <tr>
        <td align="center"><!--#include file=foot.asp --></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<%
end if
rs.close
set rs=nothing
call CloseConn()
else
%>
<script language=javascript>
history.back()
alert("没有相关内容!")
</script>
<%end if%>