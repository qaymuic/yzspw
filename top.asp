<!--
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/text.css" rel="stylesheet" type="text/css">
//-->
<%
dim rs,sql
conn.execute("update shu set hits=hits+1")
sql="select * from shu"
Set rs= Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
dim shu
shu=rs("hits")
rs.close
%>
<table width="778" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="20" background="images/topbg00.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="50%" align="left" class="TD-MENU"><font color="#FFFFFF"><script language=JavaScript class="font_white">
 today=new Date();
 function initArray(){
   this.length=initArray.arguments.length
   for(var i=0;i<this.length;i++)
   this[i+1]=initArray.arguments[i]  }
   var d=new initArray(
     "星期日",
     "星期一",
     "星期二",
     "星期三",
     "星期四",
     "星期五",
     "星期六");
document.write(
     "<font color=#ff0000 style='font-size:10pt;font-family: 宋体'> ",
     today.getYear(),"年",
     today.getMonth()+1,"月",
     today.getDate(),"日","　",
     d[today.getDay()+1],
     "</font>" );
</script></font>&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">本站已被访问&nbsp;<span class="a_color_003"><b><%=shu%></b></span>&nbsp;次</font></td>
          <td width="50%" align="right" class="TD-MENU"><a href="http://www.yzhuiyu.com/aboutus.asp" target=_blank class="a_color_003">联系我们</a>&nbsp;<font color="#FFFF00">|</font>&nbsp;<a href="ads.asp" class="a_color_003">刊登广告</a>&nbsp;<font color="#FFFF00">|</font>&nbsp;<a href="bbs/index.asp" target="_blank" class="a_color_003">商铺论坛</a></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" height="80"  border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="250"><img src="images/LOGO.gif" width="250" height="70"></td>
          <td align="center"><table width="100%" height="70"  border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td align="center"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="520" height="70">
                    <param name="movie" value="banner.swf">
                    <param name="quality" value="high">
                    <param name="menu" value="false">
                    <embed src="banner.swf" width="520" height="70" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" menu="false"></embed>
                </object></td>
              </tr>
          </table></td>
        </tr>
      </table>
        <table width="100%"  border="0" cellpadding="0" cellspacing="1" bgcolor="#ABA5FC">
          <tr>
            <td width="14%" align="center" bgcolor="#253E7C" class="TD-MENU"><a href="index.asp" class="a_menu">返回首页</a></td>
            <td width="14%" align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="splist.asp">商铺展示</A></td>
            <td width="14%" align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="splist1.asp">商铺租赁</A></td>
            <td width="14%" align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="splist2.asp">商铺买卖</A></td>
            <td width="14%" align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="pinggulist.asp">商铺评估</A></td>
            <td width="14%" align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="daikuanlist.asp">商铺贷款</A></td>
            <td width="14%" align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="paimailist.asp">商铺拍卖</A></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#253E7C" class="TD-MENU"><a href="zichanlist.asp" class="a_menu">处置资产</a></td>
            <td align="center" bgcolor="#253E7C" class="TD-MENU"><a href="yuanqulist.asp" class="a_menu">园区招商</a></td>
            <td align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="splist3.asp">工业厂房</A></td>
            <td align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="xuetanglist.asp">商铺学堂</A></td>
            <td align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="chuangyelist.asp">创业中心</A></td>
            <td align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="zhuangshilist.asp">名店风采</A></td>
            <td align="center" bgcolor="#253E7C" class="TD-MENU"><A class=a_menu href="splist4.asp">周边商铺</A></td>
          </tr>
        </table>
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="50" align="right" background="images/sousuo.gif"><table width="83%"  border="0" cellspacing="0" cellpadding="0">
                <form name="Sform1" method="post" action="spsearch.asp">
                  <tr>
                    <td align="left"><%
	   Set rs=Server.CreateObject("Adodb.RecordSet")
           sql = "select * from Special order by Specialid"
           rs.open sql,conn,1,1
		%> <select name="splb" size="1">
                    <option value="" selected>所有类别</option>
		    <%do while not rs.eof%>
                <option value="<%=trim(rs("SpecialName"))%>"><%=trim(rs("SpecialName"))%></option>
            <%
		     rs.movenext
    	     loop
             rs.close:set rs=nothing
			%></select>
                        <select name="spgqlb" id="spgqlb">
        <option value="" selected>所有交易类型</option>
        <option value="出租">出租</option>
        <option value="求租">求租</option>
        <option value="转让">转让</option>
        <option value="求购">求购</option>
        <option value="出售">出售</option>
      </select>
                        <%
	   Set rs=Server.CreateObject("Adodb.RecordSet")
           sql = "select * from a2"
           rs.open sql,conn,1,1
		%> <select name="SmallClassName" size="1">
                    <option value="" selected>所有地区</option>
		    <%do while not rs.eof%>
                <option value="<%=trim(rs("SmallClassName"))%>"><%=trim(rs("SmallClassName"))%></option>
            <%
		     rs.movenext
    	     loop
             rs.close:set rs=nothing
			%></select>
                    关键字
                    <input name="spname" type="text" size="15">
                    <input type="submit" name="Submit" value="搜 索"></td><td width="110" align="center"><a href="sppost.asp" target="_blank"><img src="images/addnew.gif" border="0"></a></td>
                  </tr>
                </form>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table>
