<!--#include file=inc/conn.asp -->
<!--#include file=inc/function.asp -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>扬州商铺网</title>
<link href="css/text.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.style1 {color: #990000}
.style2 {
	font-size: 14px;
	font-weight: bold;
	color: #FFFFFF;
}
.style3 {font-size: 14px}
-->
</style>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" class="table-body">
  <tr>
    <td><!--#include file=top.asp --><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table width="776"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="180" valign="top"><table width="89%"  border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                <tr>
                  <td bgcolor="#0066CC"><div align="center" class="style2">商铺类型</div></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF"><div align="center"><a href="splist.asp?spgqlb=&#36716;&#35753;">商铺转让</a></div></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF"><div align="center"><a href="splist.asp?spgqlb=&#20986;&#31199;">商铺出租</a></div></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF"> <div align="center"><a href="splist.asp?spgqlb=&#20986;&#21806;">商铺出售</a></div></td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#FFFFFF"><a href="splist.asp?spgqlb=&#27714;&#31199;">商铺求租</a></td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#FFFFFF"><a href="splist.asp?spgqlb=&#27714;&#36141;">商铺求购</a></td>
                </tr>
              </table>
                <br>
                <table width="89%"  border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                  <tr>
                    <td bgcolor="#0066CC"><div align="center" class="style2">商铺类型</div></td>
                  </tr>
				  <%
		   dim rs1,sql1
	       Set rs1=Server.CreateObject("Adodb.RecordSet")
           sql1 = "select * from Special order by Specialid"
           rs1.open sql1,conn,1,1
		   do while not rs1.eof
		   %>
            <tr>
                    <td bgcolor="#FFFFFF"><div align="center"><a href="splist.asp?splb=<%=rs1("SpecialName")%>"><%=trim(rs1("SpecialName"))%></a></div></td>
            </tr><%
		     rs1.movenext
    	     loop
             rs1.close:set rs1=nothing
			%>
                 </table></td>
              <td width="596" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="5">
                <tr>
                      <td>&nbsp;<a href="index.asp">首页</a>&nbsp;&gt;&gt;&gt;&nbsp;广告刊登</td>
                </tr>
              </table>
                <table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
                <tr>
                      <td width="100%" height="6" valign="top" bgcolor="#FFFFFF" class="text-p"><p align="center"><b><font color="#1F641F" style="font-size:16pt">广 
                          告 服 务<br>
                          <img src="images/zoulang_main_line.gif" width="513" height="1"></font></b></p>
                        　　网络广告的独特优势：<br>
                        　　<strong>* 广泛性：</strong> 24小时连续播出，任何人在任何地方均可随意在线浏览<br>
                        　　<strong>* 节省性：</strong> 收费低廉节约成本，随时更改广告内容，绝无资金浪费<br>
                        　　<strong>* 互动性：</strong> 受众主动点击想了解的信息，商家在线查询得到反馈信息<br>
                        　　<strong>* 目标性：</strong> 不同广告内容针对不同受众，通过点击直达可能用户<br>
                        　　<strong>* 计量性：</strong> 精准统计浏览量，受众群体清晰易辨，广告效果立显<br>
                        　　<strong>* 感官性：</strong> 图文声像多种形式应用，多媒体技术令人身临其境 
                        <p>　　互联网广告答疑 </p>
                        <p>　　<font color="#0000FF">1、网络广告能为我的产品和服务做到什麽？</font><br>
                          　　网络广告运用新兴的媒体与自由灵活的广告形式和空间，为广告主带来庞大的访问量，使品牌记忆度大大增强。 
                          网络广告更可按广告主需求指向特定用户群，从而把销售信息直接传递给目标受众，有效刺激产品及服务的销售增长。 
                          <br>
                          <br>
                          　　<font color="#0000FF">2、网络广告为什麽对我的产品和服务产生好的广告效果？</font><br>
                          　　网络广告依托网络在线的强大优势，根据客户产品和服务的特性，与相关频 道或服务结合投放广告，或举办特色在线营销活动，最大程度的吸引、影响目标人群。网络广告具有其其特殊的视觉性和互动性，图、文、声、像，活跃的表现令 
                          广告投资效果迅速显现，吸引消费者对所感兴趣的广告进行点击，深层了解 ，使网络广告达到最佳效果。<br>
                          <br>
                          　　<font color="#0000FF">3、网络是如何为我提供广告服务的？</font><br>
                          　　网络通过先进科技、低廉成本、便捷方式、高效收益为广告主提供专业服务。网络可以把丰富的广告内容带给特定的目标受众，广告主亦可通过在线调查做出详尽的用户分析，大大降低用户获取成本。<br>
                          <br>
                          　　<font color="#0000FF">4、网络广告的优势何在？</font><br>
                          　　网络广告的优势很多比如网络广告按访问量多少收取费用，并且可以根据客户产品、服务特点向 相关目标人群投放广告，让广告主的每一份投入都物有所值，因此具有最大的投资回报率。网络广告采用多媒体形式，图文并茂，感官性强，给广告主一个独立完整的空间展示产品与服务。网络广告进行中可随时生成监测报告，从而量化广告成本和效益，为今後的 
                          广告与营销活动提供依据<br>
                          <br>
                          　　<font color="#0000FF">5、网络广告的费用如何计算？</font><br>
                          　　网络广告的费用较传统广告低廉，而且投资回报率高。目前，网络广告的计费方式可按广告每被点击一次收取费用，也可按广告每被访问千次收取费用；或者按照投放时间、投放位置收取费用；另外，横幅广告和按钮广告的费用也有所不同。<br>
                          <br>
                          　　欢迎在本站投放广告，请和工作人员联系。<br>
                          <br>
                          　　广告联系：<br>
                          <br>
                          　　　　地址：文汇南路88号开发大厦房地产交易中心<br>
                          <br>
                          　　　　电话：7892731　<br>
                          <br>
                          　　　　联系人：居经理</p>
                      </td>
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
