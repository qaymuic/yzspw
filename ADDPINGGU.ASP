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
.style2 {
	font-size: 14px;
	font-weight: bold;
	color: #FFFFFF;
}
.style3 {font-size: 14px}
.style4 {color: #000000}
.style5 {font-size: 14px; font-weight: bold; color: #000000; }
.style1 {color: #FF0000}
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
              <td width="180" valign="top"><!--#include file=wzleft.asp --></td>
              <td width="596" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="5">
                <tr>
                  <td>&nbsp;<a href="index.asp">返回首页</a>&nbsp;&gt;&gt;&gt;&nbsp;商铺评估 &gt;&gt;&gt; 我要评估 </td><td align="right" width="150">&nbsp;&nbsp;</td>
                </tr>
              </table>
                <TABLE WIDTH=576 BORDER=0 CELLPADDING=0 CELLSPACING=0>
                  <TR>
                    <TD><IMG SRC="images/ytiiNew_01.gif" WIDTH=576 ALT=""></TD>
                  </TR>
                  <TR>
                    <TD valign="top" background="images/ytiiNew_02.gif"> 　<br>
                        　 
                        <span class="style1">请认真填写下列表格，以便我们及时为您服务。 带*为必填项！</span><br>
                        <br>
                        <table width="97%"  border="0" align="center" cellpadding="6" cellspacing="0">
                          <form name="eeform1" method="post" action="jgpg_save.asp">
                            <tr>
                              <td width="21%"><strong>评估目的</strong></td>
                              <td width="79%">
                                <input name="a1" type="text" id="a1">
                                <span class="style1">*</span> </td>
                            </tr>
                            <tr>
                              <td><strong>评估对象描述</strong><br>
              (详细说明您的商铺位置，楼层，面积，建筑年份等)</td>
                              <td><textarea name="a2" cols="50" rows="10" id="a2"></textarea>
                                  <span class="style1">*</span></td>
                            </tr>
                            <tr>
                              <td><strong>出评估报告时间</strong></td>
                              <td><input name="a3" type="text" id="a3">
              (填上您要求出报告时间)</td>
                            </tr>
                            <tr>
                              <td><strong>联 系 人</strong></td>
                              <td><input name="a4" type="text" id="a4">
                                  <span class="style1">*</span></td>
                            </tr>
                            <tr>
                              <td><strong>电　　话</strong></td>
                              <td><input name="a5" type="text" id="a5">
                                  <span class="style1">*</span></td>
                            </tr>
                            <tr>
                              <td><strong>地　　址</strong></td>
                              <td><input name="a6" type="text" id="a6"></td>
                            </tr>
                            <tr>
                              <td><strong>电子邮件</strong></td>
                              <td><input name="a7" type="text" id="a7"></td>
                            </tr>
                            <tr>
                              <td>&nbsp;</td>
                              <td><input type="submit" name="Submit" value="提交"></td>
                            </tr>
                          </form>
                        </table>
                        <br>
                    </TD>
                  </TR>
                  <TR>
                    <TD> <IMG SRC="images/ytiiNew_03.gif" WIDTH=576 ALT=""></TD>
                  </TR>
                </TABLE></td>
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
closeconn
%>