<!--#include file=inc/conn.asp -->
<!--#include file=inc/function.asp -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����������</title>
<link href="css/text.css" rel="stylesheet" type="text/css">

<style type="text/css">
<!--
.style11 {
	color: #FFFFFF;
	font-weight: bold;
	font-size: 14px;
}
-->
</style>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" class="table-body">
  <tr>
    <td><!--#include file=top.asp --><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
        <tr>
          <td bgcolor="#FFFFFF"><table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="100%" height="6"><FORM name='UserReg' action='userregsave.asp' method='post'>
                  <table width=100% height="556" border=0 align="center" cellpadding=0 cellspacing=1 bgcolor="#CCCCCC" class="border" style="border-collapse: collapse">
                    <TR align=center bgcolor="#0066CC" class='title'>
                      <TD height=22 colSpan=2><span class="style11"><font style10>�û�ע�����д</font></span></TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">�� �� ����<BR>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <input   maxlength=14 size=10 name=UserName style="border:1 solid;background-color:#F8FEF1" >
                          <font color="#FF0000">*</font> ���ܳ���14���ַ���7�����֣�</TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">�ܡ����룺<BR>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <INPUT   type=password maxLength=12 size=10 name=Password style="border:1 solid;background-color:#F8FEF1" >
                          <font color="#FF0000">*</font> ����6λ</TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">ȷ�����룺<BR>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <INPUT   type=password maxLength=12 size=20 name=PwdConfirm style="border:1 solid;background-color:#F8FEF1" >
                          <font color="#FF0000">*</font> ������һ��ȷ��</TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">�������⣺<BR>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <INPUT   type=text maxLength=50 size=20 name="Question" style="border:1 solid;background-color:#F8FEF1" >
                          <font color="#FF0000">*</font> �����������ʾ����</TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">����𰸣�<BR>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <INPUT   type=text maxLength=20 size=20 name="Answer" style="border:1 solid;background-color:#F8FEF1" >
                          <font color="#FF0000">* </font>�����������ʾ����𰸣�����ȡ������ </TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">�ա�������<BR>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <input name="contact"   type=text id="contact2" size=10 maxlength=20 style="border:1 solid;background-color:#F8FEF1" >
                      </TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">�ء���ַ��<BR>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <INPUT name="address"   type=text id="address" size=40 maxLength=20 style="border:1 solid;background-color:#F8FEF1" >
                      </TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">Email��<BR>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <input   maxlength=50 size=20 name=Email style="border:1 solid;background-color:#F8FEF1" >
                          <font color="#FF0000">*</font> ��ʽ<strong>:invest@muicc.com</strong></TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">��˾���ƣ�<BR>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <input name="company"   type=text id="company4" size=40 maxlength=50 style="border:1 solid;background-color:#F8FEF1" >
                      </TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">�硡������<BR>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <INPUT   maxLength=100 size=20 name=tel style="border:1 solid;background-color:#F8FEF1" ></TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%"><div align="right">�������棺<br>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <INPUT maxLength=20 size=20 name=fax style="border:1 solid;background-color:#F8FEF1" >
                      </TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%" height="38"><div align="right">�ʡ����ࣺ<br>
                      </div></TD>
                      <TD width="83%">&nbsp;
                          <INPUT maxLength=50 size=10 name=pc style="border:1 solid;background-color:#F8FEF1" ></TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD width="17%" valign="top"><div align="right">��˾��飺</div></TD>
                      <TD>&nbsp;
                          <textarea name="content" cols="40" rows="5" id="content" style="border:1 solid;background-color:#F8FEF1" ></textarea></TD>
                    </TR>
                    <TR bgcolor="#FFFFFF" class="tdbg" >
                      <TD colspan="2"><div align="right"></div>
                          <div align="center">
                            <input   type=submit value=" �� �� " name=Submit style="border:1 solid;background-color:#F8FEF1" >
&nbsp;
                <input name=Reset   type=reset id="Reset2" value=" �� �� " style="border:1 solid;background-color:#F8FEF1" >
                        </div></TD>
                    </TR>
                  </TABLE>
                  <div align="center"> </div>
              </form></td>
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