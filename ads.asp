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
                  <td bgcolor="#0066CC"><div align="center" class="style2">��������</div></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF"><div align="center"><a href="splist.asp?spgqlb=&#36716;&#35753;">����ת��</a></div></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF"><div align="center"><a href="splist.asp?spgqlb=&#20986;&#31199;">���̳���</a></div></td>
                </tr>
                <tr>
                  <td bgcolor="#FFFFFF"> <div align="center"><a href="splist.asp?spgqlb=&#20986;&#21806;">���̳���</a></div></td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#FFFFFF"><a href="splist.asp?spgqlb=&#27714;&#31199;">��������</a></td>
                </tr>
                <tr>
                  <td align="center" bgcolor="#FFFFFF"><a href="splist.asp?spgqlb=&#27714;&#36141;">������</a></td>
                </tr>
              </table>
                <br>
                <table width="89%"  border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                  <tr>
                    <td bgcolor="#0066CC"><div align="center" class="style2">��������</div></td>
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
                      <td>&nbsp;<a href="index.asp">��ҳ</a>&nbsp;&gt;&gt;&gt;&nbsp;��濯��</td>
                </tr>
              </table>
                <table width="97%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
                <tr>
                      <td width="100%" height="6" valign="top" bgcolor="#FFFFFF" class="text-p"><p align="center"><b><font color="#1F641F" style="font-size:16pt">�� 
                          �� �� ��<br>
                          <img src="images/zoulang_main_line.gif" width="513" height="1"></font></b></p>
                        ����������Ķ������ƣ�<br>
                        ����<strong>* �㷺�ԣ�</strong> 24Сʱ�����������κ������κεط����������������<br>
                        ����<strong>* ��ʡ�ԣ�</strong> �շѵ�����Լ�ɱ�����ʱ���Ĺ�����ݣ������ʽ��˷�<br>
                        ����<strong>* �����ԣ�</strong> ��������������˽����Ϣ���̼����߲�ѯ�õ�������Ϣ<br>
                        ����<strong>* Ŀ���ԣ�</strong> ��ͬ���������Բ�ͬ���ڣ�ͨ�����ֱ������û�<br>
                        ����<strong>* �����ԣ�</strong> ��׼ͳ�������������Ⱥ�������ױ棬���Ч������<br>
                        ����<strong>* �й��ԣ�</strong> ͼ�����������ʽӦ�ã���ý�弼�����������侳 
                        <p>���������������� </p>
                        <p>����<font color="#0000FF">1����������Ϊ�ҵĲ�Ʒ�ͷ�������ʲ�᣿</font><br>
                          �����������������˵�ý�����������Ĺ����ʽ�Ϳռ䣬Ϊ����������Ӵ�ķ�������ʹƷ�Ƽ���ȴ����ǿ�� 
                          ��������ɰ����������ָ���ض��û�Ⱥ���Ӷ���������Ϣֱ�Ӵ��ݸ�Ŀ�����ڣ���Ч�̼���Ʒ����������������� 
                          <br>
                          <br>
                          ����<font color="#0000FF">2��������Ϊʲ����ҵĲ�Ʒ�ͷ�������õĹ��Ч����</font><br>
                          ���������������������ߵ�ǿ�����ƣ����ݿͻ���Ʒ�ͷ�������ԣ������Ƶ ���������Ͷ�Ź�棬��ٰ���ɫ����Ӫ��������̶ȵ�������Ӱ��Ŀ����Ⱥ���������������������Ӿ��Ժͻ����ԣ�ͼ���ġ������񣬻�Ծ�ı����� 
                          ���Ͷ��Ч��Ѹ�����֣����������߶�������Ȥ�Ĺ����е��������˽� ��ʹ������ﵽ���Ч����<br>
                          <br>
                          ����<font color="#0000FF">3�����������Ϊ���ṩ������ģ�</font><br>
                          ��������ͨ���Ƚ��Ƽ��������ɱ�����ݷ�ʽ����Ч����Ϊ������ṩרҵ����������԰ѷḻ�Ĺ�����ݴ����ض���Ŀ�����ڣ���������ͨ�����ߵ��������꾡���û���������󽵵��û���ȡ�ɱ���<br>
                          <br>
                          ����<font color="#0000FF">4������������ƺ��ڣ�</font><br>
                          ��������������ƺܶ���������水������������ȡ���ã����ҿ��Ը��ݿͻ���Ʒ�������ص��� ���Ŀ����ȺͶ�Ź�棬�ù������ÿһ��Ͷ�붼������ֵ����˾�������Ͷ�ʻر��ʡ���������ö�ý����ʽ��ͼ�Ĳ�ï���й���ǿ���������һ�����������Ŀռ�չʾ��Ʒ���������������п���ʱ���ɼ�ⱨ�棬�Ӷ��������ɱ���Ч�棬Ϊ����� 
                          �����Ӫ����ṩ����<br>
                          <br>
                          ����<font color="#0000FF">5��������ķ�����μ��㣿</font><br>
                          ����������ķ��ýϴ�ͳ������������Ͷ�ʻر��ʸߡ�Ŀǰ��������ļƷѷ�ʽ�ɰ����ÿ�����һ����ȡ���ã�Ҳ�ɰ����ÿ������ǧ����ȡ���ã����߰���Ͷ��ʱ�䡢Ͷ��λ����ȡ���ã����⣬������Ͱ�ť���ķ���Ҳ������ͬ��<br>
                          <br>
                          ������ӭ�ڱ�վͶ�Ź�棬��͹�����Ա��ϵ��<br>
                          <br>
                          ���������ϵ��<br>
                          <br>
                          ����������ַ���Ļ���·88�ſ������÷��ز���������<br>
                          <br>
                          ���������绰��7892731��<br>
                          <br>
                          ����������ϵ�ˣ��Ӿ���</p>
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
