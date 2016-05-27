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
.style3 {font-size: 14px}
.style4 {
	color: #FFFFFF;
	font-size: 14px;
	font-weight: bold;
}
.style5 {font-size: 14px; color: #990000; }
.style6 {
	font-size: 12px;
	font-weight: bold;
	color: #990000;
}
.style8 {font-size: 12px; color: #990000; }
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
 // while(Wth>330 || Hgh>345){
 // 		i=i-1;
 // 		BaiFB=i/100;
//		Wth=Wth*BaiFB;
//		Hgh=Hgh*BaiFB;
 // }  
  if(Wth>330)Wth=330;
  if(Hgh>345)Hgh=345;
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
          <td><table width="776"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="10" valign="top"><%
					dim id
					id=ReplaceBadChar(request("id"))
					conn.execute "update spw set sphits=sphits+1 where id="&id
					Set rs=Server.CreateObject("Adodb.RecordSet")
					sql="select * from spw where id="&id
					'sql=sql & " order by id desc"
					rs.Open sql,conn,1,1
					  %>
                </td><td width="766" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="5">
                <tr>
                  <td>&nbsp;<a href="index.asp">返回首页</a>&nbsp;&gt;&gt;&gt;&nbsp;商铺展示&nbsp;&gt;&gt;&gt;&nbsp;<%=rs("splb")%></td>
                </tr>
              </table>
                <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
                <tr>
                  <td width="100%" height="6" valign="top" bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="3" cellspacing="2">
                      <tr bgcolor="#0066CC">
                      <td bgcolor="#0066CC"><div align="center" class="style4">商铺图片</div></td>
                      <td><div align="center"><span class="style4">商铺详细资料</span></div></td>
                      </tr>
					<tr>
                      <td width="45%" align="center" valign="middle"><%if rs("spphoto")<>"" then%><a href="<%=rs("spphoto")%>" target="_blank"><img src="<%=rs("spphoto")%>" class="img-border-1px" id="ShowImg" width="330"></a><%else%>暂无图片<%end if%></td>
                      <td width="55%" height="80" valign="top" bgcolor="#E3EDF4" class="td-tianchong-4px"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
                        <tr bgcolor="#FFFFFF">
                          <td width="31%" align="center" bgcolor="#FFFFFF"><span class="style3">名称</span></td>
                          <td width="31%" bgcolor="#FFFFFF"><%=rs("spname")%></td>
                          <td width="38%" bgcolor="#FFFFFF"><div align="right"><span class="style8">有效期:<%=rs("spendtime")%></span></div></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="center" bgcolor="#FFFFFF"><span class="style3">类型</span></td>
                          <td colspan="2"><span class="style6"><%=rs("spgqlb")%></span></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="center" bgcolor="#FFFFFF"><span class="style3">价格</span></td>
                          <td colspan="2"><%=rs("spjg")%> 万</td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="center" bgcolor="#FFFFFF"><span class="style3">面积</span></td>
                          <td colspan="2"><%=rs("spmj")%> 平方米</td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="center" bgcolor="#FFFFFF"><span class="style3">地区</span></td>
                          <td colspan="2"><%=rs("SmallClassName")%></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="center" bgcolor="#FFFFFF"><span class="style3">地址</span></td>
                          <td colspan="2"><%=rs("spaddress")%></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="center" bgcolor="#FFFFFF"><span class="style3">登记日期</span></td>
                          <td colspan="2"><%=rs("spaddtime")%></td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="center" bgcolor="#FFFFFF"><span class="style3">联系人</span></td>
                          <td colspan="2"><!--< %=rs("spren")% >-->居先生</td>
                        </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="center" bgcolor="#FFFFFF"><span class="style3">电话</span></td>
                          <td colspan="2"><!--< %=rs("spcontact")% >-->7017847  13328120965</td>
                        </tr>
                        <tr bgcolor="#0066CC">
                          <td colspan="3"><div align="center"><span class="style4">商铺其它说明</span></div></td>
                          </tr>
                        <tr bgcolor="#FFFFFF">
                          <td colspan="3"><%=rs("spcontent")%></td>
                          </tr>
                        <tr bgcolor="#FFFFFF">
                          <td align="center" bgcolor="#FFFFFF"><span class="style3">网站电话:</span></td>
                          <td colspan="2" bgcolor="#FFFFFF"><span class="style5">7892731</span> 　　　　　　　<span class="style3"> 点击次数:<%=rs("sphits")%></span></td>
                        </tr>
                      </table>                        
                        </td>
                    </tr>

                  </table></td>
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
rs.close
closeconn
%>