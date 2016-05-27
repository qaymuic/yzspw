<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>

<body>
<FORM name='UserReg' action='Useradd1save.asp' method='post'>
		
  <table width=88% border=0 align="center" cellpadding=2 cellspacing=1 bordercolor="#FFFFFF" style="border-collapse: collapse" class="border">
    <TR align=center class='title'> 
      <TD height=20 colSpan=2><font class=en><b>新用户加入</b></font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><b>用户名：</b><BR> </TD>
      <TD width="70%"> <INPUT   maxLength=14 size=10 name=UserName> </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><B>密码(至少6位)：</B><BR> </TD>
      <TD width="70%"> <INPUT   type=password maxLength=12 size=10 name=Password> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>确认密码(至少6位)：</strong><BR> </TD>
      <TD width="70%"> <INPUT   type=password maxLength=12 size=10 name=PwdConfirm> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>申请商：</strong><BR> </TD>
      <TD width="70%"> <INPUT name="company"   type=text id="company" size=40 maxLength=50> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>地址：</strong><BR> </TD>
      <TD width="70%"> <INPUT name="address"   type=text id="address" size=40 maxLength=20> 
      </TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>联系人：</strong><BR> </TD>
      <TD width="70%"><INPUT name="contact"   type=text id="contact" size=10 maxLength=20></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>Email地址：</strong><BR> </TD>
      <TD width="70%"> <INPUT   maxLength=50 size=30 name=Email> <font color="#FF0000">*</font></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>电话：</strong><BR> </TD>
      <TD width="70%"> <INPUT   maxLength=100 size=20 name=tel></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>传真：</strong><br> </TD>
      <TD width="70%"> <INPUT maxLength=20 size=20 name=fax></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%"><strong>邮编：</strong><br> </TD>
      <TD width="70%"> <INPUT maxLength=50 size=10 name=pc></TD>
    </TR>
    <TR class="tdbg" >
      <TD width="30%"><strong>简介：</strong></TD>
      <TD><textarea name="content" cols="40" rows="5" id="content"></textarea></TD>
    </TR>
    <TR class="tdbg" > 
      <TD width="30%">&nbsp;</TD>
      <TD><input   type=submit value=" 加 入 " name=Submit> &nbsp; <input name=Reset   type=reset id="Reset2" value=" 重 来 "> 
      </TD>
    </TR>
  </TABLE>
	
  <div align="center"> </div>
</form>
</body>
</html>
