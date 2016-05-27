<html>
<head>
<title>管理员登录</title>
<link rel="stylesheet" href="style.CSS">
<script language=javascript>
function SetFocus()
{
if (document.Login.UserName.value=="")
	document.Login.UserName.focus();
else
	document.Login.UserName.select();
}
function CheckForm()
{
	if(document.Login.UserName.value=="")
	{
		alert("请输入用户名！");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("请输入密码！");
		document.Login.Password.focus();
		return false;
	}
}
</script>
</head>
<body onLoad="SetFocus();">
<p>&nbsp;</p>
<form name="Login" action="ChkLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
    
  <table width="340" border="0" align="center" cellpadding="5" cellspacing="0" class="border" >
    <tr class="title"> 
        <td colspan="2" align="center"> <strong>管理员登录</strong></td>
      </tr>
      
    <tr> 
      <td height="120" colspan="2" class="tdbg">
<table width="309" border="0" cellspacing="8" cellpadding="0" align="center">
          <tr> 
            <td align="right">用户名称：</td>
            <td><input name="UserName"  type="text"  id="UserName2" size="23" maxlength="20"></td>
          </tr>
          <tr> 
            <td align="right">用户密码：</td>
            <td><input name="Password"  type="password"  size="23" maxlength="20"></td>
          </tr>
          <tr> 
            <td colspan="2"> <div align="center"> 
                <input   type="submit" name="Submit" value=" 确认 ">
                &nbsp; 
                <input name="reset" type="reset"  id="reset" value=" 清除 ">
                <br>
              </div></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
</form>
</body>
</html>
