<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<%
	Dvbbs.LoadTemplates("login")
	Dim Stats,ErrCodes,TempLateStr
	Dim username,i,sql,Rs,useremail
	Stats=split(template.Strings(25),"||")
	Dvbbs.Stats=Stats(0)
	dvbbs.head()
	ErrCodes=""
	If Request.form("username")="" Then ErrCodes=ErrCodes+"<li>"+template.Strings(6)
	If strLength(Request.form("username"))>Cint(Dvbbs.Forum_Setting(41)) or strLength(Request.form("username"))<Cint(Dvbbs.Forum_Setting(40)) Then
		TempLateStr=template.Strings(28)
		TempLateStr=Replace(TempLateStr,"{$RegMaxLength}",Dvbbs.Forum_Setting(41))
		TempLateStr=Replace(TempLateStr,"{$RegLimLength}",Dvbbs.Forum_Setting(40))
		ErrCodes=ErrCodes+"<li>"+TempLateStr
		TempLateStr=""
	Else
		username=Dvbbs.CheckStr(Trim(Request.form("username")))
		If Instr(username,"=")>0 or Instr(username,"%")>0 or Instr(username,chr(32))>0 or Instr(username,"?")>0 or Instr(username,"&")>0 or Instr(username,";")>0 or Instr(username,",")>0 or Instr(username,"'")>0 or Instr(username,",")>0 or Instr(username,chr(34))>0 or Instr(username,chr(9))>0 or Instr(username,"£†")>0 or Instr(username,"$")>0 Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(46)
		End If
		Dim RegSplitWords
		RegSplitWords=split(Dvbbs.forum_setting(4),",")
		for i = 0 to ubound(RegSplitWords)
			If instr(username,RegSplitWords(i))>0 Then
				ErrCodes=ErrCodes+"<li>"+template.Strings(46)
			End If
		next
	End If
	If Request("action")="" Then
	If IsValidEmail(trim(Request.form("email")))=false then
   		ErrCodes=ErrCodes+"<li>"+template.Strings(30)
	Else 
		useremail=Dvbbs.checkStr(Request.form("email"))
	End If
	End If
	If ErrCodes<>"" Then Showerr()
	if ErrCodes="" then
		If cint(Dvbbs.Forum_Setting(24))=1 Then
		If Request("action")="" Then
			sql="select username,useremail from [Dv_user] where username='"&username&"' or useremail='"&useremail&"'"
		Else
			sql="select username,useremail from [Dv_user] where username='"&username&"'"
		End If
		Else 
		sql="select username,useremail from [Dv_user] where username='"&username&"'"
		End If
		Set Rs=Dvbbs.execute(sql)
		If Not rs.eof and not rs.bof then
			If cint(Dvbbs.Forum_Setting(24))=1 And Rs("useremail")=useremail Then
				If Request("action")="" Then
				ErrCodes=ErrCodes+"<li>"+template.Strings(44)
				Else
				ErrCodes=ErrCodes+"<li>"+template.Strings(43)
				End If
			Else 
				ErrCodes=ErrCodes+"<li>"+template.Strings(44)
			End If
		End If 
		Rs.close:Set Rs=Nothing
	
		If ErrCodes="" Then 
			ErrCodes=template.Strings(45)
		End If
		Response.Write Replace(template.html(16),"{$Reportmsg}",ErrCodes)
	End If
Call Dvbbs.footer()


'œ‘ æ¥ÌŒÛ–≈œ¢
Sub Showerr()
Dim Show_Errmsg
	If ErrCodes<>"" Then 
		Show_Errmsg=Dvbbs.mainhtml(14)
		ErrCodes=Replace(ErrCodes,"{$color}",Dvbbs.mainSetting(1))
		Show_Errmsg=Replace(Show_Errmsg,"{$color}",Dvbbs.mainSetting(1))
		Show_Errmsg=Replace(Show_Errmsg,"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		Show_Errmsg=Replace(Show_Errmsg,"{$action}",Dvbbs.Stats)
		Show_Errmsg=Replace(Show_Errmsg,"{$ErrString}",ErrCodes)
	End If
	Response.write Show_Errmsg
End Sub
%>