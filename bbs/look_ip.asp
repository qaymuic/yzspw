<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dvbbs.LoadTemplates("dispuser")
Dim ErrCodes
Dim canlookip,canlockip
Dim lockid,Rs,SQl
canlookip=False 
canlockip=False 
If (Dvbbs.master or Dvbbs.superboardmaster or Dvbbs.boardmaster) and Cint(Dvbbs.GroupSetting(30))=1 Then
	canlookip=True
Else
	canlookip=False
End If
If Dvbbs.UserGroupID>3 And  CInt(Dvbbs.GroupSetting(30))=1 Then
	canlookip=true
End If
If Dvbbs.FoundUserPer And  Cint(Dvbbs.GroupSetting(30))=1 Then
	canlookip=True
ElseIf Dvbbs.FoundUserPer And  CInt(Dvbbs.GroupSetting(30))=0 Then
	canlookip=False
End If

if (Dvbbs.master or Dvbbs.superboardmaster or Dvbbs.boardmaster) and Cint(Dvbbs.GroupSetting(31))=1 Then 
	canlockip=True 
Else
	canlockip=False 
End If
If Dvbbs.UserGroupID>3 And  Cint(Dvbbs.GroupSetting(31))=1 Then
	canlockip=True 
End If 
If Dvbbs.FoundUserPer And CInt(Dvbbs.GroupSetting(31))=1 Then
	canlockip=True 
ElseIf Dvbbs.FoundUserPer and Cint(Dvbbs.GroupSetting(31))=0 Then
	canlockip=False 
End If 
Dvbbs.stats=template.Strings(13)
Dvbbs.Nav()
Dvbbs.Head_var 0,0,Replace(template.Strings(0),"{$MemberName}",""),"dispuser.asp"
If Not Dvbbs.ChkPost() And Request("action") <> "" Then
	Response.redirect "showerr.asp?ErrCodes=<li>您不要从外部提交数据&action=OtherErr"
End If
If request("action")="setlockip" Then
	call Setlockip()
ElseIf request("action")="unlock" Then
	call unlock()
Else
	call lookip()
End If
Showerr()
Dvbbs.Showerr()
Call Dvbbs.activeonline()
Call Dvbbs.footer()

Sub lookip()
	If Not canlookip Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(7)
		Exit sub
	End If

	Dim ip,useraddress,iGetLockIP
	ip=Request("ip")
	useraddress=lookaddress(replace(ip,"'",""))
	iGetLockIP=GetLockIP(replace(ip,"'",""))
	lockid=lockid
%>
<table class=tableborder1 cellspacing="1" cellpadding="3" align="center">
<tr align=center>
<th height=25>查看 <%=IP%>的来源</th>
</tr>
<tr><td height=25 class=tablebody1><blockquote><%=useraddress%></blockquote></td></tr>
<%If canlookip Then%>
	<tr><td height=25 class=tablebody2 align=center><B>管理操作</B>：
	<%If iGetLockIP Then%>
		<a href="?action=unlock&boardid=<%=Dvbbs.BoardID%>&id=<%=lockid%>">该用户IP已被锁定，解除锁定
	<%Else%>
		<a href="?action=setlockip&ip=<%=IP%>&boardid=<%=Dvbbs.BoardID%>">限制该IP不允许访问</a>
	<%End If%>
	</td></tr>
<%End If%>
</table>
<%
End Sub 

Sub Setlockip()
	If Not canlockip then
		ErrCodes=ErrCodes+"<li>"+template.Strings(8)
		Exit sub
	End If
	If request("reaction")="yes" Then
		Dim sip
		sip=cstr(request.form("ip1"))
		If sip<>"" Then
			If Instr(sip,"*.")>0 Then
				ErrCodes=ErrCodes+"<li>前台最多只能限制四类IP，如218.1.2.*"
				Exit Sub
			End If
			If Instr(sip,"*.*.")>0 Then
				ErrCodes=ErrCodes+"<li>前台最多只能限制四类IP，如218.1.2.*"
				Exit Sub
			End If
			If Instr(sip,"*.*.*.")>0 Then
				ErrCodes=ErrCodes+"<li>前台最多只能限制四类IP，如218.1.2.*"
				Exit Sub
			End If
			If Trim(Dvbbs.CacheData(25,0))<>"" Then
				sip=Trim(Dvbbs.CacheData(25,0)) & "|" & Replace(sip,"|","")
			End If
		End If
		If sip<>"" Then
			dvbbs.execute("update dv_setup set Forum_LockIP='"&replace(sip,"'","''")&"'")
			Dvbbs.Name="setup"
			dvbbs.reloadsetup
		End If
		sql="insert into dv_log (l_touser,l_username,l_content,l_ip,l_type) values ('-','"&Dvbbs.membername&"','用户操作：限制IP"&Dvbbs.checkstr(Request.Form("ip1"))&"-"&Dvbbs.checkstr(Request.Form("ip2"))&"','"&Dvbbs.UserTrueIP&"',6)"
		dvbbs.Execute(SQL)
		Dvbbs.Dvbbs_Suc("<li>"+template.Strings(9))
	Else
		Dim userip,ips,GetIp1,useraddress,ip
		If request("ip")<>"" then
			userip=request("ip")
			ips=Split(userIP,".")
			GetIp1=ips(0)&"."&ips(1)&"."&ips(2)&".*"
		Else  
			userip=""
			GetIp1=""
			GetIp2=""
		End If
		ip=Request("ip")
		useraddress=lookaddress(replace(request("ip"),"'",""))
%>
<table class=tableborder1 cellspacing="1" cellpadding="3" align="center">
<tr align=center>
<th height=25>锁定 <%=IP%> 的来源</th>
</tr>
<tr><td height=25 class=tablebody1><blockquote><%=useraddress%></blockquote></td></tr>
<FORM METHOD=POST ACTION="?action=setlockip&boardid='+Dvbbs.BoardID+'">
<input type=hidden name="reaction" value="yes">
<tr><td height=40 class=tablebody1>
<B>说明</B>：您可以添加多个限制IP，每个IP用|号分隔，限制IP的书写方式如202.152.12.1就限制了202.152.12.1这个IP的访问，如202.152.12.*就限制了以202.152.12开头的IP访问，同理*.*.*.*则限制了所有IP的访问。在添加多个IP的时候，请注意最后一个IP的后面不要加|这个符号，<b>在前台只能做一个星号的四类IP限制</b>
</td></tr>
<tr><td height=40 class=tablebody1>
<B>限制I&nbsp;P</B>：<input type="text" name="ip1" size="30" value="<%=GetIp1%>">&nbsp;&nbsp;<input type="submit" name="Submit" value="提 交">
</td></tr>
</FORM>
</table>
<%
	End If 
End Sub 

sub unlock()
	If Not canlockip Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(8)
		Exit sub
	End If
	Dim locklist,unlockip
	locklist=Trim(Dvbbs.CacheData(25,0))
	If locklist<>"" Then
		If Trim(request("id"))="" Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(10)
			Exit sub
		End If
		locklist = "|" & locklist & "|"
		unlockip = Replace(Replace(request("id"),"|",""),"'","")
		unlockip = "|" & unlockip
		locklist = Replace(locklist,unlockip,"")
		unlockip = Split(request("id"),".")
		If Ubound(unlockip)<>3 Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(10)
			Exit sub
		End If
		locklist = Split(locklist,"|")
		Dim i,ilocklist
		For i = 1 To Ubound(locklist)-1
			If i = 1 Then
				ilocklist = locklist(i)
			Else
				ilocklist = ilocklist & "|" & locklist(i)
			End If
		Next
		dvbbs.execute("update dv_setup set Forum_LockIP='"&replace(Trim(ilocklist),"'","")&"'")
		Dvbbs.Name="setup"
		dvbbs.reloadsetup
	End If

	sql="insert into dv_log (l_touser,l_username,l_content,l_ip,l_type) values ('-','"&Dvbbs.membername&"','用户操作：解除IP限制','"&Dvbbs.UserTrueIP&"',6)"
	Dvbbs.Execute(SQL)
	Dvbbs.Dvbbs_Suc("<li>"+template.Strings(11))
End Sub

Function lookaddress(sip)
	Dim str1,str2,str3,str4
	Dim num
	Dim irs
	If isnumeric(left(sip,2)) Then
		If sip="127.0.0.1" Then sip="192.168.0.1"
		str1=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str2=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str3=left(sip,instr(sip,".")-1)
		str4=mid(sip,instr(sip,".")+1)
		If isNumeric(str1)=0 Or isNumeric(str2)=0 Or isNumeric(str3)=0 Or isNumeric(str4)=0 Then

		Else
			num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
			Dim adb,aConnStr,AConn
			adb = "data/ipaddress.mdb"
			aConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(adb)
			Set AConn = Server.CreateObject("ADODB.Connection")
			aConn.Open aConnStr
			sql="select country,city from dv_address where ip1 <="&num&" and ip2 >="&num
			Set irs=AConn.Execute(sql)
			If irs.eof And irs.bof Then 
				lookaddress=template.Strings(12)
			Else
				Do While Not irs.eof
					lookaddress=lookaddress & "<br>" &irs(0) & irs(1)
				irs.movenext
				Loop
			End If
			irs.close
			Set irs=nothing
			Set AConn=Nothing
		End If
	Else
		lookaddress=template.Strings(12)
	End If
End Function

Function getLockIP(sip)
	getLockIP=False 
	Dim locklist
	locklist=Trim(dvbbs.CacheData(25,0))
	If locklist="" Then Exit Function
	Dim i,StrUserIP,StrKillIP
	StrUserIP=sip
	locklist=Split(locklist,"|")
	If StrUserIP="" Then Exit Function
	StrUserIP=Split(StrUserIP,".")
	If Ubound(StrUserIP)<>3 Then Exit Function
	For i= 0 to UBound(locklist)
		If locklist(i)<>"" Then 
			StrKillIP = Split(locklist(i),".")
			If Ubound(StrKillIP)<>3 Then Exit For
			getLockIP = True
			If (StrUserIP(0) <> StrKillIP(0)) And Instr(StrKillIP(0),"*")=0 Then getLockIP=False
			If (StrUserIP(1) <> StrKillIP(1)) And Instr(StrKillIP(1),"*")=0 Then getLockIP=False
			If (StrUserIP(2) <> StrKillIP(2)) And Instr(StrKillIP(2),"*")=0 Then getLockIP=False
			If (StrUserIP(3) <> StrKillIP(3)) And Instr(StrKillIP(3),"*")=0 Then getLockIP=False
			If getLockIP Then
				lockid=locklist(i)
				Exit For
			End If
		End If
	Next
End Function

'显示错误信息
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