<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim username
dim password
dim repassword
dim answer
Dvbbs.LoadTemplates("login")
Dvbbs.Stats=template.Strings(21)
Dvbbs.Nav()
Dvbbs.Head_var 0,"",template.Strings(0),""
main
Dvbbs.activeonline()
Dvbbs.Footer()
Sub main()
	Dim Rs,SQL
	If request("username")="" or request("pass")="" or request("repass")="" or request("answer")="" then
		 showerr template.Strings(22)
	Else 
		username=Dvbbs.checkStr(request("username"))
		password=Dvbbs.checkStr(request("pass"))
		repassword=md5(Dvbbs.checkStr(request("repass")),16)
		answer=md5(request("answer"),16)
		sql="select userpassword,userclass,UserGroupID from [Dv_user] where username='"&username&"' and userpassword='"&password&"' and useranswer='"&answer&"'"
		set rs=server.createobject("adodb.recordset")
		If Not IsObject(Conn) Then ConnectionDatabase
		rs.open sql,conn,1,3
		If rs.eof and rs.bof Then 
			showerr  template.Strings(23) 
			Exit  Sub 
		Else 
			If Rs("usergroupid")<4 Then
				showerr template.Strings(7) 
				Exit  Sub
			Else
				Rs("userpassword")=repassword
				Rs.Update
				Response.Write template.html(11)
				Rs.Close
				Set Rs=Nothing 
			End If 
		End If
	End If
End Sub

Sub showerr(errmsg)
	template.html(9)=Replace(template.html(9),"{$Errmsg}",errmsg)
	Response.Write template.html(9)
End Sub 
%>