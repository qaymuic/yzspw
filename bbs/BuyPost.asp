<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<%
Dvbbs.LoadTemplates("dispbbs")
Dim rootid,PostTable
Dim AnnounceID,Rs,SQL,i
PostTable=request("PostTable")
PostTable=checktable(PostTable)
If request("action")="view" Then
	Dvbbs.stats="查看购买贴子的用户"
Else
	Dvbbs.stats="购买帖子"
End If
Dvbbs.Nav
Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
rootid=Request("ID")
If rootid="" Or Not IsNumeric(rootid) Then Dvbbs.AddErrCode(30)
AnnounceID=request("replyID")
If AnnounceID="" Then
	Dvbbs.AddErrCode(30)
ElseIf Not IsNumeric(AnnounceID) Then
	Dvbbs.AddErrCode(30)
End If
If  Dvbbs.UserID=0 Then
	Dvbbs.AddErrCode(30)
End If
Dvbbs.ShowErr()
If request("action")="view" Then
	Call view()
Else
	Call main()
End If
Dvbbs.ShowErr()
Dvbbs.activeonline()
Dvbbs.footer
Sub main()
	dim re
	dim po,ii
	dim reContent
	dim strContent
	dim PostBuyUser
	po=0
	ii=0
	dim usermoney
	set rs=Dvbbs.Execute("select userWealth from [Dv_user] where userid="&Dvbbs.Userid)
	usermoney=rs(0)
	set rs=server.createobject("adodb.recordset")
	sql="select body,PostBuyUser,username,PostUserID from "&PostTable&" where Announceid="&Announceid
	rs.open sql,conn,1,3
	If rs.eof and rs.bof Then
		Dvbbs.AddErrCode(30)
		Dvbbs.ShowErr()
	Else 	
				
		strContent=Dvbbs.HTMLEncode(rs(0))
		PostBuyUser=Trim(rs(1))
		'Response.Write PostBuyUser
		'Response.End
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=True
		re.Pattern="(^.*)(\[UseMoney=*([0-9]*)\])(.*)(\[\/UseMoney\])(.*)"
		po=re.Replace(strContent,"$3")
		If IsNumeric(po) Then 
			ii=int(po) 
		Else
			ii=0
		End If
		Set re=Nothing
				
		If Dvbbs.membername=rs(2) Then
			response.write "<script>alert('呵呵，您要花钱购买自己发布的帖子嘛？');</script>"
		ElseIf  usermoney >ii then
			If (not isnull(PostBuyUser)) Or  PostBuyUser<>"" Then
				If InStr("|"&PostBuyUser&"|","|"&Dvbbs.membername&"|")>0 Then
					response.write "<script>alert('呵呵，您已经购买过了呀？');</script>"
				Else
					Dvbbs.Execute("update [Dv_user] set userWealth=userWealth-"&ii&" where userid="&Dvbbs.userid)
					Dvbbs.Execute("update [Dv_user] set userWealth=userWealth+"&ii&" where userid="&rs(3))
					If IsNull(Rs(1)) or  Rs(1)="" Then 
						rs(1)=Dvbbs.membername
					Else
						rs(1)=rs(1) & "|" & Dvbbs.membername
					End If
					Rs.Update 
					response.write "<script>alert('购买成功！');</script>"
				End If
			Else 
				Dvbbs.Execute("update [Dv_user] set userWealth=userWealth-"&ii&" where userid="&Dvbbs.userid)
				Dvbbs.Execute("update [Dv_user] set userWealth=userWealth+"&ii&" where userid="&rs(3))
				rs(1)=Dvbbs.membername
				Rs.Update
				response.write "<script>alert('购买成功！');</script>"
			End If
		Else
			response.write "<script>alert('您都没有钱呀？');</script>"
		End If
		
	End If
	Rs.Close 
	Set  Rs=Nothing
	Response.Write "<script language=""javascript"">"
	Response.Write "parent.location.href='"
	Response.Write "dispbbs.asp?boardid="&request("boardid")&"&id="&rootid&"&replyID="&AnnounceID&"&skin=1"
	Response.Write "';"
	Response.Write "</script>"
End Sub
Sub view()
	Dim PostBuyUser
	sql="select PostBuyUser from "&PostTable&" where Announceid="&Announceid
	Set rs=Dvbbs.Execute(sql)
	PostBuyUser=Trim(rs(0))
	Response.Write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1>"
	Response.Write "<TBODY><TR>"
	Response.Write "<Th height=24 colspan=1>查看购买贴子的用户</Th>"
	Response.Write "</TR>"
	Response.Write "<tr><TD class=tablebody2>"
	If (not isnull(PostBuyUser)) Or  PostBuyUser<>"" Then
		PostBuyUser=Replace(PostBuyUser,"|","<li>")
		Response.Write "<li>"&PostBuyUser		
	Else
		Response.Write "<br><li>还未有人购买！"
	End If
	Response.Write "</td></tr>"
	Response.Write "</table>"
	Set rs=Nothing
End Sub
Function checktable(Table)
	Table=Right(Trim(Table),2)
	If Not IsNumeric(table) Then Table=Right(Trim(Table),1)
	If Not IsNumeric(table) Then Dvbbs.AddErrCode(30)
	checktable="Dv_bbs"&table
End Function 
%>