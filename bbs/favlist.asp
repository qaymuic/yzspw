<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%

Dvbbs.LoadTemplates("usermanager")
Dvbbs.Stats=Dvbbs.MemberName&template.Strings(6)
Dvbbs.Nav()
Dvbbs.Head_var 0,0,template.Strings(0),"usermanager.asp"
If Dvbbs.userid=0 Then
	Dvbbs.AddErrCode(6)
	Dvbbs.Showerr()
End If
Dim ErrCodes,Rs,Sql,TempLateStr
Response.Write Template.Html(0)
TempLateStr=Split(template.html(17),"||")
TempLateStr(1)=Replace(TempLateStr(1),"{$fav_del}",template.pic(13))


If request("action")="delet" Then
	call delete()
Else
	Response.Write TempLateStr(0)
	Response.Write TempLateStr(1)
	call favlist()
End If
If ErrCodes<>"" Then Response.redirect "showerr.asp?ErrCodes="&ErrCodes&"&action=OtherErr"
Dvbbs.Showerr()
Dvbbs.ActiveOnline()
Dvbbs.Footer()
Sub favlist()
	Dim currentPage,page_count,totalrec,Pcount,PageListNum,i
	PageListNum=Cint(Dvbbs.Forum_Setting(11))
	currentPage=Request("page")
	If currentpage="" or not IsNumeric(currentpage) Then
		currentpage=1
	Else
		currentpage=clng(currentpage)
	End If
	set Rs=server.createobject("adodb.recordset")
	Sql="Select * From Dv_bookmark Where UserName='"&Dvbbs.membername&"' Order By id Desc"
	Dvbbs.SqlQueryNum=Dvbbs.SqlQueryNum+1
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open SQL,Conn,1,1
	If Rs.eof And Rs.bof Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(50)
		Exit Sub
	Else
		Rs.PageSize = PageListNum
		Rs.AbsolutePage=currentpage
		page_count=0
		totalrec=Rs.recordcount
		Do While Not Rs.eof And (Not page_count = Rs.PageSize)
		Response.Write "<script>dvbbs_favlist_loop('"&rs("url")&"','"&EncodeJS(rs("topic"))&"','"&rs("addtime")&"',"&rs("id")&")</script>"
		page_count = page_count + 1
		Rs.movenext
		Loop
	End If
	Rs.close:Set rs=nothing
	If totalrec mod PageListNum=0 Then
     	Pcount= totalrec \ PageListNum
  	Else
     	Pcount= totalrec \ PageListNum+1
  	End If
	If page_count=0 Then CurrentPage=0
	Response.Write ShowPage(CurrentPage,Pcount,totalrec,PageListNum)
	Response.Write TempLateStr(2)
End Sub

Sub delete()
If Dvbbs.chkpost=False Then
	Dvbbs.AddErrCode(16)
	Exit Sub
End If
If IsNumeric(request("id")) Then
	sql="delete from Dv_bookmark where username='"&Dvbbs.membername&"' and id="&cstr(request("id"))
	Dvbbs.execute sql
End If
Dvbbs.Dvbbs_Suc("<li>"+template.Strings(46))
Session("ispost")="0"
End Sub

'·ÖÒ³Êä³ö
Function ShowPage(CurrentPage,Pcount,totalrec,PageNum)
	Dim SearchStr
	SearchStr=Request("action")
	ShowPage=template.html(16)
	ShowPage=Replace(ShowPage,"{$colSpan}",3)
	ShowPage=Replace(ShowPage,"{$CurrentPage}",CurrentPage)
	ShowPage=Replace(ShowPage,"{$Pcount}",Pcount)
	ShowPage=Replace(ShowPage,"{$PageNum}",PageNum)
	ShowPage=Replace(ShowPage,"{$totalrec}",totalrec)
	ShowPage=Replace(ShowPage,"{$SearchStr}",SearchStr)
	ShowPage=Replace(ShowPage,"{$redcolor}",Dvbbs.mainsetting(1))
End Function

Function EncodeJS(str)
EncodeJS = Replace(Replace(Replace(Replace(str,"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")
End Function

%>