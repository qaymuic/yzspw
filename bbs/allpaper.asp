<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dvbbs.LoadTemplates("paper_even_toplist")
Dim rs,sql,i
If dvbbs.boardid=0 Then
	dvbbs.stats=template.Strings(0)
	Dvbbs.Nav()
	Dvbbs.Head_var 2,0,"",""
Else
	dvbbs.stats=template.Strings(1)
	Dvbbs.Nav()
	Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
End If

Dim TempStr
If Not(Dvbbs.boardmaster or Dvbbs.master or Dvbbs.superboardmaster) Then Response.redirect "showerr.asp?ErrCodes=<li>只有管理员才能登录。&action=OtherErr"
If Dvbbs.Forum_Setting(56)=0 Then Dvbbs.AddErrCode(52)
Dvbbs.ShowErr()
If request("action")="delpaper" Then
	call batch()
Else
	call boardeven()
End If
Dvbbs.ShowErr()
Dvbbs.ActiveOnline
Dvbbs.Footer()

Sub boardeven()
	Dim totalrec
	Dim n
	Dim currentpage,page_count,Pcount
	Pcount=0
	totalrec=0
	currentPage=request("page")
	If currentpage="" Or not IsNumeric(currentpage) Then
		currentpage=1
	Else
		currentpage=clng(currentpage)
	End If
	Dim TempArray,TempStr1,TempStr2,TempStr3
	TempStr = template.html(0)
	TempArray = Split(template.html(1),"||")
	TempStr2 = template.html(2)
	If Dvbbs.GroupSetting(27)="1" Then TempStr = Replace(TempStr,"{$manageinfo}",TempArray(2))
	TempStr = Replace(TempStr,"{$manageinfo}","")

	set rs=server.createobject("adodb.recordset")
	If dvbbs.boardid=0 Then
	sql="select * from dv_smallpaper order by s_addtime desc"
	Else
	sql="select * from dv_smallpaper where s_boardid="&dvbbs.boardid&" order by s_addtime desc"
	End If
	If Not IsObject(Conn) Then ConnectionDatabase
	rs.open sql,conn,1,1
	If rs.bof And rs.eof Then
		TempStr1 = TempArray(0)
		TempStr = Replace(TempStr,"{$pagelist}","")
	Else
		rs.PageSize = Dvbbs.Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) And (not page_count = rs.PageSize)
			TempStr3 = TempStr2
			TempStr3 = Replace(TempStr3,"{$username}",Dvbbs.HtmlEncode(rs("s_username")))
			TempStr3 = Replace(TempStr3,"{$addtime}",rs("s_addtime"))
			TempStr3 = Replace(TempStr3,"{$title}",Dvbbs.HtmlEncode(rs("s_title")))
			If Dvbbs.GroupSetting(27)="1" Then
				TempStr3 = Replace(TempStr3,"{$manageinfo1}",TempArray(1) & rs("s_hits"))
			Else
				TempStr3 = Replace(TempStr3,"{$manageinfo1}",rs("s_hits"))
			End If
			TempStr3 = Replace(TempStr3,"{$sid}",rs("s_id"))
			TempStr1 = TempStr1 & TempStr3
			page_count = page_count + 1
		rs.movenext
		wend
		Pcount=rs.PageCount
	rs.close
	set rs=nothing	
	End If
	TempStr = Replace(TempStr,"{$paperloop}",TempStr1)
	TempStr = Replace(TempStr,"{$pagelist}",template.html(3))
	TempStr = Replace(TempStr,"{$page}",currentpage)
	TempStr = Replace(TempStr,"{$Pcount}",Pcount)
	TempStr = Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	TempStr = Replace(TempStr,"{$alertcolor}",Dvbbs.mainsetting(1))
	TempStr = Replace(TempStr,"{$pagelimited}",Dvbbs.Forum_Setting(11))
	TempStr = Replace(TempStr,"{$listnum}",totalrec)
	TempStr = Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
	Response.Write TempStr
End Sub

Sub batch()
	Dim sid,fixid
	Dim adminpaper
	adminpaper=False
	If dvbbs.userid=0 Then
		Dvbbs.AddErrCode(34)
	End If
	If (dvbbs.master Or dvbbs.superboardmaster Or dvbbs.boardmaster) And Cint(dvbbs.GroupSetting(27))=1 Then
		adminpaper=True
	Else
		adminpaper=False
	End If
	If dvbbs.UserGroupID>3 And Cint(dvbbs.GroupSetting(27))=1 Then
		adminpaper=True
	End If
	If dvbbs.FoundUserPer And Cint(dvbbs.GroupSetting(27))=1 Then
		adminpaper=True
	ElseIf dvbbs.FoundUserPer And Cint(dvbbs.GroupSetting(27))=0 Then
		adminpaper=False
	End If
	If not adminpaper Then
		Dvbbs.AddErrCode(28)
	End If
	If request.form("sid")="" Then
		Dvbbs.AddErrCode(35)
	Else
		sid=replace(request.Form("sid"),"'","")
		sid=replace(sid,";","")
		sid=replace(sid,"--","")
		sid=replace(sid,")","")
		fixid=replace(sid," ","")
		fixid=replace(fixid,",","")
		If Not IsNumeric(fixid) Then
			Dvbbs.AddErrCode(35)
			Exit Sub
		End If
	End If 	
	If dvbbs.ErrCodes<>"" Then exit Sub
	Dvbbs.Execute("delete from dv_smallpaper where s_boardid="&dvbbs.boardid&" And s_id in ("&sid&")")

	Dvbbs.Name = "BoardInfo_" & Dvbbs.BoardID
	Dvbbs.LoadBoardNews_Paper(Dvbbs.BoardID)
	Dvbbs.Dvbbs_Suc(template.Strings(2))
	
End Sub
%>