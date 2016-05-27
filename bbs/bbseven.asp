<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dvbbs.LoadTemplates("paper_even_toplist")
Dim Rs,sql,i,isshow
isshow=False
If DVbbs.BoardID=0 then
	Dvbbs.stats=template.Strings(4)
	Dvbbs.nav()
	Dvbbs.Head_var 2,0,"",""
Else
	Dvbbs.stats=template.Strings(5)
	Dvbbs.nav()
	Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
End If

If Cint(Dvbbs.GroupSetting(39))=0  And Not Dvbbs.master Then Dvbbs.AddErrCode(55)
Dvbbs.ShowErr

boardeven()

Dvbbs.activeonline()
Dvbbs.footer()

Sub boardeven()
	Dim currentpage,page_count,Pcount
	Dim endpage
	Dim totalrec
	totalrec=0
	currentPage=request("page")
	If currentpage="" Or Not IsNumeric(currentpage) Then
		currentpage=1
	Else
		currentpage=clng(currentpage)
	End If
	Dim TempStr,TempStr1,TempStr2,TempStr3
	Dim TempArray
	TempStr = template.html(5)
	TempArray = Split(template.html(6),"||")
	TempStr2 = TempArray(1)

	Set Rs=Server.CreateObject("ADODB.RecordSet")
	If Dvbbs.BoardID>0 Then
		sql="select * from dv_log where l_boardid="&DVbbs.BoardID&" and l_type >2 order by l_addtime desc"
	Else
		sql="select * from dv_log where l_type > 2 order by l_addtime desc"
	End If
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open sql,conn,1,1
	If rs.bof And rs.eof Then
		TempStr1 = TempArray(0)
	Else
		chkshow()
		rs.PageSize = Dvbbs.Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		While (Not rs.eof) And (Not page_count = rs.PageSize)
			TempArray = rs("l_touser") & "||" & rs("l_content") & "||" & rs("l_username")
			TempArray = Dvbbs.HtmlEncode(TempArray)
			TempArray = Split(TempArray,"||")
			TempStr3 = TempStr2
			TempStr3 = Replace(TempStr3,"{$username}",TempArray(0))
			TempStr3 = Replace(TempStr3,"{$content}",TempArray(1))
			TempStr3 = Replace(TempStr3,"{$addtime}",rs("l_addtime"))
			If isshow or Dvbbs.MemberName=rs("l_username") Then
				TempStr3 = Replace(TempStr3,"{$postuser}","<a href=dispuser.asp?name="&TempArray(2)&" target=_blank>"&TempArray(2)&"</a>")
			Else
				TempStr3 = Replace(TempStr3,"{$postuser}","±£√‹")
			End If
			TempStr1 = TempStr1 & TempStr3
			page_count = page_count + 1
		Rs.Movenext
		Wend
	End If

  	If totalrec Mod Dvbbs.Forum_Setting(11)=0 Then
     		Pcount= totalrec \ Dvbbs.Forum_Setting(11)
  	Else
     		Pcount= totalrec \ Dvbbs.Forum_Setting(11)+1
  	End If
	TempStr = Replace(TempStr,"{$evenloop}",TempStr1)
	TempStr = Replace(TempStr,"{$pagelist}",template.html(3))
	TempStr = Replace(TempStr,"{$page}",currentpage)
	TempStr = Replace(TempStr,"{$Pcount}",Pcount)
	TempStr = Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	TempStr = Replace(TempStr,"{$alertcolor}",Dvbbs.mainsetting(1))
	TempStr = Replace(TempStr,"{$pagelimited}",Dvbbs.Forum_Setting(11))
	TempStr = Replace(TempStr,"{$listnum}",totalrec)
	TempStr = Replace(TempStr,"{$boardid}",Dvbbs.BoardID)
	Response.Write TempStr

	Rs.Close
	Set Rs=Nothing
	
End Sub
Sub chkshow()
	If Dvbbs.master or Dvbbs.superboardmaster  Then
		isshow=True
	ElseIf Dvbbs.BoardID<>0 Then 
		If Dvbbs.Board_Setting(36)<>"" and IsNumeric(Dvbbs.Board_Setting(36)) Then
			If Cint(Dvbbs.Board_Setting(36))=1  Then
				isshow=True
			Else
				isshow=False 
			End If
		End If
	Else
		isshow=False 
	End If
End Sub
%>
