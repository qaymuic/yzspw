<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<%
	If Dvbbs.IsReadonly()  And Not Dvbbs.Master Then Response.redirect "showerr.asp?action=readonly&boardid="&dvbbs.boardID&"" 
	Dim rootid,isagree,Annisagree
	TotalUseTable=CheckBoardInfo
	TotalUseTable=""
	Dim getmoney,TotalUseTable
	'设置投票所用金钱
	getmoney=Cint(Dvbbs.GroupSetting(47))
	Dvbbs.stats="帖子投票"
	Dvbbs.ShowErr()
	If Dvbbs.Userid=0 Then
		Dvbbs.AddErrCode(6)
	End If
	If CInt(Dvbbs.GroupSetting(6))=0 Then
		Dvbbs.AddErrCode(63)
	End If
	If request("id")="" Then
		Dvbbs.AddErrCode(43)
	ElseIf Not Isnumeric(request("id")) Then
		Dvbbs.AddErrCode(30)
	Else
		rootid=Clng(request("id"))
	End If
	Dvbbs.ShowErr()
	if Isnumeric(session("postagree")) then
		If Clng(session("postagree"))=Clng(rootid) Then
			Dvbbs.AddErrCode(46)
		End If
	End If
	If request("isagree")="" Then
		Dvbbs.AddErrCode(42)
	ElseIf  not Isnumeric(request("isagree")) Then
		Dvbbs.AddErrCode(35)
	Else
		isagree=request("isagree")
	End If
	Dvbbs.ShowErr()
	Main
sub main()
	Dim Rs,sql
	set rs=Dvbbs.execute("select userWealth from [Dv_user] where userid="&Dvbbs.userid)
	If Rs(0)<getmoney Then
		Dvbbs.AddErrCode(47)
		Dvbbs.ShowErr()

	Else
		Dvbbs.execute("update [Dv_user] set userWealth=userWealth-"&getmoney&" where userid="&Dvbbs.userid)
		Set Rs=Dvbbs.execute("select PostTable from Dv_topic where topicid="&Clng(rootid))
		TotalUseTable=rs(0)
		rs.close
		sql="select top 1 isagree from "&TotalUseTable&" where rootid="&rootid&" order by Announceid"
		Set Rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,3
		Dvbbs.SqlQueryNum=Dvbbs.SqlQueryNum+1
		If rs.eof and rs.bof Then
			Dvbbs.AddErrCode(30)
		Else
			If Not isnull(rs(0)) and rs(0)<>"" Then
				If InStr(rs(0),"[isubb]") > 0 Then 
					If Replace(Rs(0),"[isubb]","")<>"" Then
						Annisagree=split(Replace(Rs(0),"[isubb]",""),"|")
						If Cint(isagree)=1 Then
							isagree=Annisagree(0)+1
							rs("isagree")="[isubb]"&isagree & "|" & Annisagree(1)
						Else
							isagree=Annisagree(1)+1
							rs("isagree")="[isubb]"&Annisagree(0) & "|" & isagree
						End If
					Else
						If  Cint(isagree)=1 Then
							rs("isagree")="[isubb]1|0"
						Else
							rs("isagree")="[isubb]0|1"
						End If
					End If
				Else
					Annisagree=split(rs(0),"|")
					If Cint(isagree)=1 Then
						isagree=Annisagree(0)+1
						rs("isagree")=isagree & "|" & Annisagree(1)
					Else
						isagree=Annisagree(1)+1
						rs("isagree")=Annisagree(0) & "|" & isagree
					End If
				End If
				rs.Update 
			Else
				If  Cint(isagree)=1 Then
					rs("isagree")="1|0"
				Else
					rs("isagree")="0|1"
				End If
				rs.Update
			End If
		End If
		Rs.Close
		Set Rs=Nothing 
	End If 
	If Dvbbs.ErrCodes<>"" Then
		Dvbbs.Nav()
		Dvbbs.ShowErr()
	End If
	session("postagree")=rootid
	response.redirect "dispbbs.asp?boardid="&request("boardid")&"&id="&rootid
End Sub
%>