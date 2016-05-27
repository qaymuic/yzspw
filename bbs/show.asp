<!--#include FILE="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp" -->
<%
Dvbbs.Loadtemplates("show")
Dvbbs.stats=template.Strings(0)
Dvbbs.Nav
If Dvbbs.BoardID=0 then
	Dvbbs.Head_var 2,0,"",""
Else
	Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
End If
Dvbbs.ShowErr()

Dim TopicCount
Dim Pcount,endpage,star,page_count
Dim toback,tonext,uploadpath
Dim rsearch
Dim maxpage
Dim Tab
Dim Rs,SQl,i
Dim BoardID
BoardID=Dvbbs.BoardID
tab=Request.cookies("tab")
If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
If right(Dvbbs.Forum_Setting(76),1)<>"/" Then Dvbbs.Forum_Setting(76)=Dvbbs.Forum_Setting(76)&"/"
uploadpath=Dvbbs.Forum_Setting(76)
Dim bbsurl
Dim TempStr,TopStr
bbsurl=Dvbbs.Get_ScriptNameUrl()
If Request("star")="" or Not IsNumeric(Request("star"))  then
	star=1
Else
	star=clng(Request("star"))
	toback=star-1
	tonext=star+1
End If 
If star <=1 Then star=1
 
If tab="" or not isnumeric(tab) Then tab=4

If Request("tab")="" or not isnumeric(Request("tab")) or Request("tab")="0" Then
	Tab=tab
Else
	Tab=clng(Request("tab"))
	Response.Cookies("tab").Expires=Date+365
	response.cookies("tab")=tab
End If
maxpage=clng(tab*3)	'每页显示文件的个数
If Request("username")="" or Request("filetype")="" or Request("boardid")="" then rsearch=""

If Request("boardid")<>"" and isnumeric(Request("boardid")) Then
	If clng(Request("boardid"))<>0 Then rsearch=rsearch&"and  F_BoardID="&clng(Request("boardid"))&" "
End If

If Request("filetype")<>"" And IsNumeric(Request("filetype")) Then rsearch=rsearch&"and  F_Type="&cint(Request("filetype"))&"  "

If Request("username")<>"" Then rsearch=rsearch&"and F_Username='"&Dvbbs.checkStr(Request("username"))&"'"
If Cint(Dvbbs.GroupSetting(49))=0 then Dvbbs.AddErrCode(54)
Dvbbs.ShowErr()
TempStr = template.html(0)
TempStr = Replace(TempStr,"{$stats}",Dvbbs.Stats)
Main()
Dvbbs.ActiveOnline()
Dvbbs.NewPassword()
Dvbbs.Footer()

Sub Main()
	Dim filetype,username
	Dim TempStr1,TempStr2,TempStr3,TempArray,TempArray1,FileArray
	TempArray = Split(template.html(4),"||")
	TempArray1 = Split(template.html(6),"||")
	FileArray = Split(Dvbbs.lanstr(5),"||")
	TempStr1 = template.html(1)
	filetype=Request("filetype")
	username=Request("username")

	'取出总数
	Dim frs
	set frs=Dvbbs.Execute("select count(F_id) from DV_Upfile where F_Flag<>4 "&rsearch&" ")
	TopicCount=frs(0)
	frs.close
	set frs=Nothing
	'定义变量
	Dim t
	Dim F_ID,F_AnnounceID,F_BoardID,F_Filename,F_FileType,F_Type,F_Readme,F_DownNum,F_ViewNum,F_Flag
	Dim F_typename,showfile,golist
	Dim Edit,Rs,sql
	edit=False 
	t=0
	set rs=Dvbbs.Execute("select * from [DV_Upfile] where F_BoardID<>0 and F_BoardID<>444 and F_BoardID<>777 and F_Flag<>4 "&rsearch&" order by F_ID desc  ")
	If Not (rs.eof and rs.bof) Then
		If TopicCount Mod Cint(maxpage)=0 Then
			Pcount= TopicCount \ Cint(maxpage)
  		Else
			Pcount= TopicCount \ Cint(maxpage)+1
	  	End If
		RS.MoveFirst
		If star > Pcount Then star = Pcount

		RS.Move (star-1) * maxpage
		page_count=0
		Do While Not rs.eof and page_count < Cint(maxpage)
			TempStr2 = TempStr1
			F_AnnounceID=rs("F_AnnounceID")
			F_Type=rs("F_Type")
			F_Readme=rs("F_Readme")
			F_Flag=Rs("F_Flag")
			F_Filename=Rs("F_Viewname")
			If F_Filename="" or Isnull(F_Filename) Then
				F_Filename=Rs("F_Filename")
				If F_Flag<>1 Then 
					If isshow(Rs("F_BoardID"))=1 Then
						If Dvbbs.Forum_Setting(75)="0" Then
							F_Filename=bbsurl&"UploadFile/"&F_Filename
						Else
							F_Filename="showimg.asp?Boardid="&Rs("F_BoardID")&"&filename="&F_Filename
						End If
					Else
						F_Filename=Dvbbs.Forum_Info(6)
					End If
					
				End If
			End If			

			If Not IsNull(F_AnnounceID) And F_AnnounceID<>"" Then
				F_AnnounceID=split(F_AnnounceID,"|")
				TempStr2 = Replace(TempStr2,"{$golist}",TempArray(2))
				TempStr2 = Replace(TempStr2,"{$announceid}",F_AnnounceID(0))
				TempStr2 = Replace(TempStr2,"{$replyid}",F_AnnounceID(1))
			Else
				TempStr2 = Replace(TempStr2,"{$golist}",template.Strings(26))
			End If
			Select case F_Type
			Case 1
					F_typename=FileArray(1)
					showfile=TempArray(3)
					showfile=Replace(showfile,"{$imgurl}",Dvbbs.HTMLEncode(Dv_FilterJS(F_Filename)))
					
			Case 2
				F_typename=FileArray(2)
				showfile=TempArray(4)
				showfile=Replace(showfile,"{$fileurl}",Dvbbs.HTMLEncode(F_Filename))
			Case 3
				F_typename=FileArray(3)
				showfile=TempArray(4)
				showfile=Replace(showfile,"{$fileurl}",Dvbbs.HTMLEncode(F_Filename))
			Case 4
				F_typename=FileArray(4)
				showfile=TempArray(4)
				showfile=Replace(showfile,"{$fileurl}",Dvbbs.HTMLEncode(F_Filename))
			Case Else
				F_typename=FileArray(0)
				showfile=TempArray(5)
			End Select

			If Dvbbs.GroupSetting(48)=1 Then
				If Dvbbs.master or Dvbbs.superboardmaster or Dvbbs.boardmaster Then
					edit=True
				ElseIf rs("F_Username")=Dvbbs.membername Then
					edit=True
				Else
					edit=False
				End If
			End If

			If Dvbbs.UserID > 0 Then
				TempStr2 = Replace(TempStr2,"{$useraction}",TempArray1(0))
				If edit Then TempStr2 = Replace(TempStr2,"{$usereditinfo}",TempArray1(1))
				TempStr2 = Replace(TempStr2,"{$usereditinfo}","")
			End If
			TempStr2 = Replace(TempStr2,"{$useraction}","")
			If F_Type = 0 Then
				TempStr2 = Replace(TempStr2,"{$typenum}",rs("F_DownNum"))
			Else
				TempStr2 = Replace(TempStr2,"{$typenum}",rs("F_ViewNum"))
			End If
			If Trim(F_Readme) <> "" Then
				If Len(F_Readme)>26 Then
					F_Readme = Dvbbs.HtmlEncode(Left(F_Readme,26)) & "..."
				Else
					F_Readme = Dvbbs.HtmlEncode(F_Readme)
				End If
			Else
				F_Readme = template.Strings(1)
			End If
			TempStr2 = Replace(TempStr2,"{$filereadme}",F_Readme)
			TempStr2 = Replace(TempStr2,"{$showfile}",showfile)

			TempStr2 = Replace(TempStr2,"{$filetype}",F_Type)
			TempStr2 = Replace(TempStr2,"{$refiletype}",rs("F_FileType"))
			TempStr2 = Replace(TempStr2,"{$typename}",F_typename)
			TempStr2 = Replace(TempStr2,"{$username}",Dvbbs.HtmlEncode(rs("F_Username")))
			TempStr2 = Replace(TempStr2,"{$fid}",rs("F_ID"))
			TempStr2 = Replace(TempStr2,"{$boardid}",rs("F_BoardID"))
			If t = (tab-1) Then
				t = 0
				TempStr2 = Replace(TempStr2,"{$tableinfo}",TempArray(1))
			Else
				TempStr2 = Replace(TempStr2,"{$tableinfo}","")
				t = t + 1
			End If
			TempStr3 = TempStr3 & TempStr2
			page_count = page_count + 1
		Rs.MoveNext
		Loop
		TempStr2 = template.html(5)
		TempStr2 = Replace(TempStr2,"{$Pcount}",Pcount)
		TempStr2 = Replace(TempStr2,"{$boardid}",Dvbbs.boardid)
		TempStr2 = Replace(TempStr2,"{$star}",star)
		TempStr2 = Replace(TempStr2,"{$pagelistnum}",TopicCount)
		TempStr2 = Replace(TempStr2,"{$limitednum}",maxpage)
	Else
		TempStr = Replace(TempStr,"{$showloop}",template.html(3))
	End If
	Rs.Close
	Set Rs=Nothing
	TempStr = Replace(TempStr,"{$pagelist}",TempStr2)

	TempStr = Replace(TempStr,"{$showloop}",TempStr3)
	top 10,"F_ViewNum"
	TempStr = Replace(TempStr,"{$toplist1}",TopStr)
	top 10,"F_ID"
	TempStr = Replace(TempStr,"{$toplist2}",TopStr)
	If UserName<>"" Then TempStr = Replace(TempStr,"{$userinfo}",TempArray(0))
	TempStr = Replace(TempStr,"{$filetype}",filetype)
	TempStr = Replace(TempStr,"{$tab}",tab)
	TempStr = Replace(TempStr,"{$username}",UserName)
	TempStr = Replace(TempStr,"{$userinfo}","")

	Dim FileJump
	For t = 0 To Ubound(FileArray)
		FileJump = FileJump & "<option value=" & t
		If filetype = t Then FileJump = FileJump & "selected"
		FileJump = FileJump & "> " & FileArray(t) & "</option>"
	Next
	TempStr = Replace(TempStr,"{$filetypejump}",FileJump)

	Dim Forum_Boards,Board_Data,BoardJumpList,i,ii,Depth
	Forum_Boards=Split(Dvbbs.cachedata(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.Name="BoardInfo_" & Forum_Boards(i)
  		If Dvbbs.ObjIsEmpty() Then Dvbbs.ReloadBoardInfo(Forum_Boards(i))
		Board_Data=Dvbbs.Value
		BoardJumpList = BoardJumpList & "<option value="""&Forum_Boards(i)&""" "
		If Clng(Forum_Boards(i))=Dvbbs.BoardID Then BoardJumpList = BoardJumpList & "selected"
		BoardJumpList = BoardJumpList & ">"
		Depth=Board_Data(4,0)
		Select Case Depth
		Case 0
			BoardJumpList = BoardJumpList & "╋"
		Case 1
			BoardJumpList = BoardJumpList & "&nbsp;&nbsp;├"
		End Select
		If Depth>1 Then
			For ii=2 To Depth
				BoardJumpList = BoardJumpList & "&nbsp;&nbsp;│"
			Next
			BoardJumpList = BoardJumpList & "&nbsp;&nbsp;├"
		End If
		BoardJumpList = BoardJumpList & Board_Data(1,0)&"</option>"
	Next
	Board_Data=Null
	Forum_Boards=Null
	TempStr = Replace(TempStr,"{$boardjump}",BoardJumpList)
	TempStr = Replace(TempStr,"{$alertcolor}",Dvbbs.mainsetting(1))
	Response.Write TempStr
End sub

'top 排行调用
Sub top(num,Str)
	TopStr = ""
	Dim TempStr1,TempStr2,TempStr3,tmptitle
	Dim k
	k=0
	TempStr1 = template.html(2)
	Set Rs=Dvbbs.Execute("select top "&num&" F_ID,F_BoardID,F_Filename,F_Readme from Dv_Upfile Where F_BoardID<>0 and F_BoardID<>444 and F_BoardID<>777 order by "&str&" desc,F_addTime desc")
	Do While Not Rs.Eof
		TempStr2 = TempStr1
		TempStr2 = Replace(TempStr2,"{$fid}",Rs(0))
		TempStr2 = Replace(TempStr2,"{$boardid}",Rs(1))
		If isshow(Rs("F_BoardID"))=1 Then
			If Rs(3)<>""  Then
				tmptitle=Rs(3) 
			Else
				tmptitle=Rs(2) 
			End If
			tmptitle=cutStr(tmptitle,16)
		Else
			tmptitle="==隐藏或认证版内容=="
		End If
		TempStr2 = Replace(TempStr2,"{$title}",tmptitle)
		TempStr3 = TempStr3 & TempStr2
		K=K+1
		
	Rs.MoveNext
	Loop
	Set Rs=Nothing
	TopStr = TempStr3
	TempStr3 = ""
	TempStr2 = ""
	TempStr1 = ""
End Sub
Function reUBBCode(strContent)
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	strContent=replace(strContent,"&nbsp;"," ")
	re.Pattern="(\[QUOTE\])(.*)(\[\/QUOTE\])"
	strContent=re.Replace(strContent,"$2")
	re.Pattern="(\[point=*([0-9]*)\])(.*)(\[\/point\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[post=*([0-9]*)\])(.*)(\[\/post\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[power=*([0-9]*)\])(.*)(\[\/power\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[usercp=*([0-9]*)\])(.*)(\[\/usercp\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[money=*([0-9]*)\])(.*)(\[\/money\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[replyview\])(.*)(\[\/replyview\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[usemoney=*([0-9]*)\])(.*)(\[\/usemoney\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="\[username=(.[^\[]*)\](.[^\[]*)\[\/username\]"
	strContent=re.Replace(strContent,"&nbsp;")
	strContent=replace(strContent,"<I></I>","")
	set re=Nothing
	reUBBCode=Dvbbs.ChkBadWords(strContent)
End Function
'截取指定字符
Function cutStr(str,strlen)
	Str=reUBBCode(Str)
	'去掉所有HTML标记
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="<(.[^>]*)>"
	str=re.Replace(str,"")	
	set re=Nothing
	Dim l,t,c,i
	l=Len(str)
	t=0
	For i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		If c>255 Then
			t=t+2
		Else
			t=t+1
		End If
		If t>=strlen Then
			cutStr=left(str,i)&"..."
			Exit For
		Else
			cutStr=str
		End If
	Next
End Function
Function Dv_FilterJS(v)
	If Not Isnull(V) Then
		Dim t
		Dim re
		Dim reContent
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(&#)"
		t=re.Replace(v,"<I>&#</I>")
		re.Pattern="(script)"
		t=re.Replace(t,"<I>script</I>")
		re.Pattern="(js:)"
		t=re.Replace(t,"<I>js:</I>")
		re.Pattern="(value)"
		t=re.Replace(t,"<I>value</I>")
		re.Pattern="(about:)"
		t=re.Replace(t,"<I>about:</I>")
		re.Pattern="(file:)"
		t=re.Replace(t,"<I>file:</I>")
		re.Pattern="(Document.cookie)"
		t=re.Replace(t,"<I>Documents.cookie</I>")
		re.Pattern="(vbs:)"
		t=re.Replace(t,"<I>vbs:</I>")
		re.Pattern="(on(mouse|Exit|error|click|key))"
		t=re.Replace(t,"<I>on$2</I>")
		Dv_FilterJS=t
		Set Re=Nothing
	End If 
End Function
Function isshow(BID)
	If BID=0 Or IsNull(bid) Then
		isshow=0
		Exit Function
	End If
	If Not InStr((","&Dvbbs.cachedata(27,0)&","),(","&bid&","))>0 Then
		isshow=0
		Exit Function
	End If
	Dim Board_Data
	Dvbbs.Name="BoardInfo_" & BID
  	If Dvbbs.ObjIsEmpty() Then Dvbbs.ReloadBoardInfo(BID)
	Board_Data=Dvbbs.Value
	Board_Data(16,0)=Split(Board_Data(16,0),",")
	isshow=1
	If Board_Data(16,0)(1)="1" Then
		isshow=0
	ElseIf Board_Data(16,0)(2)="1" Then
		isshow=0
		
	End If
End Function
%>