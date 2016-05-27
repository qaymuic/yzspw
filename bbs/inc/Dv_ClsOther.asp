<%
Rem 除首页外通用函数
'Dvbbs.Board_Setting(40)是否继承上级版主，顺带取出上级论坛版面信息
'最多只取向上的10级版面信息
'输出导航菜单字串
Function CheckBoardInfo()
	Dim i
	Dvbbs.Boardmaster =False
	If Dvbbs.BoardID>0 and Dvbbs.BoardParentID>0 Then	
		Dim TempData,NavStr
		If Not IsArray(Dvbbs.Board_Data(22,0)) Then
			If Clng(Dvbbs.Board_Data(22,0))=Dvbbs.BoardID And Dvbbs.Board_Data(2,0)>0 Then
				Dvbbs.Name = "BoardInfo_" & Dvbbs.BoardID
				Call Dvbbs.LoadBoardParentStr (Dvbbs.Board_Data(3,0))
				Dvbbs.Board_Data = Dvbbs.Value
			End If
		End If		
		TempData=Dvbbs.Board_Data(22,0)
		If Cstr(Dvbbs.Board_Data(21,0))=Cstr(Dvbbs.BoardID) Then
			Dvbbs.Name = "BoardInfo_" & Dvbbs.BoardID
			Call Dvbbs.LoadBoardList(Dvbbs.BoardID,1)
			Call Dvbbs.LoadBoardList(Dvbbs.BoardID,0)
			Dvbbs.Board_Data = Dvbbs.Value
		End If
		If Dvbbs.Master Then
			Dvbbs.Boardmaster=True
		ElseIf Dvbbs.Superboardmaster Then
			Dvbbs.Boardmaster=True
		ElseIf Dvbbs.UserGroupID =3 And Not Trim(Dvbbs.BoardMasterList) = "" Then
			If Instr("|"&Dvbbs.BoardMasterList&"|","|"&Dvbbs.Membername&"|")>0 Then
				Dvbbs.Boardmaster=True
			End If
		End If
	ElseIf Dvbbs.BoardID>0 and Dvbbs.UserID>0 Then
		If Dvbbs.Master Then
			Dvbbs.Boardmaster=True
		ElseIf Dvbbs.Superboardmaster Then
			Dvbbs.Boardmaster=True
		ElseIf Dvbbs.UserGroupID =3 And Not Trim(Dvbbs.BoardMasterList) = "" Then
			If Instr("|"&lcase(Dvbbs.BoardMasterList)&"|","|"&lcase(Dvbbs.Membername)&"|")>0 Then
				Dvbbs.Boardmaster=True
			End If
		End If
	End If
	If Dvbbs.BoardID>0 and Dvbbs.BoardParentID>0 Then
	For i=0 To Ubound(TempData,2)
			If i=0 Then
				If Dvbbs.GroupSetting(37)="1" Then
					NavStr=" <a href=""list.asp?boardid="&TempData(0,i)&""" onMouseOver=""showmenu(event,'"&Dvbbs.Board_Data(21,0)&"')"">"& TempData(1,i) &"</a> "
				Else
					NavStr=" <a href=""list.asp?boardid="&TempData(0,i)&""" onMouseOver=""showmenu(event,'"&Dvbbs.Board_Data(26,0)&"')"">"& TempData(1,i) &"</a> "
				End If
			Else
				NavStr=NavStr& "→ <a href=""list.asp?boardid="&TempData(0,i)&""">"& TempData(1,i) &"</a> "
			End If
			If Cint(Dvbbs.Board_Setting(40))=1 And Not Dvbbs.Boardmaster Then
				If Dvbbs.UserGroupID =3  And Trim(TempData(2,i))<>"" Then
					If instr("|"&lcase(TempData(2,i))&"|","|"&lcase(Dvbbs.membername)&"|")>0 Then
						Dvbbs.Boardmaster=True
					Else
						Dvbbs.Boardmaster=False 
					End If
				End If
			End If
			If i>9 Then Exit For
	Next
	CheckBoardInfo=NavStr
	End If
	Call GetBoardPermission()
	'Response.Write Dvbbs.Boardmaster
End Function
Rem 获得版面用户组权限配置
Public Sub GetBoardPermission()
	Dim Rs,IsGroupSetting
	IsGroupSetting = Dvbbs.IsGroupSetting
	If IsGroupSetting<>"" And Not IsNull(IsGroupSetting) Then
		IsGroupSetting = "," & IsGroupSetting & ","
		
		If InStr(IsGroupSetting,"," & Dvbbs.UserGroupID & ",")>0 Then
			Set Rs=Dvbbs.Execute("Select PSetting From Dv_BoardPermission Where Boardid="&Dvbbs.Boardid&" And GroupID="&Dvbbs.UserGroupID)
			If Not (Rs.Eof And Rs.Bof) Then
				Dvbbs.GroupSetting = Split(Rs(0),",")
			End If
			Set Rs=Nothing
		End If
		If Dvbbs.UserID>0 And InStr(IsGroupSetting,",0,")>0 Then
			Set Rs=Dvbbs.execute("Select Uc_Setting From Dv_UserAccess Where Uc_Boardid="&Dvbbs.BoardID&" And uc_UserID="&Dvbbs.Userid)
			If Not(Rs.Eof And Rs.Bof) Then
				Dvbbs.UserPermission=Split(Rs(0),",")
				Dvbbs.GroupSetting = Split(Rs(0),",")
				Dvbbs.FoundUserPer=True
			End If
			Set Rs=Nothing
		End If
	End If
	If Dvbbs.Boardmaster Then Exit Sub
	Call Chkboardlogin()
End Sub
Rem 能否进入论坛的判断
Public Sub Chkboardlogin()
	If Dvbbs.Board_Setting(1)="1" And Dvbbs.GroupSetting(37)="0" Then Dvbbs.AddErrCode(26)
	If Dvbbs.GroupSetting(0)="0"  Then Dvbbs.AddErrCode(27)
	'访问论坛限制(包括文章、积分、金钱、魅力、威望、精华、被删数、注册时间)
	Dim BoardUserLimited
	BoardUserLimited = Split(Dvbbs.Board_Setting(54),"|")
	If Ubound(BoardUserLimited)=8 Then
		'文章
		If Trim(BoardUserLimited(0))<>"0" And IsNumeric(BoardUserLimited(0)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户发贴最少为 <B>"&BoardUserLimited(0)&"</B> 才能进入&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(8))<Clng(BoardUserLimited(0)) Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户发贴最少为 <B>"&BoardUserLimited(0)&"</B> 才能进入&action=OtherErr"
		End If
		'积分
		If Trim(BoardUserLimited(1))<>"0" And IsNumeric(BoardUserLimited(1)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户积分最少为 <B>"&BoardUserLimited(1)&"</B> 才能进入&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(22))<Clng(BoardUserLimited(1)) Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户积分最少为 <B>"&BoardUserLimited(1)&"</B> 才能进入&action=OtherErr"
		End If
		'金钱
		If Trim(BoardUserLimited(2))<>"0" And IsNumeric(BoardUserLimited(2)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户金钱最少为 <B>"&BoardUserLimited(2)&"</B> 才能进入&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(21))<Clng(BoardUserLimited(2)) Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户金钱最少为 <B>"&BoardUserLimited(2)&"</B> 才能进入&action=OtherErr"
		End If
		'魅力
		If Trim(BoardUserLimited(3))<>"0" And IsNumeric(BoardUserLimited(3)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户魅力最少为 <B>"&BoardUserLimited(3)&"</B> 才能进入&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(23))<Clng(BoardUserLimited(3)) Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户魅力最少为 <B>"&BoardUserLimited(3)&"</B> 才能进入&action=OtherErr"
		End If
		'威望
		If Trim(BoardUserLimited(4))<>"0" And IsNumeric(BoardUserLimited(4)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户威望最少为 <B>"&BoardUserLimited(4)&"</B> 才能进入&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(24))<Clng(BoardUserLimited(4)) Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户威望最少为 <B>"&BoardUserLimited(4)&"</B> 才能进入&action=OtherErr"
		End If
		'精华
		If Trim(BoardUserLimited(5))<>"0" And IsNumeric(BoardUserLimited(5)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户精华最少为 <B>"&BoardUserLimited(5)&"</B> 才能进入&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(28))<Clng(BoardUserLimited(5)) Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户精华最少为 <B>"&BoardUserLimited(5)&"</B> 才能进入&action=OtherErr"
		End If
		'删贴
		If Trim(BoardUserLimited(6))<>"0" And IsNumeric(BoardUserLimited(6)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户被删贴少于 <B>"&BoardUserLimited(6)&"</B> 才能进入&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(27))>Clng(BoardUserLimited(6)) Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户被删贴少于 <B>"&BoardUserLimited(6)&"</B> 才能进入&action=OtherErr"
		End If
		'注册时间
		If Trim(BoardUserLimited(7))<>"0" And IsNumeric(BoardUserLimited(7)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户注册时间大于 <B>"&BoardUserLimited(7)&"</B> 分钟才能进入&action=OtherErr"
			If DateDiff("s",Dvbbs.MyUserInfo(14),Now)<Clng(BoardUserLimited(7))*60 Then Response.redirect "showerr.asp?ErrCodes=<li>本版面设置了用户注册时间大于 <B>"&BoardUserLimited(7)&"</B> 分钟才能进入&action=OtherErr"
		End If
		
	End If
	'认证版块判断Board_Setting(2)
	If Dvbbs.Board_Setting(2)="1" Then
		If Dvbbs.UserID=0 Then
			Dvbbs.AddErrCode(24)
			Dvbbs.showerr()
		Else
			Dim Boarduser,Canlogin,i
			Canlogin = False
			BoardUser = Dvbbs.boarduser
			If Ubound(Boarduser)=-1 Then	'为空时值等于-1
				Canlogin = False
			Else
				For i = 0 To Ubound(Boarduser)
					If Trim(Lcase(Boarduser(i))) = Trim(Lcase(Dvbbs.MemberName)) Then
						Canlogin = True
						Exit For
					End If				
				Next
			End If
		End If
		'If Dvbbs.Board_Setting(46) <> "0"  And Not Canlogin Then
			'Response.Redirect "pay_boardlimited.asp?boardid=" & Dvbbs.BoardID
		If Not Canlogin Then
			Dvbbs.AddErrCode(25)	
		End If
	End If
	Dvbbs.showerr()
End Sub
%>