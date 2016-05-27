<%
'=========================================================
' File: Dv_ClsMain.asp
' Version:7.0 sp2
' Date: 2004-6-30
' Script Written by dvbbs.net
'=========================================================
' Copyright (C) 2003,2004 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
'========================================
' 更新说明，加强过滤，加入对Chr(0)的过滤=
' 同时解决封IP中伪造cookies信息         = 
' 和通过访问一下管理页躲过封IP的问题    =
'========================================
Dim Ad_3(100),i3
Class Cls_Forum
	Rem Const
	Public BoardID,SqlQueryNum,Forum_Info,Forum_Setting,Forum_user,Forum_Copyright,Forum_ads,Forum_ChanSetting
	Public Forum_sn,Forum_Version,Stats,StyleName,ErrCodes,NowUseBBS,Cookiepath
	Public lanstr,mainhtml,mainsetting,sysmenu,mainpic
	Public MyUserInfo,UserToday,BoardJumpList,BoardList,CacheData,Maxonline
	Public UserGroupID,Lastlogin,GroupSetting,FoundUserPer
	Public Vipuser,Boardmaster,Superboardmaster,Master,FoundIsChallenge,FoundUser
	Public ScriptName,MemberName,MemberWord,MemberClass,UserHidden,UserID,UserTrueIP,UserPermission
	Public sendmsgnum,sendmsgid,sendmsguser,Page_Admin,Forum_AdLoop3
	Public BadWords,rBadWord,Forum_emot,Forum_PostFace,Forum_UserFace,SkinID,Forum_PicUrl
	Private adcode_1,adcode_2,adcode_4,ScriptTrueUrl,Forum_CSS,Main_Sid,ReloadCount,Nowstats,CssID
	Public Reloadtime,CacheName,savelog
	Private LocalCacheName,Cache_Data,IsTopTable,CookiesSid,BoardInfoData
	Public Board_Setting,boarduser,LastPost,Board_Ads,Board_user,BoardType,IsGroupSetting,BoardMasterList,Board_Data,Sid,Boardreadme,BoardRootID,BoardParentID
	Rem Sub 
	Private Sub Class_Initialize()
		savelog=0'设置为1的时候会记录攻击或错误错信息。
		SqlQueryNum = 0
		Reloadtime=14400
		CacheName=Replace(Replace(Replace(Server.MapPath("index.asp"),"index.asp",""),":",""),"\","")
		ReloadCount=0
		IsTopTable = 0
		Forum_sn = LCase(Replace(Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL"),Split(request.ServerVariables("SCRIPT_NAME"),"/")(ubound(Split(request.ServerVariables("SCRIPT_NAME"),"/"))),""))
		Vipuser = False:Boardmaster = False
		Superboardmaster = False:Master = False:FoundIsChallenge = False:FoundUser = False
		BoardID = Request("BoardID")
		If IsNumeric(BoardID) = 0 or BoardID = "" Then BoardID = 0
		BoardID = Clng(BoardID)
		MemberName = checkStr(Trim(Request.Cookies(Forum_sn)("username")))
		MemberWord = checkStr(Trim(Request.Cookies(Forum_sn)("password")))
		UserHidden = Request.Cookies(Forum_sn)("userhidden")
		UserID = Trim(Request.Cookies(Forum_sn)("UserID"))
		If IsNumeric(UserHidden) = 0 or Userhidden = "" Then UserHidden = 2
		If IsNumeric(UserID) = 0 Or UserID="" Then UserID=0
		UserID = Clng(UserID)
		UserTrueIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If UserTrueIP = "" Then UserTrueIP = Request.ServerVariables("REMOTE_ADDR")
		UserTrueIP = CheckStr(UserTrueIP)
		Dim Tmpstr
		Tmpstr = Request.ServerVariables("PATH_INFO")
		Tmpstr = Split(Tmpstr,"/")
		ScriptName = Lcase(Tmpstr(UBound(Tmpstr)))
		MemberClass = checkStr(Request.Cookies(Forum_sn)("userclass"))
		Page_Admin=False
		If InStr(ScriptName,"showerr")>0 Or InStr(ScriptName,"login")>0 Or InStr(ScriptName,"admin_")>0 Then Page_Admin=True
		sendmsgnum=0:sendmsgid=0:sendmsguser=""	
	End Sub
	Private Sub class_terminate()
		If IsObject(Conn) Then Conn.Close:Set Conn = Nothing
	End Sub
	Public Property Let Name(ByVal vNewValue)
		LocalCacheName = LCase(vNewValue)
	End Property
	Public Property Let Value(ByVal vNewValue)
		If LocalCacheName<>"" Then 
			ReDim Cache_Data(2)
			Cache_Data(0)=vNewValue
			Cache_Data(1)=Now()
			Application.Lock
			Application(CacheName & "_" & LocalCacheName) = Cache_Data
			Application.unLock
		Else
			Err.Raise vbObjectError + 1, "DvbbsCacheServer", " please change the CacheName."
		End If
	End Property
	Public Property Get Value()
		If LocalCacheName<>"" Then 
			Cache_Data=Application(CacheName & "_" & LocalCacheName)	
			If IsArray(Cache_Data) Then
				Value=Cache_Data(0)
			Else
				Err.Raise vbObjectError + 1, "DvbbsCacheServer", " The Cache_Data("&LocalCacheName&") Is Empty."
			End If
		Else
			Err.Raise vbObjectError + 1, "DvbbsCacheServer", " please change the CacheName."
		End If
	End Property
	Public Function ObjIsEmpty()
		ObjIsEmpty=True	
		Cache_Data=Application(CacheName & "_" & LocalCacheName)
		If Not IsArray(Cache_Data) Then Exit Function
		If Not IsDate(Cache_Data(1)) Then Exit Function
		If DateDiff("s",CDate(Cache_Data(1)),Now()) < (60*Reloadtime) Then ObjIsEmpty=False		
	End Function
	Public Sub DelCahe(MyCaheName)
		Application.Lock
		Application.Contents.Remove(CacheName&"_"&MyCaheName)
		Application.unLock
	End Sub
	'取得基本设置数据
	Public Sub GetForum_Setting()
		Name="setup"
		If ObjIsEmpty() Then ReloadSetup()
		CacheData=value
		'每日更新数据
		'DelCahe "Date"
		'第一次起用论坛或者重启IIS的时候加载缓存
		Name="Date"
		If ObjIsEmpty() Then
			value=Date()
			Call ReloadAllForumInfo
			Call ReloadAllBoardInfo
		Else
			If Cstr(value) <> Cstr(Date()) Then
				Call ReloadAllForumInfo
				Call ReloadAllBoardInfo
				Name="setup"
				Call ReloadSetup()
				CacheData=value
			End If
		End If
		Dim Setting
		Setting=CacheData(1,0)
		Setting = Split(Setting,"|||")
		Forum_Info = Setting(0)
		Forum_Info = Split (Forum_Info,",")
		Forum_Setting = Setting(1)
		Forum_Setting = Split (Forum_Setting,",")
		Forum_user = Setting(2)
		Forum_user = Split (Forum_user,",")
		Forum_Copyright = Setting(3)
		Forum_ChanSetting = CacheData(24,0)
		Forum_ChanSetting = Split(Forum_ChanSetting,",")
		Forum_Version = CacheData(18,0)
		BadWords = Split(CacheData(3,0),"|")
		rBadWord = Split(CacheData(4,0),"|")
		Main_Sid=CacheData(17,0)
		Maxonline = CacheData(5,0)
		NowUseBBS = CacheData(19,0)
		Cookiepath = CacheData(26,0)
		'IP锁定
		If Request.Cookies(Forum_sn & "Kill")("kill") = "1" Then
			If Not Page_Admin Then
				Response.Redirect "showerr.asp?action=iplock"
				Exit Sub
			End If
		ElseIf Not ( Request.Cookies(Forum_sn & "Kill")("kill") = "0" And Not IsEmpty(Session(CacheName & "UserID")) ) Then
			Call ChecKIPlock
			If Request.Cookies(Forum_sn & "Kill")("kill") = "1" Then
				If Not Page_Admin Then
					Response.Redirect "showerr.asp?action=iplock"
					Exit Sub
				End If
			End If
		End If
		'关闭论坛相关部分
		If Forum_Setting(21)="1" And Not Page_Admin Then Response.redirect "showerr.asp?action=stop"		
		Dim OpenTime,ischeck
		'判断BoardID的值，获取对应的设置
		If BoardID>0 Then
			If Not InStr((","&cachedata(27,0)&","),(","&BoardID&","))>0 Then
				Response.Write "错误的版面参数"
  				Response.End
			End If
			Name="BoardInfo_" & BoardID
  			If ObjIsEmpty() Then ReloadBoardInfo(BoardID)
			Board_Data = Value
			boarduser = Split(Board_Data(13,0) & "",",")
			Board_Ads = Split(Board_Data(17,0),"$")
			Board_user = Split(Board_Data(18,0),",")
			Forum_user = Board_User
			board_Setting = Split(Board_Data(16,0),",")
			LastPost = Split(Board_Data(14,0),"$")
			BoardType = Board_Data(1,0)
			IsGroupSetting = Board_Data(19,0)
			BoardMasterList = Board_Data(8,0)
			BoardRootID = Board_Data(5,0)
			BoardParentID=Board_Data(2,0)
			Sid = Board_Data(15,0)
			Boardreadme=Board_Data(7,0)
			If Len(Board_Setting(22))< 24 Then
				Board_Setting(22)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
			End If
			OpenTime=Split(Board_Setting(22),"|")
			setting=Board_Setting(21)
			Forum_ads =Board_Ads
			ischeck=Clng(Board_Setting(18))
			If Board_Setting(50)<>"0" And Board_Setting(50)<>"" Then Response.Redirect Board_Setting(50)
			If IsNumeric(Board_Data(21,0)) And CLng(Board_Data(6,0)) > 0 And CInt(Board_Data(4,0))< 2 Then Call LoadBoardList(BoardID,1)
			If IsNumeric(Board_Data(26,0)) And CLng(Board_Data(6,0)) > 0 And CInt(Board_Data(4,0))< 2 Then Call LoadBoardList(BoardID,0)
			'杨铮注：Board_Data(6,0) 为子论坛个数，当为空值时便会出错，检查 Dv_Board 表 Child 字段。
		Else
			Forum_ads =  CacheData(2,0)
			Forum_ads = Split(Forum_ads,"$")
			If Len(Forum_Setting(70))< 24 Then
				Forum_Setting(70)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
			End If
			OpenTime=Split(Forum_Setting(70),"|")
			setting=Forum_Setting(69)
			ischeck=Forum_Setting(26)
			If Not IsNumeric(ischeck) Then ischeck=0
			ischeck=CLng(ischeck)		
		End If
		'定时开放判断
		If Not Page_Admin And Cint(setting)=1 Then
			If OpenTime(Hour(Now))="0" Then
				Response.redirect "showerr.asp?action=stop&boardid="&Dvbbs.BoardID&""	 
			End If
		End If
		'在线人数限制
		If ischeck > 0 And Not Page_Admin Then
			If MyBoardOnline.Forum_Online > ischeck And BoardID=0 Then
				If Not IsONline(Membername,1) Then Response.Redirect "showerr.asp?action=limitedonline&lnum="&ischeck
			End If
			If BoardID<> 0 Then
				If (Not IsONline(Membername,1)) And MyBoardOnline.Board_Online > ischeck Then Response.Redirect "showerr.asp?action=limitedonline&lnum="&ischeck
			End If
		End If
		If Forum_ChanSetting(0)="1" And Forum_ChanSetting(1)="1" Then Get_Chan_Ad
	End Sub
	Public Function IsReadonly()
		IsReadonly=False
		Dim TimeSetting
		If Forum_Setting(69)="2" Then
			TimeSetting=split(Forum_Setting(70),"|")
			If TimeSetting(Hour(Now))="0" Then
				IsReadonly=True
				Exit Function
			End If
		End If
		If boardid<>0 Then 
			If Board_Setting(21)="2" Then
				TimeSetting=split(Board_Setting(22),"|")
				If TimeSetting(Hour(Now))="0" Then
					IsReadonly=True
				End If
			End If
		End If 
	End Function
	Public Function IsONline(UserName,action)
		IsONline=False
		If Trim(UserName)="" Then Exit Function
		If IsArray(Session(CacheName & "UserID")) And action=1 Then
			If Session(CacheName & "UserID")(0)="Dvbbs" Then
				IsONline=True:Exit Function 
			End If
		End If
		Dim Rs
		Set Rs =Execute("Select Count(*) From Dv_Online Where Username='"&UserName&"'")
		If Rs(0)<> 0 Then IsONline=True
		Set rs=Nothing  
	End Function  
	Public Sub ReloadSetup()
		Dim SQL,Rs,i
		SQL = "Select * from [Dv_setup] "
		Set Rs = Execute(SQL)
		value = Rs.GetRows(1)
		Set Rs = Nothing
	End Sub 
	Public Sub ReloadTemplateslist()		
		Dim Rs,SQL,tmpdata
		SQL = "select ID,StyleName from [Dv_Style]"
		Set Rs = Execute(SQL)
		tmpdata = Rs.GetString(,,"|||","@@@","")
		tmpdata = Left(tmpdata,Len(tmpdata)-3)	
		Set Rs = Nothing 
		value=tmpdata
	End Sub
	Public Sub LoadTemplates(Page_Fields)
		Dim Style_Pic,Main_Style,TempStyle
		CookiesSid = Request.Cookies("skin")("SkinID_"&BoardID)
		If Not IsNumeric(CookiesSid) Or CookiesSid = "" Then
			If BoardID = 0 Then 
				SkinID = Main_Sid
			Else
				SkinID = sid
			End If
		Else
			SkinID=CookiesSid
		End If
		SkinID=CLng(SkinID)
		Name="StyleName"&SkinID
		If ObjIsEmpty() Then TemplatesToCache ("StyleName")
		StyleName=value
		Name="Forum_CSS"&SkinID
		If ObjIsEmpty() Then TemplatesToCache ("Forum_CSS")
		'风格换肤修改
		CssID=Request.Cookies("skin")("cssid_"&BoardID)
		If Not IsNumeric(CssID) OR CssID="" Then 
			If boardid=0 Then
				CssID=CacheData(30,0)
			Else
				CssID=Board_Data(25,0)
			End If
		End If
		If CssID="" Or Not IsNumeric(CssID) Then CssID=0
		CssID=CLng(CssID)
		TempStyle = value
		TempStyle = Split(TempStyle,"@@@")
		If CssID > UBound(Split(TempStyle(1),"|||"))-1 Then
			CssID = 0
		End If
		Forum_CSS = Split(TempStyle(1),"|||")(CssID)		'风格内容
		Forum_PicUrl = Split(TempStyle(2),"|||")(CssID)		'图片路径
		Name = "Main_Style"&SkinID
		If ObjIsEmpty() Then TemplatesToCache ("Main_Style")
		Main_Style = Replace(value,"{$PicUrl}",Forum_PicUrl)		'风格图片路径替换
		If Not (Instr(ScriptName,"index")>0 Or Instr(ScriptName,"list")>0 Or Page_Admin) Then
			Name = "Style_Pic"&SkinID
			If ObjIsEmpty() Then TemplatesToCache ("Style_Pic")
			Style_Pic = Replace(value,"{$PicUrl}",Forum_PicUrl)		'风格图片路径替换
			Style_Pic = Split(Style_Pic,"@@@")
			Dim TmpArray(10),i
			For i=0 to UBound(Style_Pic)
				TmpArray(i) = Style_Pic(i)
			Next 
			Forum_UserFace = TmpArray(0)
			Forum_PostFace = TmpArray(1)
			Forum_Emot = TmpArray(2)
		End If
		If Page_Fields<>"" Then
			Name="page_"&Page_Fields&SkinID
			If ObjIsEmpty() Then TemplatesToCache ("page_"&Page_Fields)
			Template.value = value
		End If
		Main_Style = Split(Main_Style,"@@@")
		mainhtml = Split(Main_Style(0),"|||")
		lanstr = Split(Main_Style(1),"|||")
		mainpic = Split(Main_Style(2),"|||")
		mainsetting = Split(mainhtml(0),"||")
		Forum_CSS = Replace(Forum_CSS,"{$width}",mainsetting(0))
		Forum_CSS = Replace(Forum_CSS,"{$PicUrl}",Forum_PicUrl)
	End Sub
	Public Sub TemplatesToCache(Page_Fields)
		Dim Rs,SQL
		SQL = "Select "&Page_Fields&" from [Dv_Style] where id = " & SkinID
		Set Rs = Execute(SQL)
		If Not Rs.EOF Then
			value=Rs(0)&""
		Else
			'处理错误
			If boardid<>0 Then
				If Cint(SkinID)=Cint(sid) Then Fixsid()
			Else
				If SkinID=CInt(CacheData(17,0)) Then
					Call FixSetupsid()		
				End if
			End If
			Response.redirect "cookies.asp?action=stylemod&SkinID=0&boardid="&Boardid					
		End If
		Set Rs = Nothing
	End Sub
	Private Sub Fixsid()
		Dim Rs,SQL
		SQL = "Select Count(*) from [Dv_Style] where id = " & sid
		Set Rs = Execute(SQL)
		If Rs(0)=0 Then
			'把该版的SID更新为系统缺省的值
			Execute("Update Dv_Board Set Sid="&CLng(CacheData(17,0))&" where BoardID="&BoardID&"")
			'更新该版面的缓存
			ReloadBoardCache BoardID,CacheData(17,0),15,0
		End If
		Set Rs = Nothing
	End Sub 
	Private Sub FixSetupsid()
		Dim Rs,SQL
		SQL = "Select Top 1 ID from [Dv_Style] Order by ID"
		Set Rs = Execute(SQL)
		If Rs.EOF Then
			Response.Write "论坛模板数据是空的，请添加。"
			Response.End 	
		Else
			ReloadSetupCache Rs(0),17
			Execute("Update Dv_Setup Set Forum_Sid="&Rs(0)&"")
		End If
		Set rs=Nothing 
	End Sub 
	Rem 判断发言是否来自外部
	Public Function ChkPost()
		Dim server_v1,server_v2
		Chkpost=False 
		server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
		server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
		If Mid(server_v1,8,len(server_v2))=server_v2 Then Chkpost=True 
	End Function
	'每日更新信息，简单更新
	Public Sub ReloadAllForumInfo()
		'数据库部分
		If value <> "1900-1-1" Then 
			value="1900-1-1"
			Dim Rs,LastPostInfo,TempStr,i
			Dim Forum_YesterdayNum,Forum_TodayNum,Forum_LastPost,Forum_MaxPostNum,Forum_MaxPostDate
			Set Rs=Execute("Select Top 1 Forum_YesterdayNum,Forum_TodayNum,Forum_LastPost,Forum_MaxPostNum From Dv_Setup")
			Forum_YesterdayNum=Rs(0)
			Forum_TodayNum=Rs(1)
			Forum_LastPost=Rs(2)
			Forum_MaxPostNum=Rs(3)
			Set Rs=Nothing
			LastPostInfo = Split(Forum_LastPost,"$")
			If Not IsDate(LastPostInfo(2)) Then LastPostInfo(2)=Now()	
			If DateDiff("d",CDate(LastPostInfo(2)),Now())<>0 Then'最后发帖时间不是今天，	
				TempStr=LastPostInfo(0)&"$"&LastPostInfo(1)&"$"&Now()&"$"&LastPostInfo(3)&"$"&LastPostInfo(4)&"$"&LastPostInfo(5)&"$"&LastPostInfo(6)&"$"&LastPostInfo(7)
				Execute("Update Dv_Setup Set Forum_YesterdayNum="&Forum_TodayNum&",Forum_LastPost='"&TempStr&"',Forum_TodayNum=0")
				ReloadSetupCache 0,9
				ReloadSetupCache Forum_TodayNum,11
				ReloadSetupCache TempStr,15
			End If
			If Forum_TodayNum >Forum_MaxPostNum Then
				Execute("Update Dv_Setup Set Forum_MaxPostNum=Forum_TodayNum,Forum_MaxPostDate="&SqlNowString)
				ReloadSetupCache Forum_TodayNum,12'日最高发帖
				ReloadSetupCache Now(),13 '最高发帖日期
			End If
			LoadBoardsInfo()
		End If
		Name="Date"
		value=Date()
	End Sub
	'使用一个查询更新所有版面的缓存
	Public Sub LoadBoardsInfo()
		Dim Rs,BoardData(26,0),i,GetData,SQL,LastPostInfo,TempStr,IsUpdate
		IsUpdate=0
		SQL="select boardid,BoardType,ParentID,ParentStr,Depth,RootID,Child,readme,BoardMaster,PostNum,TopicNum,indexIMG,todayNum,boarduser,LastPost,Sid,Board_Setting,Board_Ads,Board_user,IsGroupSetting,BoardTopStr,BoardID As TempStr,BoardID As TempStr1,BoardID As TempStr2,BoardID As TempStr3,cid from Dv_board"
		If Not IsObject(Conn) Then ConnectionDatabase
		Set Rs=Server.CreateObject("ADODB.RecordSet")
		Rs.Open SQL,Conn,1,3
		Do While Not Rs.Eof
			LastPostInfo = Split(Rs(14),"$")
			If Not IsDate(LastPostInfo(2)) Then LastPostInfo(2)=Now()
			If DateDiff("d",LastPostInfo(2),Now())<>0 Then
				Rs("LastPost")=LastPostInfo(0)&"$"&LastPostInfo(1)&"$"&LastPostInfo(2)&"$"&LastPostInfo(3)&"$"&LastPostInfo(4)&"$"&LastPostInfo(5)&"$"&LastPostInfo(6)&"$"&LastPostInfo(7)
				Rs("TodayNum")=0
				Rs.UpDate
				IsUpdate=1
			End If
			Name="BoardInfo_" & Rs(0)
			For i=0 to Rs.Fields.Count-1
				BoardData(i,0)=Rs(i)
			Next
			value = BoardData
			GetData = Value
			IsUpdate=0
		Rs.MoveNext
		Loop
		Rs.Close
		Set Rs=Nothing
	End Sub 
	'更新总设置表部分缓存数组，入口：更新内容、数组位置
	Public Function ReloadSetupCache(MyValue,N)
		CacheData(N,0) = MyValue
		Name="setup"
		value=CacheData
	End Function
	'更新用户资料缓存(缓存用户名,是否需要添加)[0=不添加,只作清理,1=需要添加]
	Public Sub NeedUpdateList(username,act)
		Dim Tmpstr,TmpUsername
		Name="NeedToUpdate"
		If ObjIsEmpty() Then 
		Value=""
		End If
		Tmpstr=Value
		TmpUsername=","&username&","
		Tmpstr=Replace(Tmpstr,TmpUsername,",")
		Tmpstr=Replace(Tmpstr,",,",",")
		IF act=1 Then 
			If IsONline(username,0) Then
				If Tmpstr="" Then
					Tmpstr=TmpUsername
				Else
					Tmpstr=Tmpstr&TmpUsername
				End If
			End If
		End If
		Tmpstr=Replace(Tmpstr,",,",",")
		Value=Tmpstr
	End Sub
	'写入客人session
	Public Sub LetGuestSession()
		Dim StatUserID,UserSessionID
		StatUserID = checkStr(Trim(Request.Cookies(Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
		StatUserID = Ccur(StatUserID)
		Response.Cookies(Forum_sn).Expires=DateAdd("s",3600,Now())
		Response.Cookies(Forum_sn).path=cookiepath
		Response.Cookies(Forum_sn)("StatUserID") = StatUserID
		'客人=SessionID+活动时间+发帖时间+版面ID
		StatUserID = StatUserID & "_" & Now & "_" & Now & "_" & BoardID
		Session(CacheName & "UserID") = Split(StatUserID,"_")
	End Sub 
	'根据页面来判断是否需要执行TrueCheckUserLogin
	Public Function NeedChecklongin()
		NeedChecklongin=True
		If UserID>0 Then
			If InStr(ScriptName,"admin_")>0 Then Exit Function
			Dim pagelist
			pagelist=",post.asp,usermanager.asp,mymodify.asp,modifypsw.asp,modifyadd.asp,usersms.asp,"
			pagelist=pagelist & "friendlist.asp,favlist.asp,myfile.asp,friendlist.asp,recycle.asp,"
			pagelist=pagelist & "fileshow.asp,bbseven.asp,dispuser.asp,savepost.asp,"
			If InStr(pagelist,","&ScriptName&",")>0 Then Exit Function
		End If
		NeedChecklongin=False
	End Function 
	'验证用户登陆
	Public Sub CheckUserLogin()
		If Not IsArray(Session(CacheName & "UserID")) Then
			If UserID > 0 Then 
				TrueCheckUserLogin
			Else
				Call LetGuestSession()
			End If	
		Else
			If UserID >0  Then
				Dim NeedToUpdate,toupdate
				toupdate=False
				Name="NeedToUpdate"
				If Not ObjIsEmpty() Then 
					NeedToUpdate=","&Value&","
					If InStr(NeedToUpdate,","&MemberName&",")>0 Then
						Call NeedUpdateList(MemberName,0)
						toupdate=True
					End If
				End If
				
				If NeedChecklongin Or (UserID >0 And Not Session(CacheName & "UserID")(0)="Dvbbs" ) Or toupdate Then
					TrueCheckUserLogin
				End If
			End If
		End If
		If Session(CacheName & "UserID")(0) = "Dvbbs" Then
			GetCacheUserInfo
		Else
			MyUserInfo = Session(CacheName & "UserID")
			UserGroupID = 7
			Lastlogin = Now()
		End If	
		GetGroupSetting
	End Sub
	'系统分配随机密码
	Public Function Createpass()
		Dim Ran,i,LengthNum
		LengthNum=16
		Createpass=""
		For i=1 To LengthNum
			Randomize
			Ran = CInt(Rnd * 2)
			Randomize
			If Ran = 0 Then
				Ran = CInt(Rnd * 25) + 97
				Createpass =Createpass& UCase(Chr(Ran))
			ElseIf Ran = 1 Then
				Ran = CInt(Rnd * 9)
				Createpass = Createpass & Ran
			ElseIf Ran = 2 Then
				Ran = CInt(Rnd * 25) + 97
				Createpass =Createpass& Chr(Ran)
			End If
		Next
	End Function
	'更新用户验证密码
	Public Sub NewPassword()
		If UserID=0 Then Exit Sub	
		Response.Write "<iframe width=""0"" height=""0"" src=""newpass.asp"" name=""Dvnewpass""></iframe>"
	End Sub
	Public Sub NewPassword0()
		If UserID=0 Then Exit Sub
		If Not Response.IsClientConnected Then
			Exit Sub
		End If
		Dim TruePassWord,usercookies
		usercookies=Request.Cookies(Dvbbs.Forum_sn)("usercookies")
		TruePassWord=Createpass
		If (Isnull(usercookies) or usercookies="") And Not Isnumeric(usercookies) Then usercookies=0
		Select Case Cint(usercookies)
			Case 0
				Response.Cookies(Forum_sn)("usercookies") = usercookies
			Case 1
   				Response.Cookies(Forum_sn).Expires=Date+1
				Response.Cookies(Forum_sn)("usercookies") = usercookies
			Case 2
				Response.Cookies(Forum_sn).Expires=Date+31
				Response.Cookies(Forum_sn)("usercookies") = usercookies
			Case 3
				Response.Cookies(Forum_sn).Expires=Date+365
				Response.Cookies(Forum_sn)("usercookies") = usercookies
		End Select
		Response.Cookies(Forum_sn).path=cookiepath
		Response.Cookies(Forum_sn)("username") = MemberName
		Response.Cookies(Forum_sn)("UserID") = UserID
		Response.Cookies(Forum_sn)("userclass") = checkStr(Request.Cookies(Forum_sn)("userclass"))
		Response.Cookies(Forum_sn)("userhidden") = UserHidden
		Response.Cookies(Forum_sn)("password") = TruePassWord
		'检查写入是否成功如果成功则更新数据
		If checkStr(Trim(Request.Cookies(Forum_sn)("password")))=TruePassWord Then
			Execute("UpDate [Dv_user] Set TruePassWord='"&TruePassWord&"' where UserID="&UserID)
			MemberWord = TruePassWord
			Dim iUserInfo
			iUserInfo = Session(CacheName & "UserID")
			iUserInfo(35) = TruePassWord
			Session(CacheName & "UserID") = iUserInfo
		End If
	End Sub
	Public Sub TrueCheckUserLogin()
	'Session(CacheName & "UserID")用户资料=0dvbbs+1刷新时间+2发帖时间+3所在版面ID+4用户ID+5用户名+6用户密码+7用户邮箱+8用户文章数+9用户主题数+10用户性别+11用户头像+12用户头像宽+13用户头像高+14用户注册时间+15用户最后登陆时间+16用户登陆次数+17用户状态+18用户等级+19用户组ID+20用户组名+21用户金钱+22用户积分+23用户魅力+24用户威望+25用户生日+26最后登陆IP+27用户被删除数+28用户精华数+29用户隐身状态+30用户短信情况+31用户阳光会员+32用户手机+33用户组图标+34用户头衔+35验证密码+36用户今日信息+37用户待发帖子数据+38Dvbbs
		Dim Rs,SQL
		Sql="Select UserID,UserName,UserPassword,UserEmail,UserPost,UserTopic,UserSex,UserFace,UserWidth,UserHeight,JoinDate,LastLogin,UserLogins,Lockuser,Userclass,UserGroupID,UserGroup,userWealth,userEP,userCP,UserPower,UserBirthday,UserLastIP,UserDel,UserIsBest,UserHidden,UserMsg,IsChallenge,UserMobile,TitlePic,UserTitle,TruePassWord,UserToday"
		Sql=Sql+" From [Dv_User] Where UserID = " & UserID
		Set Rs = Execute(Sql)
		If Rs.Eof And Rs.Bof Then
			Rs.Close:Set Rs = Nothing
			UserID = 0
			EmptyCookies
			LetGuestSession()
		Else
			MyUserInfo=Rs.GetString(,1, "|||","","")
			Rs.Close:Set Rs = Nothing
			If IsArray(Session(CacheName & "UserID")) Then

				MyUserInfo = "Dvbbs|||"& Now & "|||" & Session(CacheName & "UserID")(2) &"|||"& BoardID &"|||"& MyUserInfo &"||||||Dvbbs"
			Else
				MyUserInfo = "Dvbbs|||"& Now & "|||" & DateAdd("s",-3600,Now()) &"|||"& BoardID &"|||"& MyUserInfo &"||||||Dvbbs"
			End If
			MyUserInfo = Split(MyUserInfo,"|||")
			If Trim(MyUserInfo(35)) = Memberword And Trim(MyUserInfo(5)) =Membername Then
				Session(CacheName & "UserID") = MyUserInfo
				Memberword = MyUserInfo(35)
				GetCacheUserInfo()
			Else
				If IsArray(Session(CacheName & "UserID")) Then
					If Session(CacheName & "UserID")(0)="Dvbbs" Then
						If Trim(Session(CacheName & "UserID")(4))=Trim(MyUserInfo(4)) And Trim(Session(CacheName & "UserID")(5))=Trim(MyUserInfo(5)) And Trim(Session(CacheName & "UserID")(6))=Trim(MyUserInfo(6)) Then
							Call NewPassword0()
						End If 
					Else
						UserID = 0
						EmptyCookies
						LetGuestSession()
					End If
				Else
					UserID = 0
					EmptyCookies
					LetGuestSession()
				End If 
			End If
		End If
	End Sub
	'用户登录成功后，采用本函数读取用户数组并判断一些常用信息
	Public Sub GetCacheUserInfo()
		MyUserInfo = Session(CacheName & "UserID")
		UserID = Clng(MyUserInfo(4))
		MemberName = MyUserInfo(5)
		Lastlogin = MyUserInfo(15)
		If Not IsDate(LastLogin) Then LastLogin = Now()
		UserGroupID = Cint(MyUserInfo(19))
		If Trim(MyUserInfo(36))="" Then
			Execute("Update [Dv_User] Set UserToday='0|0|0' Where UserID = " & UserID)
			MyUserInfo(36) = "0|0|0"
			UserToday = Split(MyUserInfo(36),"|")
		Else
			UserToday = Split(MyUserInfo(36),"|")
			If Ubound(UserToday) <> 2 Then
				Execute("Update [Dv_User] Set UserToday='0|0|0' Where UserID = " & UserID)
				MyUserInfo(36) = "0|0|0"
				UserToday = Split(MyUserInfo(36),"|")
			End If
		End If
		Select Case UserGroupID
		Case 4
			Vipuser = True
		Case 3
			Boardmaster = True
		Case 2
			Superboardmaster = True
		Case 1
			Master = True
		End Select
		If MyUserInfo(31) = "1" Then FoundIsChallenge = True
		If DateDiff("d",LastLogin,Now())<>0 Then
			Execute("Update [Dv_User] Set UserToday='0|0|0',LastLogin = " & SqlNowString & " Where UserID = " & UserID)
			MyUserInfo(36) = "0|0|0"
			LastLogin = Now()
		End If
		If Userhidden = 2 and DateDiff("s",Lastlogin,Now())>Clng(Forum_Setting(8))*60 Then
			Execute("Update [Dv_User] Set UserLastIP = '" & UserTrueIP & "',LastLogin = " & SqlNowString & " Where UserID = " & UserID)
			Lastlogin = Now()
		End If
		sendmsgnum=0:sendmsgid=0:sendmsguser=""
		If MyUserInfo(30)<>"" Then
			Dim Usermsg
			Usermsg=Split(MyUserInfo(30),"||")
			If Ubound(Usermsg)=2 Then
				sendmsgnum=Usermsg(0)
				sendmsgid=Usermsg(1)
				sendmsguser=Usermsg(2)
			End If
		End If
		FoundUser=True
		MyUserInfo(15)=Lastlogin
		Session(CacheName & "UserID")=MyUserInfo
	End Sub
	Public Sub EmptyCookies()
		Response.Cookies(Forum_sn)("usercookies") = 0
		Response.Cookies(Forum_sn).path=cookiepath
		Response.Cookies(Forum_sn)("username") = ""
		Response.Cookies(Forum_sn)("UserID") = 0
		Response.Cookies(Forum_sn)("userclass") = ""
		Response.Cookies(Forum_sn)("userhidden") = 2
		Response.Cookies(Forum_sn)("password") = ""
	End Sub
	Private Sub GetGroupSetting()
		Name="GroupSetting_"& UserGroupID
		If ObjIsEmpty() Then 
			Dim Rs,SQL
			SQL = "Select GroupSetting From [Dv_UserGroups] where UserGroupID = " & UserGroupID
			Set Rs = Execute(SQL)
			If Rs.Eof Then
				Set Rs=Nothing
				SQL = "Select GroupSetting From [Dv_UserGroups] where UserGroupID = 4"
				Set Rs = Execute(SQL)
				value=Rs(0)
			Else
				value=Rs(0)
			End If
		End If
		GroupSetting = Split(value,",")
		If Cint(GroupSetting(0))=0 And Not Page_Admin Then AddErrCode "8":Showerr()
		If BoardID <> 0 And Not ScriptName="showerr.asp" Then BoardInfoData=CheckBoardInfo()
	End Sub 
	Public Sub ActiveOnline()
		Dim ReflashPageLastTime,LastVisiBoardID
		ReflashPageLastTime = Session(CacheName & "UserID")(1)
		LastVisiBoardID = Clng(Session(CacheName & "UserID")(3))
		If Not IsDate(ReflashPageLastTime) Then ReflashPageLastTime = Now()
		'当在120秒内刷新同一个页面则不更新online数据
		If DateDiff("s",ReflashPageLastTime,Now()) < 120 And LastVisiBoardID = BoardID  And Not InStr(ScriptName,"showerr")>0 Then Exit Sub
		'更新数组
		ReflashPageLastTime = Session(CacheName & "UserID")
		ReflashPageLastTime(1) = Now()
		ReflashPageLastTime(3) = Dvbbs.BoardID
		Session(CacheName & "UserID") = ReflashPageLastTime
		UserActiveOnline
	End Sub
	Private Sub UserActiveOnline()
		Dim Actcome,SQl,Rs
		Dim MyGroupID,uip,BrowserType,StatsStr
			uip = UserTrueIP
        	StatsStr = Stats
        	StatsStr = Replace(StatsStr, "'", "")
        	StatsStr = Replace(StatsStr, Chr(0), "")
        	StatsStr = Replace(StatsStr, "--", "——")
        	StatsStr = Left(StatsStr, 250)
		If FoundIsChallenge and Cint(Forum_ChanSetting(0))=1 Then
			MyGroupID = 9999
		Else
			MyGroupID = UserGroupID
		End If
		If UserID = 0 Then
			Dim StatUserID
			StatUserID = Session(CacheName & "UserID")(0)
			SQL = "Select ID,Boardid From [Dv_Online] Where ID = " & Ccur(StatUserID)
			Set Rs = Execute(SQL)
			If Rs.Eof And Rs.Bof Then
				If CInt(Forum_Setting(36)) = 0 Then
					Actcome = ""
				Else
					Actcome = address(uip)
				End If
				Set BrowserType=new Cls_Browser
				SQL = "Insert Into [Dv_Online](ID,Username,Userclass,Ip,Startime,Lastimebk,Boardid,Browser,Stats,Usergroupid,Actcome,Userhidden) Values (" & StatUserID & ",'客人','客人','" & UserTrueIP & "'," & SqlNowString & "," & SqlNowString & "," & Boardid & ",'" & BrowserType.platform&"|"&BrowserType.Browser&BrowserType.version & "','" & StatsStr & "',7,'" & Actcome & "'," & Userhidden & ")"
				'更新缓存总在线数据
				MyBoardOnline.Forum_Online=MyBoardOnline.Forum_Online+1
				Name="Forum_Online"
				value=MyBoardOnline.Forum_Online
				Set BrowserType=Nothing 
			Else
				SQL = "Update [Dv_Online] Set Lastimebk = " & SqlNowString & ",Boardid = " & Boardid & ",Stats = '" & StatsStr & "' Where ID = " & Ccur(StatUserID)
			End If
			Rs.Close
			Set Rs = Nothing
			Execute(SQL)
		Else
			SQL = "Select ID,Boardid From [DV_Online] Where UserID = " & UserID
			Set Rs = Execute(SQL)
			If Rs.Eof And Rs.Bof Then
				If CInt(forum_setting(36)) = 0 Then
					Actcome = ""
				Else
					Actcome = address(uip)
				End If
				Set BrowserType=new Cls_Browser
				SQL = "Insert Into [Dv_Online](ID,Username,Userclass,Ip,Startime,Lastimebk,Boardid,Browser,Stats,Usergroupid,Actcome,Userhidden,UserID) Values (" & Session.SessionID & ",'" & Membername & "','" & Memberclass & "','" & UserTrueIP & "'," & SqlNowString & "," & SqlNowString & "," & Boardid & ",'" & BrowserType.platform&"|"&BrowserType.Browser&BrowserType.version & "','" & StatsStr & "'," & MyGroupID & ",'" & Actcome & "'," & Userhidden & "," & UserID & ")"
				Set BrowserType=Nothing
				'更新缓存总在线数据
				MyBoardOnline.Forum_Online=MyBoardOnline.Forum_Online+1
				Name="Forum_Online"
				Dvbbs.value=MyBoardOnline.Forum_Online
				'更新缓存总用户在线数据
				MyBoardOnline.Forum_UserOnline=MyBoardOnline.Forum_UserOnline+1
				Name="Forum_UserOnline"
				value=MyBoardOnline.Forum_UserOnline
			Else
				SQL = "Update [Dv_Online] Set Lastimebk = " & SqlNowString & ",Boardid = " & Boardid & ",Stats = '" & StatsStr & "' Where UserID = " & UserID
			End If
			Rs.Close
			Set Rs = Nothing
			Execute(SQL)
		End If	
		'更新在线峰值
		If CLng(MyBoardOnline.Forum_Online) > CLng(Maxonline) Then
			Execute("update [Dv_setup] set Forum_Maxonline="&CLng(MyBoardOnline.Forum_Online)&",Forum_MaxonlineDate="& SqlNowString) 
			CacheData(5,0)=MyBoardOnline.Forum_Online
			CacheData(6,0)=Now()
			Name="setup"
			value=CacheData
		End If 
		Rem 删除超时用户
		MyBoardOnline.OnlineQuery
	End Sub
	Public Sub Nav()
		Head()
		ShowTopTable()
		IsTopTable = 1
	End Sub
	Public Sub head()
		'建立缓存
		Name="head_"&SkinID
		If ObjIsEmpty() Then
			value= Replace(Replace(mainhtml(1),"{$keyword}",Replace(Forum_info(8),"|",",")),"{$description}",Forum_info(10))&vbNewLine
		End If
		Response.Write Value
		Nowstats=stats
		Dim re
		Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="<(.[^>]*)>"
			If BoardID > 0 And ScriptName<>"printpage.asp" Then Stats=BoardType&"-"&Stats
			Stats=re.Replace(Stats, "")
			re.Pattern=""""
			Stats=re.Replace(Stats, "&quot;")
		Set Re=Nothing
		Stats=Replace(Stats,chr(13),"")
		Response.Write "<title>"
		Response.Write Forum_Info(0)
		Response.Write "-"	
		Response.Write stats						
		Response.Write "</title>"
		Response.Write vbNewLine
		Response.Write Forum_CSS
		Response.Write mainhtml(2)
		'论坛防刷新设置
		If Cint(Forum_Setting(19))=1 And Not Page_Admin Then
			Dim DoReflashPage
			DoReflashPage=false
			If Trim(Forum_Setting(64))<>"" And InStr(LCase(Forum_Setting(64)),ScriptName) >0 Then DoReflashPage=True
			If (Not IsEmpty(Session(CacheName & "UserID")(1))) and Cint(Forum_Setting(20))>0 and DoReflashPage Then
				If DateDiff("s",Session(CacheName & "UserID")(1),Now())<Cint(Forum_Setting(20)) Then
					Response.Write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT="&Forum_Setting(20)&"><br>本页面起用了防刷新机制，请不要在"&Forum_Setting(20)&"秒内连续刷新本页面<BR>正在打开页面，请稍后……"
					Response.End
				Else
					DoReflashPage=Session(CacheName & "UserID")
					DoReflashPage(1)=Now()
					Session(CacheName & "UserID")=DoReflashPage
				End If
			ElseIf IsEmpty(Session(CacheName & "UserID")(1)) and Cint(Forum_Setting(20))>0 and DoReflashPage Then
				DoReflashPage=Session(CacheName & "UserID")
				DoReflashPage(1)=Now()
				Session(CacheName & "UserID")=DoReflashPage
			End If
		End If
	End Sub 
	Public Sub ShowTopTable()
		Dim TempStr,ForumMenu
		If Forum_ChanSetting(0)="1" And Forum_ChanSetting(1)="1"  Then 
			If Forum_ChanSetting(2)="1" Then
				TempStr = mainhtml(3)
				TempStr = Replace(TempStr,"{$top}",adcode_4)
			End If 
			If Forum_ChanSetting(3)="1" Then Forum_ads(0)=adcode_1
		End If
		Name="Templateslist"
		If ObjIsEmpty() Then ReloadTemplateslist()
		If UserID = 0 Then 
			sysmenu = mainhtml(7)
		Else
			sysmenu = Replace(mainhtml(6),"{$username}",Membername)
			If UserHidden=2 Then
				sysmenu = Replace(sysmenu,"{$hiddeninfo}",lanstr(3))
			Else
				sysmenu = Replace(sysmenu,"{$hiddeninfo}",lanstr(4))
			End If
			If Master Then
				sysmenu = Replace(sysmenu,"{$manageinfo}",mainhtml(10))
			Else
				sysmenu = Replace(sysmenu,"{$manageinfo}","")
			End If
			If Forum_ChanSetting(0)="1" Then
				Dim RayMenuInfo,RayMenu
				RayMenuInfo = Split(mainhtml(11),"||")
				If Forum_ChanSetting(2)=2 Then RayMenu = Replace(Replace(RayMenuInfo(3),"{$channame}",CacheData(23,0)),"{$forumurl}",Forum_Info(1))
				If FoundIsChallenge Then
					RayMenu = RayMenu & RayMenuInfo(1)
				Else
					RayMenu = RayMenu & RayMenuInfo(2)
				End If
				RayMenu = Replace(RayMenuInfo(0),"{$raymenu}",RayMenu)
				sysmenu = Replace(sysmenu,"{$raymenuinfo}",RayMenu)
			Else
				sysmenu = Replace(sysmenu,"{$raymenuinfo}","")
			End If
			sysmenu = Replace(sysmenu,"{$userid}",UserID)
		End If
		Dim tmpstr,i,outstr,ioutstr,SkinID1,Csslist,CssName,k,Tempstr1,Tempstr2
		mainhtml(9)=Replace(Replace(Replace(Replace(mainhtml(9),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")
		tmpstr=Split(value,"@@@")
		mainhtml(9) = Split(mainhtml(9),"||")
		outstr =mainhtml(9)(2)
		ioutstr=mainhtml(9)(0)
		mainhtml(9)(5)=Replace(mainhtml(9)(5),"{$boardid}",BoardID)
		SkinID1=SkinID
		For i = 0 to UBound(tmpstr)
			tmpstr(i) = Split(tmpstr(i),"|||")
			SkinID=tmpstr(i)(0)
			Name="Forum_CSS"&SkinID
			If ObjIsEmpty() Then
				TemplatesToCache ("Forum_CSS")
			End If
			Csslist=Value
			Csslist=split(Csslist,"@@@")
			CssName=split(Csslist(0),"|||")
			Tempstr2=Replace(mainhtml(9)(4),"{$skinid}",SkinID)
			If SkinID1<>Cint(tmpstr(i)(0)) Then 
				Tempstr2=Replace(Tempstr2,"{$skinname}",tmpstr(i)(1))
			Else
				mainhtml(9)(1)=Replace(mainhtml(9)(1),"{$skinname}",tmpstr(i)(1))
				mainhtml(9)(1)=Replace(mainhtml(9)(1),"{$alertcolor}",mainsetting(1))
				Tempstr2=Replace(Tempstr2,"{$skinname}",mainhtml(9)(1))
			End If
			Tempstr1=""
			For k=0 to UBound(CssName)-1
				If k=CssID And SkinID1=Cint(tmpstr(i)(0)) Then 
					mainhtml(9)(6)=Replace(mainhtml(9)(6),"{$alertcolor}",mainsetting(1))
					Tempstr1=Tempstr1&Replace(Replace(Replace(mainhtml(9)(6),"{$skinid}",SkinID),"{$cssid}",k),"{$cssname}",CssName(k))
				Else
					Tempstr1=Tempstr1&Replace(Replace(Replace(mainhtml(9)(5),"{$skinid}",SkinID),"{$cssid}",k),"{$cssname}",CssName(k))
				End If	
			Next 
			Tempstr1=Replace(Tempstr2,"{$cssinfo}",Tempstr1)			
			ioutstr=ioutstr&Replace(mainhtml(9)(3),"{$csslist}",Tempstr1)
		Next	
		SkinID=SkinID1	
		outstr=Replace(outstr,"{$sylelist}",ioutstr)
		sysmenu = Replace(sysmenu,"{$syles}",outstr)
		TempStr = TempStr & mainhtml(4)
		TempStr = Replace(TempStr,"{$width}",mainsetting(0))
		TempStr = Replace(TempStr,"{$link}",Forum_Info(1))
		If Boardid<>0 Then 
			If Board_Setting(51)="" Or Board_Setting(51) = "0"  Then
				TempStr = Replace(TempStr,"{$logo}",Forum_Info(6))
			Else
				TempStr = Replace(TempStr,"{$logo}",Board_Setting(51))
			End If
		Else
			TempStr = Replace(TempStr,"{$logo}",Forum_Info(6))
		End If
		If Trim(Forum_info(7))<>"0" And Trim(Forum_info(7))<>""  Then
			TempStr = Replace(TempStr,"{$mailto}",Forum_Info(7))
		Else
			TempStr = Replace(TempStr,"{$mailto}","mailto:" & Forum_Info(5))
		End If
		TempStr = Replace(TempStr,"{$title}",Forum_Info(0) & "-" & Replace(stats,"'","\'"))
		TempStr = Replace(TempStr,"{$top_ads}",Forum_ads(0))
		TempStr = Replace(TempStr,"{$menu}",sysmenu)
		TempStr = Replace(TempStr,"{$boardid}",boardid)
		TempStr = Replace(TempStr,"{$alertcolor}",mainsetting(1))
		Name = "ForumPlusMenu"&SkinID
		If ObjIsEmpty() Then ReloadForumPlusMenu()
		ForumMenu = Value
		TempStr = Replace(TempStr,"{$plusmenu}",ForumMenu)
		Response.Write TempStr		
		TempStr = ""				
	End Sub 
	Public Sub Head_var(IsBoard,idepth,GetTitle,GetUrl)
		Dim NavStr,AllBoardList
		If Dvbbs.BoardID=0 Then
			BoardReadme=lanstr(2) & " <b>" & Forum_Info(0) & "</b>"
		End if
		If GroupSetting(37)="0" Then
			Name = "BoardJumpList_g"
			If ObjIsEmpty() Then LoadBoardJumpList(0)
		Else
			Name = "BoardJumpList"
			If ObjIsEmpty() Then LoadBoardJumpList(1)
		End If
		BoardJumpList = Value
		BoardJumpList = Replace(BoardJumpList,"{BoardID="&BoardID&"}","selected")
		If GroupSetting(37)="0" Then
			Name = "MyAllBoardList_g"
			If ObjIsEmpty() Then LoadAllBoardList(0)
		Else
			Name = "MyAllBoardList"
			If ObjIsEmpty() Then LoadAllBoardList(1)
		End If
		AllBoardList = Value
		If BoardID>0 Then
			NavStr = " <a href="&Forum_Info(11)&" onMouseOver=""showmenu(event,'"&AllBoardList&"',0)"" style=""CURSOR:hand"">"&Forum_info(0)&"</a> → "
		Else
			NavStr = " <a href="&Forum_Info(11)&">"&Forum_info(0)&"</a> → "
		End If
		If IsBoard=1 Then
			If GroupSetting(37)="0" Then
				BoardList = Board_Data(26,0)
			Else
				BoardList = Board_Data(21,0)
			End If
			BoardType = Replace(Replace(BoardType,Chr(39),"&#39;"),Chr(34), "&#34;")
			If BoardParentID=0 Then
				NavStr = NavStr & " <a href=""list.asp?boardid="&BoardID&""" onMouseOver=""showmenu(event,'"&BoardList&"',0)"">"&BoardType&"</a>"
			Else
				If ScriptName="dispbbs.asp" Then 
					NavStr = NavStr & BoardInfoData & " → <a href=""list.asp?boardid="&BoardID&"&page="&Request("page")&""">"&BoardType&"</a>"
				Else
					NavStr = NavStr & BoardInfoData & " → <a href=""list.asp?boardid="&BoardID&""">"&BoardType&"</a>"
				End If
			End If
			NavStr = NavStr & " → " & Nowstats
		Elseif IsBoard=2 Then
			NavStr = NavStr & Nowstats
		Else
			NavStr = NavStr & "<a href="&GetUrl&">"&GetTitle&"</a> → " & Nowstats
		End If
		BoardReadme=Replace(Replace(Replace(BoardReadme&"","\n",""),"\r",""),"\","")
		NavStr = Replace(mainhtml(5),"{$nav}",NavStr)
		NavStr = Replace(NavStr,"{$width}",mainsetting(0))
		NavStr = Replace(NavStr,"{$boardreadme}",BoardReadme)
		If UserID>0 Then
			'sendmsgnum,sendmsgid,sendmsguser
			IsBoard = Split(mainhtml(12),"||")
			If Clng(SendMsgNum)>0 Then
				BoardReadme = IsBoard(0)
				If Forum_Setting(10)=1 Then
					BoardReadme = BoardReadme & IsBoard(1) & IsBoard(2)
				Else
					BoardReadme = BoardReadme & IsBoard(2)
				End If
				BoardReadme = Replace(BoardReadme,"{$smsid}",sendmsgid)
				BoardReadme = Replace(BoardReadme,"{$sender}",sendmsguser)
				BoardReadme = Replace(BoardReadme,"{$newmsgnum}",sendmsgnum)
				NavStr = Replace(NavStr,"{$umsg}",BoardReadme)
			Else
				NavStr = Replace(NavStr,"{$umsg}",IsBoard(3))
			End If
		Else
			NavStr = Replace(NavStr,"{$umsg}","")
		End If
		NavStr = Replace(NavStr,"{$alertcolor}",mainsetting(1))
		NavStr = Replace(NavStr,"{$showstr}","")
		Response.Write NavStr
	End Sub
	Private Function LoadBoardJumpList(Act)'参数，1读全部，0读非隐藏
		Dim Forum_Boards,i,ii,Depth,Board_Datas,b_setting
		Forum_Boards=Split(CacheData(27,0),",")
		For i=0 To Ubound(Forum_Boards)
			Name="BoardInfo_" & Forum_Boards(i)
			If ObjIsEmpty() Then ReloadBoardInfo(Forum_Boards(i))
			Board_Datas = Value
			b_setting=split(Board_Datas(16,0),",")
			If b_setting(1)<>"1" Or Act=1 Then
				BoardJumpList = BoardJumpList & "<option value=""list.asp?boardid="&Forum_Boards(i)&""" {BoardID="&Forum_Boards(i)&"}>"
				Depth=Board_Datas(4,0)
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
				BoardJumpList = BoardJumpList & Replace(Replace(Board_Datas(1,0),Chr(39),"&#39;"),Chr(34), "&#34;") &"</option>"
			End If
		Next
		If Act=1 Then 
			Name="BoardJumpList"
		Else
			Name="BoardJumpList_g"
		End If
		value=BoardJumpList
		Forum_Boards=Null
		Board_Datas=Null
	End Function
	Private Function LoadAllBoardList(Act)'参数，1读全部，0读非隐藏
		Dim Forum_Boards,MyAllBoardList,i,ii,Depth,Board_Datas,b_setting
		Forum_Boards=Split(CacheData(27,0),",")
		For i=0 To Ubound(Forum_Boards)
			Name="BoardInfo_" & Forum_Boards(i)
			If ObjIsEmpty() Then ReloadBoardInfo(Forum_Boards(i))
			Board_Datas = Value
			b_setting=split(Board_Datas(16,0),",")
			If b_setting(1)<>"1" Or Act=1 Then
				Depth=Board_Datas(4,0)
				MyAllBoardList = MyAllBoardList & "<a href=list.asp?boardid="&Forum_Boards(i)&">"
				Select Case Depth
					Case 0
						MyAllBoardList = MyAllBoardList & "╋"
					Case 1
						MyAllBoardList = MyAllBoardList & "&nbsp;&nbsp;├"
				End Select
				If Depth>1 Then
					For ii=2 To Depth
						MyAllBoardList = MyAllBoardList & "&nbsp;&nbsp;│"
					Next
					MyAllBoardList = MyAllBoardList & "&nbsp;&nbsp;├"
				End If
				MyAllBoardList = MyAllBoardList & Server.htmlencode(Board_Datas(1,0)) & "</a><br>"
			End If
		Next
		If Act=1 Then
			Name="MyAllBoardList"
		Else
			Name="MyAllBoardList_g"
		End If
		value=Replace(Replace(MyAllBoardList,"'","\'"),Chr(34), "&#34;")
		Forum_Boards=Null
		Board_Datas=Null
	End Function
	Public Sub AddErrCode(ErrCode)
		If ErrCodes = "" Then
			ErrCodes = ErrCode
		Else
			ErrCodes = ErrCodes & "," & ErrCode
		End If
	End Sub 
	Public Sub Showerr()
		If ErrCodes<>"" Then Response.redirect "showerr.asp?BoardID="&boardid&"&ErrCodes="&ErrCodes&"&action="&server.URLEncode(Stats)
	End Sub 
	Public Sub Footer()
		Dim Tmp,CaCheInfo
		'CaCheInfo =  "<li>"
		'CaCheInfo = CaCheInfo & "共使用了" & Application.Contents.Count & "个缓存对象。"
		Tmp = mainhtml(8)
		If Forum_Setting(30) = "1" Then 
			Dim Endtime
			Endtime = Timer()	
			Tmp = Replace(Tmp,"{$runtime}","<br>执行时间：" & FormatNumber((Endtime-Startime)*1000,5) & "毫秒。查询数据库" & SqlQueryNum & "次。"& CaCheInfo)
		Else
			Tmp = Replace(Tmp,"{$runtime}","")
		End If
		Tmp = Replace(Tmp,"{$color}",mainsetting(1))
		Tmp = Replace(Tmp,"{$width}",mainsetting(0))
		Tmp = Replace(Tmp,"{$powered}","Powered By ：<a href = ""http://www.dvbbs.net/download.asp"" target = ""_blank"">Dvbbs Version " & Forum_Version & "</a> Sp2")
		Tmp = Replace(Tmp,"{$Footer_ads}",Forum_ads(1))
		If Forum_ChanSetting(0)="1" And Forum_ChanSetting(1)="1" And Forum_ChanSetting(4)="1" And IsTopTable=1 Then
			Tmp = Replace(Tmp,"{$ad}","<BR>" & adcode_2)
		Else
			Tmp = Replace(Tmp,"{$ad}","")
		End If 
		Tmp = Replace(Tmp,"{$copyright}",Forum_Copyright)
		Tmp = Replace(Tmp,"{$StyleName}",StyleName)
		If Forum_ChanSetting(0)="1" Then  
			Tmp = Replace(Tmp,"{$server}","<td align = right><a href = ""http://www.ray5198.com"" target = _blank title = ""本论坛所提供的互动服务由北京阳光加信科技有限公司提供""><img src = ""images/rayslogo.GIF"" border = 0></a></td>")
		Else
			Tmp = Replace(Tmp,"{$server}","")
		End If
		'Response.Write CaCheInfo
		'//------------------------------------------------------------------------------
		'//论坛访问量统系
		If ScriptName="list.asp" or ScriptName="index.asp" Then
		Dim RayPostAct,RayUpCount,RayMaxCount,Forum_url,RaySubjection,Board_Datas,FrameBody
		Dim PostStr
		RayMaxCount=100		'定义更新概率
		RaySubjection=False
		Forum_url=Get_ScriptNameUrl
		If ScriptName="index.asp" Then
			Name="RayUpCount"
			If Dvbbs.ObjIsEmpty() Then
				Value=1
			Else
				RayUpCount=Value
				If Not IsNumeric(RayUpCount) Then
					Value=1
				Else
					Value=RayUpCount+1
				End If
			End If
			RayUpCount=Value
			If RayUpCount >= RayMaxCount Then
				RaySubjection=True
				RayUpCount=1
				Value=1
				FrameBody="?PostType=0&forumname="&Server.htmlencode(Forum_Info(0))
				FrameBody=FrameBody+"&forumurl="&Forum_url
				FrameBody=FrameBody+"&forumlogincount="&Dvbbs.CacheData(10,0)
				FrameBody=FrameBody+"&foruminlinecount="&MyBoardOnline.Forum_Online
				FrameBody=FrameBody+"&forumtitlecount="&CacheData(8,0)
				FrameBody=FrameBody+"&forumvisitprob=1"
				FrameBody=FrameBody+"&forumemail="&Forum_Info(5)
				FrameBody=FrameBody+"&forumtag=host"
			End If
		ElseIf ScriptName="list.asp" Then
			Name="BoardInfo_" & Boardid
			Board_Datas=Value
			If Not IsNumeric(Board_Data(24,0)) Then
				Board_Datas(24,0)=1
			Else
				Board_Datas(24,0)=Board_Datas(24,0)+1
			End If
			If Board_Datas(24,0) >= RayMaxCount Then
				RaySubjection=True
				Board_Datas(24,0)=1
				FrameBody="?PostType=1&forumchildname="&Boardtype
				FrameBody=FrameBody+"&forumchildurl="&Forum_url&"list.asp?boardid="&Boardid
				FrameBody=FrameBody+"&forumchildtitlecount="&Board_Datas(9,0)
				FrameBody=FrameBody+"&foruminlinecount="&MyBoardOnline.Forum_Online
				FrameBody=FrameBody+"&forumlogincount="&Dvbbs.CacheData(10,0)
				FrameBody=FrameBody+"&Forumvisitprob=1"
				FrameBody=FrameBody+"&forumchildtag=subjection"
			End If
			Value=Board_Datas
		End If
		If RaySubjection Then
			Response.Write "<iframe id=""RayCount"" src=""RayPost.asp"&FrameBody&""" width=0 height=0></iframe>"
		End If
		End If
		Response.Write Tmp
		'//------------------------------------------------------------------------------
	End Sub
	Public Function Dvbbs_Suc(sucmsg)
		Dim TempStr
		TempStr = mainhtml(13)
		TempStr = Replace(TempStr,"{$sucmsg}",sucmsg)
		TempStr = Replace(TempStr,"{$returnurl}",Request.ServerVariables("HTTP_REFERER"))
		Response.Write TempStr
		TempStr = ""
	End Function
	Public Function Execute(Command)
		If Not IsObject(Conn) Then ConnectionDatabase
		'检查权限,防止注入攻击。
		If InStr(LCase(Command),"dv_admin")>0 And Left(ScriptName,6)<> "admin_" Then 
			If savelog=1 Then
				Response.Write SaveSQLLOG(Command,"")
			End If
			Command=Replace(LCase(Command),"dv_admin","dv<i>"&Chr(95)&"</i>admin") 
		End If				
		If IsDeBug = 0 Then 
			On Error Resume Next
			Set Execute = Conn.Execute(Command)
			If Err Then
				err.Clear
				Set Conn = Nothing
				If savelog=1 Then
					Response.Write SaveSQLLOG(Command,"查询数据的时候发现错误，请检查您的查询代码是否正确。<br>基于安全的理由，只显示本信息，要查看详细的错误信息，请修改您的程序文件conn.asp。把""Const IsDeBug = 0""改为：""Const IsDeBug = 1""")
				Else
					Response.Write "查询数据的时候发现错误，请检查您的查询代码是否正确。"
				End If
				Response.End
			End If
		Else
			'Response.Write command & "<br>"
			Set Execute = Conn.Execute(Command)
		End If	
		SqlQueryNum = SqlQueryNum+1
	End Function
	'记录查询错误事件
	Public Function SaveSQLLOG(sCommand,message)
		Dim lConnStr,lConn,ldb,SQL,RS
		ldb = "data/DvSQLLOG.mdb"
		lConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ldb)
		Set lConn = Server.CreateObject("ADODB.Connection")
		lConn.Open lConnStr
		Set Rs = Server.CreateObject("adodb.recordset")
		Sql="select * from dv_sql_log"
		Rs.open sql,lconn,1,3
		Rs.addnew
		Rs("ScriptName")=ScriptName
		Rs("S_Info")=Left(sCommand,255)
		Rs("ip")=UserTrueIP
		Rs.update
		Rs.close
		lConn.Execute(SQL)
		lConn.Close
		Set lConn = Nothing 
		SaveSQLLOG = message
	End Function
	Public Sub ChecKIPlock()
		Dim IPlock
		IPlock = False
		Dim locklist
		locklist=Trim(CacheData(25,0))
		If locklist="" Then Exit Sub
		Dim i,StrUserIP,StrKillIP
		StrUserIP=UserTrueIP
		locklist=Split(locklist,"|")
		If StrUserIP="" Then Exit Sub
		StrUserIP=Split(UserTrueIP,".")
		If Ubound(StrUserIP)<>3 Then Exit Sub
		For i= 0 to UBound(locklist)
			locklist(i)=Trim(locklist(i))
			If locklist(i)<>"" Then 
				StrKillIP = Split(locklist(i),".")
				If Ubound(StrKillIP)<>3 Then Exit For
				IPlock = True
				If (StrUserIP(0) <> StrKillIP(0)) And Instr(StrKillIP(0),"*")=0 Then IPlock=False
				If (StrUserIP(1) <> StrKillIP(1)) And Instr(StrKillIP(1),"*")=0 Then IPlock=False
				If (StrUserIP(2) <> StrKillIP(2)) And Instr(StrKillIP(2),"*")=0 Then IPlock=False
				If (StrUserIP(3) <> StrKillIP(3)) And Instr(StrKillIP(3),"*")=0 Then IPlock=False
				If IPlock Then Exit For
			End If
		Next
		Response.Cookies(Forum_sn & "Kill").Expires = DateAdd("s", 360, Now())
		Response.Cookies(Forum_sn & "Kill").Path = Cookiepath
		If IPlock Then
			Response.Cookies(Forum_sn & "Kill")("kill") = "1"
		Else
			Response.Cookies(Forum_sn & "Kill")("kill") = "0"
		End If
	End Sub
	'IP/来源
	Public Function address(sip)
		Dim aConnStr,aConn,adb
		Dim str1,str2,str3,str4
		Dim  num
		Dim country,city
		Dim irs,SQL
		If IsNumeric(Left(sip,2)) Then
			If sip="127.0.0.1" Then sip="192.168.0.1"
			str1=Left(sip,InStr(sip,".")-1)
			sip=mid(sip,instr(sip,".")+1)
			str2=Left(sip,instr(sip,".")-1)
			sip=Mid(sip,InStr(sip,".")+1)
			str3=Left(sip,instr(sip,".")-1)
			str4=Mid(sip,instr(sip,".")+1)
			If isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 Then
			Else		
				num=CLng(str1)*16777216+CLng(str2)*65536+CLng(str3)*256+CLng(str4)-1
				adb = "data/ipaddress.mdb"
				aConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(adb)
				Set AConn = Server.CreateObject("ADODB.Connection")
				aConn.Open aConnStr
	
				sql="select top 1 country,city from dv_address where ip1 <="&num&" and ip2 >="&num&""
				Set irs=aConn.execute(sql)
				If irs.EOF And irs.bof Then
					country="亚洲"
					city=""
				Else
					country=irs(0)
					city=irs(1)
				End If
				Set irs=Nothing
				Set aConn = Nothing 
				SqlQueryNum = SqlQueryNum+1
			End If
			address=country&city
		Else 
			address="未知"
		End If
	End Function
	'显示验证码
	Public Function GetCode()
		Dim test
		On Error Resume Next
		Set test=Server.CreateObject("Adodb.Stream")
		Set test=Nothing
		If Err Then
			Dim zNum
			Randomize timer
			zNum = cint(8999*Rnd+1000)
			Session("GetCode") = zNum
			GetCode=Dvbbs.mainhtml(15)& Session("GetCode")		
		Else
			GetCode= Dvbbs.mainhtml(15)&"<img src=""DV_getcode.asp"">"		
		End If
	End Function
	'检查验证码是否正确
	Public Function CodeIsTrue()
		Dim CodeStr
		CodeStr=Trim(Request("CodeStr"))
		If CStr(Session("GetCode"))=CStr(CodeStr) And CodeStr<>""  Then
			CodeIsTrue=True
			Session("GetCode")=empty
		Else
			CodeIsTrue=False
			Session("GetCode")=empty
		End If	
	End Function
	'用于用户发布的各种信息过滤，带脏话过滤
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) Then
			fString = replace(fString, ">", "&gt;")
			fString = replace(fString, "<", "&lt;")
			fString = Replace(fString, CHR(32), " ")		'&nbsp;
			fString = Replace(fString, CHR(9), " ")			'&nbsp;
			fString = Replace(fString, CHR(34), "&quot;")
			'fString = Replace(fString, CHR(39), "&#39;")	'单引号过滤
			fString = Replace(fString, CHR(13), "")
			fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
			fString = Replace(fString, CHR(10), "<BR> ")
			fString=ChkBadWords(fString)
			HTMLEncode = fString
		End If
	End Function
	'用于论坛本身的过滤，不带脏话过滤
	Public Function iHTMLEncode(fString)
		If Not IsNull(fString) Then
			fString = replace(fString, ">", "&gt;")
			fString = replace(fString, "<", "&lt;")
			fString = Replace(fString, CHR(32), " ")
			fString = Replace(fString, CHR(9), " ")
			fString = Replace(fString, CHR(34), "&quot;")
			'fString = Replace(fString, CHR(39), "&#39;")
			fString = Replace(fString, CHR(13), "")
			fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
			fString = Replace(fString, CHR(10), "<BR> ")
			iHTMLEncode = fString
		End If
	End Function
	Public Function strLength(str)
		If isNull(str) Or Str = "" Then
			StrLength = 0
			Exit Function
		End If
		Dim WINNT_CHINESE
		WINNT_CHINESE=(len("例子")=2)
		If WINNT_CHINESE Then
			Dim l,t,c
			Dim i
			l=len(str)
			t=l
			For i=1 To l
				c=asc(mid(str,i,1))
				If c<0 Then c=c+65536
				If c>255 Then t=t+1
			Next
			strLength=t
		Else 
			strLength=len(str)
		End If
	End Function
	Public Function ChkBadWords(Str)
		If IsNull(Str) Then Exit Function
		Dim i
		For i = 0 To Ubound(BadWords)
			If i > UBound(rBadWord) Then
				Str = Replace(Str,BadWords(i),"*")
			Else
				Str = Replace(Str,BadWords(i),rBadWord(i))
			End If
		Next
		ChkBadWords = Str
	End Function
	Public Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
	End Function

	Public Function Get_Chan_Ad()
		Dim TempData,i
		Dim rndnum
		Dim Temp_Ad,Forum_AdLoop1,Forum_AdLoop2
		Temp_Ad = Split(CacheData(22,0),"||")
		If Temp_Ad(0)<>"" Then
			Forum_AdLoop1=Split(Temp_Ad(0),",")
		Else
			Forum_AdLoop1=Split("",",")
		End If
		If Temp_Ad(1)<>"" Then
			Forum_AdLoop2=Split(Temp_Ad(1),",")
		Else
			Forum_AdLoop2=Split("",",")
		End If
		Forum_AdLoop3 = Temp_Ad(2)
		'顶部banner
		Randomize
		rndnum=Cint(Ubound(Forum_AdLoop1)*rnd+1)
		If UBound(Forum_AdLoop1)=-1 Then
			adcode_1=""
		Else 
			Name = "ForumAdCode1"
			If ObjIsEmpty() Then LoadForumAdCode1
			If IsArray(Value) And Forum_ChanSetting(3)="1" Then
				TempData=Value
				adcode_1=ReCssUrl(TempData(1,rndnum-1))
			Else
				adcode_1=""
			End If
		End If
		'尾部通栏
		Randomize
		rndnum=Cint(Ubound(Forum_AdLoop2)*rnd+1)
		If UBound(Forum_AdLoop2)=-1 Then
			adcode_2=""
		Else
			Name = "ForumAdCode2"
			If ObjIsEmpty() Then LoadForumAdCode2
			If IsArray(Value) And Forum_ChanSetting(4)="1" Then
				TempData=Value
				adcode_2=ReCssUrl(TempData(1,rndnum-1))
			Else
				adcode_2=""
			End If
		End If
		Name = "ForumAdCode3"
		If ObjIsEmpty() Then LoadForumAdCode3
		If IsArray(Value) And Forum_ChanSetting(2)="1" Then
			TempData=Value
			adcode_4=ReCssUrl(TempData(1,i))
		Else
			adcode_4=""
		End If
		i3 = 0
		If Forum_AdLoop3<>"" And Forum_ChanSetting(5)="1" And Instr(ScriptName,"dispbbs")>0 Then
			Name = "TopicAdCode"
			If ObjIsEmpty() Then LoadTopicAdCode
			If IsArray(Value) Then
				TempData = Value
				For i=0 To Ubound(TempData,2)
					If TempData(1,i)=239 Or TempData(1,i)=240 Or TempData(1,i)=1 Or TempData(1,i)=2 Then
						ad_3(i3)=" "
					Else
						ad_3(i3)=ReCssUrl(TempData(0,i))
					End If
					i3 = i3 + 1
				Next
			End If
		End If
		If i3=0 Then Ad_3(0)=" "
	End Function
	Private Function LoadTopicAdCode()
		Dim Rs
		Set Rs=Execute("Select a_adcode,a_id From Dv_AdCode Where a_id In ("&Forum_AdLoop3&")")
		If Not Rs.Eof Then
			Value = Rs.GetRows(-1)
		Else
			Value = ""
		End If
		Set Rs=Nothing
	End Function
	Private Function LoadForumAdCode1()
		Dim Rs
		Set Rs=Execute("Select a_address,a_adcode,a_id From Dv_AdCode Where a_address='0001'")
		If Not Rs.Eof Then
			Value = Rs.GetRows(-1)
		Else
			Value = ""
		End If
		Set Rs=Nothing
	End Function
	Private Function LoadForumAdCode2()
		Dim Rs
		Set Rs=Execute("Select a_address,a_adcode,a_id From Dv_AdCode Where a_address='0002'")
		If Not Rs.Eof Then
			Value = Rs.GetRows(-1)
		Else
			Value = ""
		End If
		Set Rs=Nothing
	End Function
	Private Function LoadForumAdCode3()
		Dim Rs
		Set Rs=Execute("Select a_address,a_adcode,a_id From Dv_AdCode Where a_address='0004'")
		If Not Rs.Eof Then
			Value = Rs.GetRows(-1)
		Else
			Value = ""
		End If
		Set Rs=Nothing
	End Function
	Public Function ReCssUrl(str)
		if str="" then exit function
		str=replace(str,"%css%","Get_Css.asp?SkinID="&SkinID)
		str=replace(str,"%url%",Forum_info(1))
		If CacheData(23,0)="" or isnull(CacheData(23,0)) Then
		str=replace(str,"%username%","dvbbs")
		str=replace(str,"%mouseId%","dvbbs")
		Else
		str=replace(str,"%username%",CacheData(23,0))
		str=replace(str,"%mouseId%",CacheData(23,0))
		End If
		ReCssUrl=str
	End Function
	Public Function ReloadBoardInfo(lBoardID)
		If lBoardID=0 Then Exit Function
		'数组(21)TempStr用来记录版面的下拉菜单,(22)TempStr1用来保存该版面的导航,(23)TempStr2用来保存该版面的新闻和小字报,(24)TempStr3版块点击统计
		Dim Rs
		Set Rs=Execute("select BoardID,BoardType,ParentID,ParentStr,Depth,RootID,Child,readme,BoardMaster,PostNum,TopicNum,indexIMG,todayNum,boarduser,LastPost,Sid,Board_Setting,Board_Ads,Board_user,IsGroupSetting,BoardTopStr,BoardID As TempStr,BoardID As TempStr1,BoardID As TempStr2,BoardID As TempStr3,cid,BoardID As TempStr4 from Dv_board where BoardID="&lBoardID)
  			If Not Rs.Eof Then
  				Name = "BoardInfo_" & lBoardID
   				Value = Rs.GetRows(1)
  			Else
  				'自动修正所有版面的boards数
  				Call ReloadAllBoardInfo()
  				'Response.Redirect "index.asp"
  			End If
  		Rs.Close
  		Set Rs = Nothing
 	End Function
	'缓存版面公告和小字报信息
	Public Function LoadBoardNews_Paper(lBoardID)
		Dim tRs,bgs,MyGetData,TempStr,NoAnn,NoColor
		If Not IsArray(lanstr) Then
			NoAnn = "当前没有公告"
		Else
			NoAnn = lanstr(9)
		End If
		If Not IsArray(mainsetting) Then
			NoColor = "blue"
		Else
			NoColor = mainsetting(10)
		End If
		Set tRs=Execute("Select Top 1 title,addtime,bgs From [Dv_bbsnews] Where boardid="&lBoardID&" Order By ID Desc")
		If tRs.BOF And tRs.EOF Then
			TempStr = NoAnn & "|||"
		Else
			bgs=tRs(2)
			If bgs="" or IsNull(bgs) Then
				TempStr=tRs(0) & "|||" & tRs(1)
			Else
				TempStr="<img src=Skins/Default/filetype/mid.gif border=0><bgsound src="&bgs&" border=0>"&tRs(0)&"|||"&tRs(1)
			End if
		End If
		'小字报部分
		If IsSqlDataBase=1 Then
			Set tRs=Execute("Select Top 5 S_id as id,S_username as postuser,S_title as topic From Dv_Smallpaper Where Datediff(D,S_addtime,"&SqlNowString&")<=1 And S_boardid="&lBoardID&" Order By S_addtime Desc")
		Else
			Set tRs=Execute("Select Top 5 S_id as id,S_username as postuser,S_title as topic From Dv_Smallpaper Where Datediff('D',S_addtime,"&SqlNowString&")<=1 And S_boardid="&lBoardID&" Order By S_addtime Desc")
		End If
		If tRs.Eof And tRs.Bof Then
			TempStr=TempStr & "|||"
		Else
			Dim TempData,i
			TempData=tRs.GetRows(-1)
			For i=0 To Ubound(TempData,2)
				If i=0 Then
					TempStr = TempStr & "|||&nbsp;&nbsp;<font color="&NoColor&">"&HtmlEncode(TempData(1,i))&"</font>：<a href=javascript:openScript(""viewpaper.asp?id="&TempData(0,i)&"&boardid="&BoardID&""",500,400)>"&HtmlEncode(TempData(2,i))&"</a>&nbsp;&nbsp;"
				Else
					TempStr = TempStr & "&nbsp;&nbsp;<font color="&NoColor&">"&HtmlEncode(TempData(1,i))&"</font>：<a href=javascript:openScript(""viewpaper.asp?id="&TempData(0,i)&"&boardid="&BoardID&""",500,400)>"&HtmlEncode(TempData(2,i))&"</a>&nbsp;&nbsp;"
				End If
			Next
		End If
		MyGetData = Value
		MyGetData(23,0) = TempStr
		Value = MyGetData
		Set tRs=Nothing
	End Function
	'缓存导航相关信息
	Public Sub LoadBoardParentStr(MyParentStr)
		Dim tRs,GetData,MyGetData
		Set tRs=Execute("Select Boardid,Boardtype,Boardmaster,Parentid From Dv_Board Where Boardid In ("&MyParentStr&") Order By Orders")
		If Not tRs.Eof Then
			GetData = tRs.GetRows(-1)
			MyGetData = Value
			MyGetData(22,0) = GetData
			value=MyGetData
		End If
		Set tRs = Nothing
	End Sub
	'对应Dvbbs.Board_Data(21,0)，Act=1.导航菜单缓存;Dvbbs.Board_Data(26,0)，Act=0不含隐藏论坛的导航菜单缓存;
	Public Sub LoadBoardList(lBoardID,Act)
		Dim Forum_Boards,i,ii,Depth,Board_Datas,MyBoardList,MyBoardRootID,MyBoard_Data,b_setting
		If lBoardID=0 Then Exit Sub
		Name="BoardInfo_" & lBoardID
		MyBoard_Data=value
		MyBoardRootID=Clng(MyBoard_Data(5,0))
		Forum_Boards=Split(CacheData(27,0),",")
		For i=0 To Ubound(Forum_Boards)
			Name="BoardInfo_" & Forum_Boards(i)
			If ObjIsEmpty() Then ReloadBoardInfo(Forum_Boards(i))
			Board_Datas = Value
			b_setting=split(Board_Datas(16,0),",")
			If b_setting(1)<>"1" Or Act=1 Then
				Depth=Board_Datas(4,0)
				If MyBoardRootID = Board_Datas(5,0) And (Not Board_Datas(2,0) = 0) Then MyBoardList = MyBoardList & "<a href=list.asp?boardid="&Forum_Boards(i)&">"
				Select Case Depth
				Case 0
					If MyBoardRootID = Board_Datas(5,0) And (Not Board_Datas(2,0) = 0) Then MyBoardList = MyBoardList & "╋"
				Case 1
					If MyBoardRootID = Board_Datas(5,0) And (Not Board_Datas(2,0) = 0) Then MyBoardList = MyBoardList & "&nbsp;&nbsp;├"
				End Select
				If Depth>1 Then
					For ii=2 To Depth
						If MyBoardRootID = Board_Datas(5,0) And (Not Board_Datas(2,0) = 0) Then MyBoardList = MyBoardList & "&nbsp;&nbsp;│"
					Next
					If MyBoardRootID = Board_Datas(5,0) And (Not Board_Datas(2,0) = 0) Then MyBoardList = MyBoardList & "&nbsp;&nbsp;├"
				End If
				If MyBoardRootID = Board_Datas(5,0) And (Not Board_Datas(2,0) = 0) Then MyBoardList = MyBoardList & Server.htmlencode(Board_Datas(1,0)) & "</a><br>"
			End If
		Next
		Name="BoardInfo_" & lBoardID
		MyBoard_Data=value
		If Act=1 Then
			MyBoard_Data(21,0)=Replace(Replace(MyBoardList,"'","\'"),Chr(34), "&quot;")
			Board_Data(21,0)=MyBoard_Data(21,0)
		Else
			MyBoard_Data(26,0)=Replace(Replace(MyBoardList,"'","\'"),Chr(34), "&quot;")
			Board_Data(26,0)=MyBoard_Data(26,0)
		End If
		value=MyBoard_Data
		Forum_Boards=Null
		Board_Datas=Null
	End Sub
	Public Sub ReloadAllBoardInfo()
		Dim Rs,Boards
		Set Rs=Execute("Select BoardID From Dv_Board Order By RootID,Orders")
		If Not Rs.Eof Then
			Boards=Rs.GetString(,-1, "",",","")
			Boards=Left(Boards,Len(Boards)-1)
		End If
		Rs.close:Set Rs=Nothing
		Execute("Update dv_Setup Set Forum_Boards='"&Boards&"'")
		ReloadSetupCache Boards,27
	End Sub 
	'更新分版面部分缓存数组，入口：版面ID、更新内容、数组位置、更新方式，0直接赋值，1数值相加
	Public Sub ReloadBoardCache(lBoardID,MyValue,N,act)
		If lBoardID=0 Then Exit Sub
		If lBoardID=444 Or lBoardID=777 Or lBoardID="" Then
  			Response.Write "错误的版面参数"
  			Response.End
		End If
		Dim Tmpdata
		Name="BoardInfo_" & lBoardID
		If ObjIsEmpty() Then ReloadBoardInfo(lBoardID)
		Tmpdata=Value
		If act=1 And IsNumeric(Tmpdata(N,0)) And IsNumeric(MyValue) Then
			Tmpdata(N,0)=CLng(Tmpdata(N,0))+MyValue
		ElseIf act=2 And IsNumeric(Tmpdata(N,0)) And IsNumeric(MyValue) Then
			Tmpdata(N,0)=CLng(Tmpdata(N,0))-MyValue
		Else
			Tmpdata(N,0) = MyValue
   		End If
   		Value=Tmpdata
	End Sub
	Public Function ReloadForumPlusMenu()
		Dim Rs,tRs,TempMenu,TempMenu1,MSetting
		Name="ForumPlusMenu"&SkinID
		Set Rs=Dvbbs.Execute("Select * From Dv_Plus Where Plus_Type='0' and Isuse=1 Order By ID")
		If Rs.Eof And Rs.Bof Then
			Value=""
			Exit Function
		End If
		Do While Not Rs.Eof
			MSetting=Split(Split(Rs("Plus_Setting"),"|||")(0),"|")
			Set tRs=Dvbbs.Execute("Select * From Dv_Plus Where Plus_Type='"&Rs("ID")&"' and Isuse=1 Order By ID")
			If tRs.Eof And tRs.Bof Then
				Select Case MSetting(0)
				Case 0
					TempMenu = TempMenu & " <img src="&mainpic(18)&" align=absmiddle> <a href="""&Rs("MainPage")&""" title="""&Rs("Plus_CopyRight")&""">"&Rs("Plus_Name")&"</a>"
				Case 1
					TempMenu = TempMenu & " <img src="&mainpic(18)&" align=absmiddle> <a href="""&Rs("MainPage")&""" title="""&Rs("Plus_CopyRight")&""" target=_blank>"&Rs("Plus_Name")&"</a>"
				Case 2
					TempMenu = TempMenu & " <img src="&mainpic(18)&" align=absmiddle> <a href=""JavaScript:openScript('"&Rs("MainPage")&"',"&MSetting(1)&","&MSetting(2)&")"" title="""&Rs("Plus_CopyRight")&""">"&Rs("Plus_Name")&"</a>"
				Case 3
					TempMenu = TempMenu & " <img src="&mainpic(18)&" align=absmiddle> <a href=""JavaScript:openScript('"&Rs("MainPage")&"',screen.width,screen.height)"" title="""&Rs("Plus_CopyRight")&""">"&Rs("Plus_Name")&"</a>"
				End Select
			Else
				Do While Not tRs.Eof
					MSetting=Split(Split(tRs("Plus_Setting"),"|||")(0),"|")
					Select Case MSetting(0)
					Case 0
						TempMenu1 = TempMenu1 & "<div class=menuitems><a href="&tRs("MainPage")&" title="&tRs("Plus_CopyRight")&">"&tRs("Plus_Name")&"</a></div>"
					Case 1
						TempMenu1 = TempMenu1 & "<div class=menuitems><a href="&tRs("MainPage")&" title="&tRs("Plus_CopyRight")&" target=_blank>"&tRs("Plus_Name")&"</a></div>"
					Case 2
						TempMenu1 = TempMenu1 & "<div class=menuitems><a href=JavaScript:openScript(\'"&tRs("MainPage")&"\',"&MSetting(1)&","&MSetting(2)&") title="&tRs("Plus_CopyRight")&">"&tRs("Plus_Name")&"</a></div>"
					Case 3
						TempMenu1 = TempMenu1 & "<div class=menuitems><a href=JavaScript:openScript(\'"&tRs("MainPage")&"\',screen.width,screen.height) title="&tRs("Plus_CopyRight")&">"&tRs("Plus_Name")&"</a></div>"
					End Select
				tRs.MoveNext
				Loop
				MSetting=Split(Split(Rs("Plus_Setting"),"|||")(0),"|")
				Select Case MSetting(0)
				Case 0
					TempMenu = TempMenu & " <img src="&mainpic(18)&" align=absmiddle> <a href="""&Rs("MainPage")&""" title="""&Rs("Plus_CopyRight")&""" onMouseOver=""showmenu(event,'"&TempMenu1&"')"">"&Rs("Plus_Name")&"</a>"
				Case 1
					TempMenu = TempMenu & " <img src="&mainpic(18)&" align=absmiddle> <a href="""&Rs("MainPage")&""" title="""&Rs("Plus_CopyRight")&""" target=_blank onMouseOver=""showmenu(event,'"&TempMenu1&"')"">"&Rs("Plus_Name")&"</a>"
				Case 2
					TempMenu = TempMenu & " <img src="&mainpic(18)&" align=absmiddle> <a href=""JavaScript:openScript('"&Rs("MainPage")&"',"&MSetting(1)&","&MSetting(2)&")"" title="""&Rs("Plus_CopyRight")&""" onMouseOver=""showmenu(event,'"&TempMenu1&"')"">"&Rs("Plus_Name")&"</a>"
				Case 3
					TempMenu = TempMenu & " <img src="&mainpic(18)&" align=absmiddle> <a href=""JavaScript:openScript('"&Rs("MainPage")&"',screen.width,screen.height)"" title="""&Rs("Plus_CopyRight")&""" onMouseOver=""showmenu(event,'"&TempMenu1&"')"">"&Rs("Plus_Name")&"</a>"
				End Select
				TempMenu1=""
			End If
			Rs.MoveNext
		Loop
		Value=TempMenu
		Set tRs=Nothing
		Set Rs=Nothing
	End Function
	'取得带端口的URL
	Property Get Get_ScriptNameUrl()
		If request.servervariables("SERVER_PORT")="80" Then
			Get_ScriptNameUrl="http://" & request.servervariables("server_name")&replace(lcase(request.servervariables("script_name")),ScriptName,"")
		Else
			Get_ScriptNameUrl="http://" & request.servervariables("server_name")&":"&request.servervariables("SERVER_PORT")&replace(lcase(request.servervariables("script_name")),ScriptName,"")
		End If
	End Property
End Class
Class cls_Templates
	Public html,Strings,pic
	Public Property Let Value(ByVal vNewValue)
		Dim tmpstr:tmpstr = vNewValue
		tmpstr = Replace(tmpstr,"{$PicUrl}",Dvbbs.Forum_PicUrl)
		tmpstr = Split(tmpstr,"@@@")
		html = Split(tmpstr(0),"|||"):Strings = Split(tmpstr(1),"|||"):pic = Split(tmpstr(2),"|||")
	End Property
End Class
Class cls_UserOnlne
	Public Forum_Online,Forum_UserOnline,Forum_GuestOnline
	Private l_Online,l_GuestOnline
	Private Sub Class_Initialize()
		Dvbbs.Name="Forum_Online"
		Dvbbs.Reloadtime=60
		If Dvbbs.ObjIsEmpty() Then ReflashOnlineNum
		Dvbbs.Name="Forum_Online"
		Forum_Online = Dvbbs.Value
		Dvbbs.Name="Forum_UserOnline"
		If Dvbbs.ObjIsEmpty() Then ReflashOnlineNum
		Forum_UserOnline=Dvbbs.Value
		If Forum_Online < 0  Or Forum_UserOnline < 0 Or Forum_UserOnline > Forum_Online Then ReflashOnlineNum
		Forum_GuestOnline = Forum_Online - Forum_UserOnline
		l_Online=-1:l_GuestOnline=-1
		Dvbbs.Reloadtime=14400
	End Sub
	Public Sub OnlineQuery()
		Dim SQL,SQL1
		Dim TempNum,TempNum1
		Dvbbs.Name="delOnline_time"
		If Dvbbs.ObjIsEmpty() Then Dvbbs.Value=Now()
		If DateDiff("s",Dvbbs.Value,Now()) > Clng(Dvbbs.Forum_Setting(8))*10 Then
			Dvbbs.Value=Now()
			If Not IsObject(Conn) Then ConnectionDatabase
			If IsSqlDataBase = 1 Then
				SQL = "Delete From [DV_Online] Where UserID=0 And Datediff(Mi, Lastimebk, " & SqlNowString & ") > " & Clng(Dvbbs.Forum_Setting(8))
				SQL1 = "Delete From [DV_Online] Where UserID>0 And Datediff(Mi, Lastimebk, " & SqlNowString & ") > " & Clng(Dvbbs.Forum_Setting(8))
			Else
				SQL = "Delete From [Dv_Online] Where UserID=0 And Datediff('s', Lastimebk, " & SqlNowString & ") > " & Dvbbs.Forum_Setting(8) & "*60" 
				SQL1 = "Delete From [Dv_Online] Where UserID>0 And Datediff('s', Lastimebk, " & SqlNowString & ") > " & Dvbbs.Forum_Setting(8) & "*60"
			End If
			Conn.Execute SQL,TempNum
			Conn.Execute SQL1,TempNum1
			Dvbbs.SqlQueryNum = Dvbbs.SqlQueryNum + 2
			'如果删除客人数大于0，则应该更新总数
			If TempNum>0 Then
				'更新缓存总在线数据
				Forum_Online = Forum_Online - TempNum
				Forum_GuestOnline = Forum_GuestOnline - TempNum
			End If
			'如果删除用户数大于0，则应该更新总数和用户数
			If TempNum1>0 Or  TempNum>0 Then
				'更新缓存总在线数据
				Forum_Online = Forum_Online - TempNum1
				Forum_UserOnline = Forum_UserOnline - TempNum1
				
			End If
			Dvbbs.Name="Forum_Online"
			Dvbbs.Value=Forum_Online
			'更新缓存总用户在线数据
			Dvbbs.Name="Forum_UserOnline"
			Dvbbs.Value=Forum_UserOnline
			Forum_Online = Forum_Online - TempNum1
		End If
	End Sub
	'刷新在线数据缓存
	Public Sub ReflashOnlineNum
		Dim Rs
		Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Online")
		Dvbbs.Value=Rs(0)
		Forum_Online = Dvbbs.Value
		Dvbbs.Name="Forum_UserOnline"
		Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Online Where UserID>0")
		Dvbbs.Value=Rs(0)
		Forum_UserOnline = Dvbbs.Value
		Set Rs=Nothing
	End Sub
	'查询在某版面的在线总数
	Public Property Get Board_Online
		Board_Online=Board_UserOnline+Board_GuestOnline
	End Property
	Public Property Get Board_GuestOnline
		If l_GuestOnline=-1 Then
			Dim Rs
			Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Online where BoardID="&Dvbbs.BoardID&" and UserID=0")
			l_GuestOnline=Rs(0):Set Rs= Nothing 
		End If
		Board_GuestOnline=l_GuestOnline
	End Property
	Public Property Get Board_UserOnline
		If l_Online=-1 Then
			Dim Rs
			Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Online where BoardID="&Dvbbs.BoardID&" and UserID>0")
			l_Online=Rs(0):Set Rs= Nothing 
		End If
		Board_UserOnline=l_Online
	End Property
End Class
'Session(Dvbbs.CacheName & "Cls_Browser") 0:Browser+|||+1:version+|||+2:platform
Class Cls_Browser
	Public Browser,version ,platform,IsSearch
	Private Sub Class_Initialize()
		Dim Agent,Tmpstr
		IsSearch = False
		If Not IsEmpty(Session(Dvbbs.CacheName & "Cls_Browser")) Then
			Tmpstr = Split(Session(Dvbbs.CacheName & "Cls_Browser"),"|||")
			Browser = Dvbbs.checkStr(Tmpstr(0))
			version = Dvbbs.checkStr(Tmpstr(1))
			platform = Dvbbs.checkStr(Tmpstr(2))
			If Tmpstr(3)="1" Then 
				IsSearch = True
			End If
			Exit Sub
		End If
		Browser="unknown"
		version="unknown"
		platform="unknown"
		Agent=Request.ServerVariables("HTTP_USER_AGENT")
		'Agent="Opera/7.23 (X11; Linux i686; U)  [en]"	
		If Left(Agent,7) ="Mozilla" Then '有此标识为浏览器
			Agent=Split(Agent,";")
			If InStr(Agent(1),"MSIE")>0 Then
				Browser="Microsoft Internet Explorer "
				version=Trim(Left(Replace(Agent(1),"MSIE",""),6))
			ElseIf InStr(Agent(4),"Netscape")>0 Then 
				Browser="Netscape "
				tmpstr=Split(Agent(4),"/")
				version=tmpstr(UBound(tmpstr))
			ElseIf InStr(Agent(4),"rv:")>0 Then
				Browser="Mozilla "
				tmpstr=Split(Agent(4),":")
				version=tmpstr(UBound(tmpstr))
				If InStr(version,")") > 0 Then 
					tmpstr=Split(version,")")
					version=tmpstr(0)
				End If
			End If
			If InStr(Agent(2),"NT 5.2")>0 Then
				platform="Windows 2003"
			ElseIf InStr(Agent(2),"Windows CE")>0 Then
				platform="Windows CE"
			ElseIf InStr(Agent(2),"NT 5.1")>0 Then
				platform="Windows XP"
			ElseIf InStr(Agent(2),"NT 4.0")>0 Then
				platform="Windows NT"
			ElseIf InStr(Agent(2),"NT 5.0")>0 Then
				platform="Windows 2000"
			ElseIf InStr(Agent(2),"NT")>0 Then
				platform="Windows NT"
			ElseIf InStr(Agent(2),"9x")>0 Then
				platform="Windows ME"
			ElseIf InStr(Agent(2),"98")>0 Then
				platform="Windows 98"
			ElseIf InStr(Agent(2),"95")>0 Then
				platform="Windows 95"
			ElseIf InStr(Agent(2),"Win32")>0 Then
				platform="Win32"
			ElseIf InStr(Agent(2),"Linux")>0 Then
				platform="Linux"
			ElseIf InStr(Agent(2),"SunOS")>0 Then
				platform="SunOS"
			ElseIf InStr(Agent(2),"Mac")>0 Then
				platform="Mac"
			ElseIf UBound(Agent)>2 Then
				If InStr(Agent(3),"NT 5.1")>0 Then
					platform="Windows XP"
				End If 
				If InStr(Agent(3),"Linux")>0 Then
					platform="Linux"
				End If
			End If
			If InStr(Agent(2),"Windows")>0 And platform="unknown" Then
				platform="Windows"
			End If
		ElseIf Left(Agent,5) ="Opera" Then '有此标识为浏览器
			Agent=Split(Agent,"/")
			Browser="Mozilla "
			tmpstr=Split(Agent(1)," ")
			version=tmpstr(0)
			If InStr(Agent(1),"NT 5.2")>0 Then
				platform="Windows 2003"
			ElseIf InStr(Agent(1),"Windows CE")>0 Then
				platform="Windows CE"
			ElseIf InStr(Agent(1),"NT 5.1")>0 Then
				platform="Windows XP"
			ElseIf InStr(Agent(1),"NT 4.0")>0 Then
				platform="Windows NT"
			ElseIf InStr(Agent(1),"NT 5.0")>0 Then
				platform="Windows 2000"
			ElseIf InStr(Agent(1),"NT")>0 Then
				platform="Windows NT"
			ElseIf InStr(Agent(1),"9x")>0 Then
				platform="Windows ME"
			ElseIf InStr(Agent(1),"98")>0 Then
				platform="Windows 98"
			ElseIf InStr(Agent(1),"95")>0 Then
				platform="Windows 95"
			ElseIf InStr(Agent(1),"Win32")>0 Then
				platform="Win32"
			ElseIf InStr(Agent(1),"Linux")>0 Then
				platform="Linux"
			ElseIf InStr(Agent(1),"SunOS")>0 Then
				platform="SunOS"
			ElseIf InStr(Agent(1),"Mac")>0 Then
				platform="Mac"
			ElseIf UBound(Agent)>2 Then
				If InStr(Agent(3),"NT 5.1")>0 Then
					platform="Windows XP"
				End If 
				If InStr(Agent(3),"Linux")>0 Then
					platform="Linux"
				End If
			End If
		Else
			'识别搜索引擎
			Dim botlist,i
			Botlist="Google,Isaac,SurveyBot,Baiduspider,ia_archiver,P.Arthur,FAST-WebCrawler,Java,Microsoft-ATL-Native,TurnitinBot,WebGather,Sleipnir"
			Botlist=split(Botlist,",")
			For i=0 to UBound(Botlist)
				If InStr(Agent,Botlist(i))>0  Then 
					platform=Botlist(i)&"搜索器"
					IsSearch=True
					Exit For
				End If
			Next 
		End If
		If version<>"unknown" Then 
			Dim Tmpstr1
			Tmpstr1=Trim(Replace(version,".",""))
			If Not IsNumeric(Tmpstr1) Then
				version="unknown"
			End If
		End If
		If IsSearch Then
			Browser=""
			version=""
			Session(Dvbbs.CacheName & "Cls_Browser") = Browser &"|||"& version &"|||"& platform&"|||1"
		Else
			Session(Dvbbs.CacheName & "Cls_Browser") = Browser &"|||"& version &"|||"& platform&"|||0"
		End If
	End Sub
End Class
%>