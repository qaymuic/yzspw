<!--#include file="Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/email.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.LoadTemplates("login")
Dim comeurl
dvbbs.stats=template.Strings(1)
Dvbbs.Nav()
dvbbs.Head_var 0,0,template.Strings(0),"login.asp"
Dim TruePassWord
TruePassWord=Dvbbs.Createpass
Select Case request("action")
Case "chk"
	Dvbbs_ChkLogin
	Dvbbs.Showerr()
Case "redir"
	redir
	Dvbbs.Showerr()
Case "save_redir_reg"
	call save_redir_reg()
	Dvbbs.Showerr()
Case Else
	Main
End Select

Dvbbs.ActiveOnline
Dvbbs.Footer()

Function Main()
	Dim TempStr
	TempStr = template.html(0)
	If Dvbbs.forum_setting(79)="0" Then
		TempStr = Replace(TempStr,"{$getcode}","")
	Else
		template.html(23)=Replace(template.html(23),"{$codestr}",Dvbbs.GetCode())
		TempStr = Replace(TempStr,"{$getcode}",template.html(23))
	End If
	If Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(10)=1 Then
		TempStr = Replace(TempStr,"{$rayuserlogin}",template.html(1))
	Else
		TempStr = Replace(TempStr,"{$rayuserlogin}","")
	End If
	Dim Comeurl,tmpstr
	If Request.ServerVariables("HTTP_REFERER")<>"" Then 
		tmpstr=split(Request.ServerVariables("HTTP_REFERER"),"/")
		Comeurl=tmpstr(UBound(tmpstr))
	Else
		Comeurl="index.asp"
	End If
	TempStr = Replace(TempStr,"{$comeurl}",Comeurl)
	Response.Write TempStr
	TempStr=""
End Function

Function Dvbbs_ChkLogin

	Dim UserIP
	Dim username
	Dim userclass
	Dim password
	Dim article
	Dim usercookies
	Dim mobile
	Dim chrs,i
	If Dvbbs.forum_setting(79)="1" Then
		If Not Dvbbs.CodeIsTrue() Then
			 Response.redirect "showerr.asp?ErrCodes=<li>验证码校验失败，请返回刷新页面后再输入验证码。&action=OtherErr"
		End If
	End If
	UserIP=Dvbbs.UserTrueIP
	mobile=trim(Dvbbs.CheckStr(request("mobile")))
	if mobile<>"" and request("username")="" then
		if len(mobile)<>11 then
			Dvbbs.AddErrCode(9)
		end if
	end if
	if mobile<>"" then
		if len(mobile)<>11 then mobile=""
	end if
	If request("username")="" Then
		If request("mobile")="" Then
			Dvbbs.AddErrCode(10)
		End If
	Else
		username=trim(Dvbbs.CheckStr(request("username")))
	End If
	If request("password")="" and mobile="" Then
		Dvbbs.AddErrCode(11)
	Else
		password=md5(trim(Dvbbs.CheckStr(request("password"))),16)
	End If
	If Dvbbs.ErrCodes<>"" Then Exit Function
	usercookies=request("CookieDate")
	'判断更新cookies目录
	Dim cookies_path_s,cookies_path_d,cookies_path
	cookies_path_s=split(Request.ServerVariables("PATH_INFO"),"/")
	cookies_path_d=ubound(cookies_path_s)
	cookies_path="/"
	For i=1 to cookies_path_d-1
		If not (cookies_path_s(i)="upload" or cookies_path_s(i)="admin") Then cookies_path=cookies_path&cookies_path_s(i)&"/"
	Next
	If dvbbs.cookiepath<>cookies_path Then
		cookies_path=replace(cookies_path,"'","")
		Dvbbs.execute("update dv_setup set Forum_Cookiespath='"&cookies_path&"'")
		Dim setupData 
		Dvbbs.CacheData(26,0)=cookies_path
		Dvbbs.Name="setup"
		Dvbbs.value=Dvbbs.CacheData
	End If
	If ChkUserLogin(username,password,mobile,usercookies,1)=false Then
		'本地验证未通过，使用手机号登录的
		If mobile<>"" Then
			challenge_check mobile,password
			Exit Function
		'本地验证未通过，使用用户名登录的，并且是高级用户则继续主服务器验证流程
		Else
			set chrs=Dvbbs.Execute("select UserMobile,IsChallenge from [Dv_User] where username='"&username&"' and IsChallenge=1")
			If chrs.eof and chrs.bof Then
				Dvbbs.AddErrCode(12)
				Exit Function
			Else
				challenge_check chrs("UserMobile"),password
				Exit Function
			End If
			set chrs=nothing
		End If
	End If

	Dim comeurlname
	If instr(lcase(request("comeurl")),"reg.asp")>0 or instr(lcase(request("comeurl")),"login.asp")>0 or trim(request("comeurl"))="" Then
		comeurlname=""
		comeurl="index.asp"
	Else
		comeurl=request("comeurl")
		comeurlname="<li><a href="&request("comeurl")&">"&request("comeurl")&"</a></li>"
	End If

	Dim TempStr
	TempStr = template.html(2)
	If Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(10)=1 And Dvbbs.Forum_ChanSetting(12)=1 Then
		TempStr = Replace(TempStr,"{$ray_logininfo}",template.html(3))
	Else
		TempStr = Replace(TempStr,"{$ray_logininfo}","")
	End If
	TempStr = Replace(TempStr,"{$comeurl}",comeurl)
	TempStr = Replace(TempStr,"{$comeurlinfo}",comeurlname)
	TempStr = Replace(TempStr,"{$forumname}",Dvbbs.Forum_Info(0))
	Response.Write TempStr
	TempStr=""

End Function

'全网认证
Function challenge_check(mobile,password)
	If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(10)=1) Then
		Dvbbs.AddErrCode(13)
		exit function
	End If
	Dim rs
	Dim MyForumID
	Dim PostChanWord
	set rs=Dvbbs.Execute("select top 1 * from Dv_ChallengeInfo")
	MyForumID=rs("D_ForumID")
	PostChanWord=Get_ChallengeWord

	Dim TempStr,TempArray
	TempArray = Split(template.html(19),"||")
	TempStr = TempArray(0)
	TempStr = Replace(TempStr,"{$mobile}",mobile)
	TempStr = Replace(TempStr,"{$password}",password)
	TempStr = Replace(TempStr,"{$MyForumID}",MyForumID)
	TempStr = Replace(TempStr,"{$serverurl}",Dvbbs.Get_ScriptNameUrl())
	TempStr = Replace(TempStr,"{$PostChanWord}",PostChanWord)
	TempStr = Replace(TempStr,"{$remobile}",left(mobile,3)&"xxx"&right(mobile,5))
	TempStr = Replace(TempStr,"{$usermobile}",left(mobile,3)&"xxx"&right(mobile,5))
	If PassWord<>"" Then
		TempStr = Replace(TempStr,"{$ifpassnull}",TempArray(1))
	Else
		TempStr = Replace(TempStr,"{$ifpassnull}","")
	End If
	Response.Write TempStr
	TempStr = ""
	set rs=nothing
End Function

Function redir()

	Dim ErrorCode,ErrorMsg
	Dim remobile,rechallengeWord,retokerWord,reuserpassword
	Dim resex,reqq,reemail,reusername
	Dim challengeWord_key,rechallengeWord_key
	Dim userclass
	Dim rs

	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))
	remobile=trim(Dvbbs.CheckStr(request("mobile")))
	reuserpassword=trim(Dvbbs.CheckStr(request("forumPwd")))
	rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
	retokerWord=trim(request("tokenWord"))
	resex=trim(Dvbbs.CheckStr(request("sex")))
	If resex="F" Then 
		resex=1
	Else
		resex=0
	End If
	reqq=trim(Dvbbs.CheckStr(request("qq")))
	reemail=trim(Dvbbs.CheckStr(request("email")))
	reusername=trim(Dvbbs.CheckStr(request("username")))

	Session("re_challenge_reg_temp")=checkreal(remobile) & "|||" & checkreal(reuserpassword) & "|||" & checkreal(resex) & "|||" & checkreal(reqq) & "|||" & checkreal(reemail) & "|||" & checkreal(reusername)

	select case ErrorCode
	case 100
		challengeWord_key=session("challengeWord_key")
		If challengeWord_key=retokerWord Then
			set rs=Dvbbs.Execute("select UserMobile,IsChallenge,userid,userclass,username from [Dv_User] where UserMobile='"&remobile&"' and IsChallenge=1")
			If rs.eof and rs.bof Then
				'不是本论坛高级用户，引导其注册
				Call redir_reg_1()
				Exit Function
			Else
				Dvbbs.Execute("update [Dv_User] set UserPassword='"&md5(reuserpassword,16)&"' where UserMobile='"&remobile&"' and IsChallenge=1")
				dvbbs.userid=rs(2)
				userclass=rs(3)
				reusername=rs(4)
			End If
		Else
			Dvbbs.AddErrCode(14)
			'challengeWord_key & "," & retokerWord & "," & md5(Session("challengeWord") & ":" & "raynetwork",32) & "<br>原始随机数："&Session("challengeWord")&",返回随机数："&rechallengeWord&""
			Exit Function
		End If
	case 101
		Dvbbs.AddErrCode(15)
		Exit Function
	case Else
		Dvbbs.AddErrCode(14)
		Exit Function
	end select

	Dim TempStr
	TempStr = template.html(20)
	If Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(10)=1 And Dvbbs.Forum_ChanSetting(12)=1 Then
		TempStr = Replace(TempStr,"{$ray_logininfo}",template.html(3))
	Else
		TempStr = Replace(TempStr,"{$ray_logininfo}","")
	End If
	TempStr = Replace(TempStr,"{$reuserpassword}",reuserpassword)
	TempStr = Replace(TempStr,"{$forumname}",Dvbbs.Forum_Info(0))
	Response.Write TempStr
	TempStr=""
	Dim StatUserID,UserSessionID
	StatUserID = Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("StatUserID")))
	If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
		StatUserID = Replace(Dvbbs.UserTrueIP,".","")
		UserSessionID = Replace(Startime,".","")
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
		StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
	End If
	StatUserID = Ccur(StatUserID)
	'客人=SessionID+活动时间+发贴时间+版面ID
	Session(Dvbbs.CacheName & "UserID") = Split(StatUserID & "_" & Now & "_" & Now & "_" & Dvbbs.BoardID,"_")
	Response.Cookies(Dvbbs.Forum_sn).Expires=DateAdd("s",3600,Now())
	Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
	Response.Cookies(Dvbbs.Forum_sn)("StatUserID") = StatUserID
	Response.Cookies(Dvbbs.Forum_sn)("usercookies") = "0"
	Response.Cookies(Dvbbs.Forum_sn)("username") = reusername
	Response.Cookies(Dvbbs.Forum_sn)("userid") = dvbbs.UserID
	Response.Cookies(Dvbbs.Forum_sn)("password") = TruePassWord
	Response.Cookies(Dvbbs.Forum_sn)("userclass") = userclass
	Response.Cookies(Dvbbs.Forum_sn)("userhidden") = 2
	rem 清除图片上传数的限制
	response.cookies("upNum")=0
	Response.Cookies(Dvbbs.Forum_sn).path=dvbbs.cookiepath
	
End Function

sub redir_reg_1()

	If Session("re_challenge_reg_temp")="" Then
		Dvbbs.AddErrCode(14)
		exit sub
	End If

	Dim re_challenge_reg_temp
	re_challenge_reg_temp=split(Session("re_challenge_reg_temp"),"|||")

	Dim TempStr
	TempStr = template.html(21)
	TempStr = Replace(TempStr,"{$maxuserlength}",Dvbbs.Forum_Setting(41))
	TempStr = Replace(TempStr,"{$minuserlength}",Dvbbs.Forum_Setting(40))
	TempStr = Replace(TempStr,"{$reusername}",re_challenge_reg_temp(5))
	TempStr = Replace(TempStr,"{$width}",Dvbbs.mainsetting(0))
	Response.Write TempStr
end sub

sub save_redir_reg()
	If Session("re_challenge_reg_temp")="" Then
		Dvbbs.AddErrCode(14)
		exit sub
	End If

	Dim username,sex,pass1,pass2,password
	Dim useremail,face,width,height
	Dim oicq,sign,showRe,birthday
	Dim mailbody,sendmsg,rndnum,num1
	Dim quesion,answer,topic
	Dim userinfo,usersetting
	Dim userclass,UserIM
	Dim re_challenge_reg_temp
	Dim rs,sql,i,namebadword,SplitWords
	re_challenge_reg_temp=split(Session("re_challenge_reg_temp"),"|||")

	If request("name")="" or Dvbbs.strLength(request("name"))>Cint(Dvbbs.Forum_setting(41)) or Dvbbs.strLength(request("name"))<Cint(Dvbbs.Forum_setting(40)) Then
		Dvbbs.AddErrCode(17)
	Else
		username=trim(request("name"))
	End If

	namebadword="=^%^?^&^;^,^'^^$^|^@@@^###"
	namebadword=split(namebadword,"^")
	For i=0 To Ubound(namebadword)
		If Instr(username,namebadword(i))>0 Then
			Dvbbs.AddErrCode(18)
			Exit For
		End If
	Next
	If Instr(request("name"),chr(32))>0 Or Instr(request("name"),chr(34))>0 or Instr(request("name"),chr(9))>0 Then
		Dvbbs.AddErrCode(18)
	End If

	SplitWords=split(Dvbbs.RegSplitWords,",")
	For i = 0 To ubound(splitwords)
		If instr(username,splitwords(i))>0 Then
			Dvbbs.AddErrCode(19)
			Exit For
		End If
	Next
	sex=re_challenge_reg_temp(2)
	password=md5(re_challenge_reg_temp(1),16)
	useremail=re_challenge_reg_temp(4)
	showRe=1
	face="images/userface/image1.gif"
	width=32
	height=32

	If request.Form("birthyear")="" or request.form("birthmonth")="" or request.form("birthday")="" Then
		birthday=""
	Else
		birthday=trim(Request.Form("birthyear"))&"-"&trim(Request.Form("birthmonth"))&"-"&trim(Request.Form("birthday"))
		If not isdate(birthday) Then birthday=""
	End If

	userinfo=checkreal(request.Form("realname")) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
	usersetting=request.Form("setuserinfo") & "|||" & request.Form("setusertrue") & "|||" & showRe

	If Dvbbs.ErrCodes<>"" Then exit sub
	Dim titlepic
	set rs=Dvbbs.Execute("select usertitle,grouppic from Dv_UserGroups where not minarticle=-1 And ParentGID=4 order by minarticle")
	userclass=rs(0)
	titlepic=rs(1)
	UserIM = "|||"&re_challenge_reg_temp(3)&"|||||||||||||||"
	set rs=server.createobject("adodb.recordset")
	sql="select * from [Dv_User] where username='"&username&"' or usermobile='"&re_challenge_reg_temp(0)&"'"
	rs.open sql,conn,1,3
	If not rs.eof and not rs.bof Then
		Dvbbs.AddErrCode(21)
		Exit Sub
	Else
		rs.addnew
		rs("IsChallenge")=1
		rs("username")=username
		rs("userpassword")=password
		rs("TruePassWord")=TruePassWord
		rs("useremail")=useremail
		rs("userclass")=userclass
		rs("titlepic")=titlepic
		rs("UserMobile")=re_challenge_reg_temp(0)
		Rs("UserIM")=UserIM
		Rs("UserPost")=0
		Rs("usergroupid")=4
		rs("lockuser")=0
		Rs("Usersex")=sex
		rs("JoinDate")=NOW()
		rs("Userface")=replace(face,"'","")
		rs("UserWidth")=width
		rs("UserHeight")=height
		rs("UserLogins")=1
		Rs("lastlogin")=NOW()
		rs("userWealth")=Dvbbs.Forum_user(0)
		rs("userEP")=Dvbbs.Forum_user(5)
		rs("usercP")=Dvbbs.Forum_user(10)
		rs("userinfo")=userinfo
		rs("usersetting")=usersetting
		rs("UserFav")="陌生人,我的好友,黑名单"
		rs.update
		Dvbbs.Execute("update Dv_Setup set Forum_usernum=Forum_usernum+1,Forum_lastuser='"&username&"'")
	End If
	rs.close
	set rs=Dvbbs.Execute("select top 1 userid from [Dv_User] order by userid desc")
	dvbbs.userid=rs(0)
	set rs=nothing
	Dvbbs.Name="setup"
	Dvbbs.ReloadSetup

	If Dvbbs.Forum_Setting(47)=1 Then
		'on error resume next
		'发送注册邮件
		Dim getpass
		topic=Replace(template.Strings(35),"{$Forumname}",Dvbbs.Forum_Info(0))

		mailbody = template.html(17)
		mailbody = Replace(mailbody,"{$username}",Dvbbs.HtmlEncode(username))
		mailbody = Replace(mailbody,"{$password}",password)
		mailbody = Replace(mailbody,"{$copyright}",Dvbbs.Forum_Copyright)
		mailbody = Replace(mailbody,"{$version}",Dvbbs.Forum_Version)
		select case Cint(Dvbbs.Forum_Setting(2))
		case 0
			sendmsg=template.Strings(36)
		case 1
			call jmail(useremail,topic,mailbody)
		case 2
			call Cdonts(useremail,topic,mailbody)
		case 3
			call aspemail(useremail,topic,mailbody)
		case Else
			sendmsg=template.Strings(36)
		end select
		If SendMail="OK" Then
			If cint(Dvbbs.Forum_Setting(23))=1 Then
				sendmsg=template.Strings(38)
			Else
				sendmsg=template.Strings(39)
			End If
		Else
			sendmsg=template.Strings(37)
		End If
		Dvbbs.ErrCodes=""
	End If

	If Dvbbs.Forum_Setting(46)=1 Then
		'发送注册短信
		Dim sender,title,body,UserMsg,MsgID
		sender=Dvbbs.Forum_info(0)
		title=Dvbbs.Forum_info(0)&"欢迎您的到来"

		body = template.html(18)
		body = Replace(body,"{$Forumname}",Dvbbs.Forum_Info(0))
		'response.write body
		sql="insert into dv_message(incept,sender,title,content,sendtime,flag,issend) values('"&username&"','"&sender&"','"&title&"','"&body&"',"&SqlNowString&",0,1)"
		Dvbbs.Execute(sql)
		Set rs=Dvbbs.execute("select top 1 ID from [Dv_message] order by ID desc")
		MsgID=rs(0)
		Rs.close:Set Rs=Nothing
		UserMsg="1||"& MsgID &"||"& sender
		Dvbbs.execute("UPDATE [Dv_User] Set UserMsg='"&Dvbbs.CheckStr(UserMsg)&"' WHERE UserID="&Dvbbs.userid)
	End If

	If cint(Dvbbs.Forum_Setting(25))=1 Then

	Else
		Response.Cookies(Dvbbs.Forum_sn).path=dvbbs.cookiepath
		Response.Cookies(Dvbbs.Forum_sn)("username")=""
		Response.Cookies(Dvbbs.Forum_sn)("password")=""
		Response.Cookies(Dvbbs.Forum_sn)("userclass")=""
		Response.Cookies(Dvbbs.Forum_sn)("userid")=""
		Response.Cookies(Dvbbs.Forum_sn)("userhidden")=""
		Response.Cookies(Dvbbs.Forum_sn)("usercookies")=""
		

		Dim StatUserID,UserSessionID
		StatUserID = Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(Dvbbs.UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
		StatUserID = Ccur(StatUserID)
		'客人=SessionID+活动时间+发贴时间+版面ID
		Session(Dvbbs.CacheName & "UserID") = Split(StatUserID & "_" & Now & "_" & Now & "_" & Dvbbs.BoardID,"_")
		Response.Cookies(Dvbbs.Forum_sn).Expires=DateAdd("s",3600,Now())
		Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
		Response.Cookies(Dvbbs.Forum_sn)("StatUserID") = StatUserID
 		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = 0
		Response.Cookies(Dvbbs.Forum_sn)("username") = username
		Response.Cookies(Dvbbs.Forum_sn)("password") = TruePassWord
		Response.Cookies(Dvbbs.Forum_sn)("userclass") = userclass
		Response.Cookies(Dvbbs.Forum_sn)("userid") = dvbbs.userid
		Response.Cookies(Dvbbs.Forum_sn)("userhidden") = 2
		Dvbbs.Execute("delete from dv_online where username='"&dvbbs.membername&"' Or id="&StatUserID&"")
	End If

	Dim TempStr
	TempStr = template.html(22)
	If Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(10)=1 And Dvbbs.Forum_ChanSetting(12)=1 Then
		TempStr = Replace(TempStr,"{$ray_logininfo}",template.html(3))
	Else
		TempStr = Replace(TempStr,"{$ray_logininfo}","")
	End If
	TempStr = Replace(TempStr,"{$reuserpassword}",re_challenge_reg_temp(1))
	TempStr = Replace(TempStr,"{$sendmsg}",sendmsg)
	TempStr = Replace(TempStr,"{$forumname}",Dvbbs.Forum_Info(0))
	Response.Write TempStr
	TempStr=""
	Session("re_challenge_reg_temp")=""

end sub

Function checkreal(v)
Dim w
If not isnull(v) Then
	w=replace(v,"|||","§§§")
	checkreal=w
End If
End Function


Rem ==========论坛登录函数=========
Rem 判断用户登录
Function ChkUserLogin(username,password,mobile,usercookies,ctype)

	Dim rsUser,article,userclass,titlepic
	Dim userhidden,lastip,UserLastLogin
	Dim UserGrade,GroupID,ClassSql,FoundGrade
	Dim regname,iMyUserInfo
	Dim sql,sqlstr,GroupID_Q

	FoundGrade=False
	lastip=Dvbbs.UserTrueIP
	userhidden=request.form("userhidden")
	If not isnumeric(userhidden) and userhidden="" Then userhidden=2
	ChkUserLogin=false
	If mobile<>"" Then
		sqlstr=" UserMobile='"&mobile&"'"
	Else
		sqlstr=" UserName='"&username&"'"
	End If
	'Session(Dvbbs.CacheName & "UserID")用户资料=0dvbbs+1刷新时间+2发贴时间+3所在版面ID+4用户ID+5用户名+6用户密码+7用户邮箱+8用户文章数+9用户主题数+10用户性别+11用户头像+12用户头像宽+13用户头像高+14用户注册时间+15用户最后登陆时间+16用户登陆次数+17用户状态+18用户等级+19用户组ID+20用户组名+21用户金钱+22用户积分+23用户魅力+24用户威望+25用户生日+26最后登陆IP+27用户被删除数+28用户精华数+29用户隐身状态+30用户短信情况+31用户阳光会员+32用户手机+33用户组图标+34用户头衔+35验证密码+36用户今日信息+37+临时数据+38Dvbbs
	Sql="Select UserID,UserName,UserPassword,UserEmail,UserPost,UserTopic,UserSex,UserFace,UserWidth,UserHeight,JoinDate,LastLogin,UserLogins,Lockuser,Userclass,UserGroupID,UserGroup,userWealth,userEP,userCP,UserPower,UserBirthday,UserLastIP,UserDel,UserIsBest,UserHidden,UserMsg,IsChallenge,UserMobile,TitlePic,UserTitle,TruePassWord,UserToday "
	Sql=Sql+" From [Dv_User] Where "&sqlstr&""
	set rsUser=Dvbbs.Execute(sql)
	If rsUser.eof and rsUser.bof Then
		ChkUserLogin=false
		Exit Function
	Else
		iMyUserInfo=rsUser.GetString(,1, "|||", "", "")
		rsUser.Close:Set rsUser = Nothing
	End If
	iMyUserInfo = "Dvbbs|||"& Now & "|||" & Now &"|||"& Dvbbs.BoardID &"|||"& iMyUserInfo &"||||||Dvbbs"
	iMyUserInfo = Split(iMyUserInfo,"|||")
	If trim(password)<>trim(iMyUserInfo(6)) Then
			ChkUserLogin=false
	ElseIf iMyUserInfo(17)=1 Then
			ChkUserLogin=false
	ElseIf iMyUserInfo(19)=5 Then
			ChkUserLogin=false
	Else
			ChkUserLogin=True
			Session(Dvbbs.CacheName & "UserID") = iMyUserInfo
			Dvbbs.UserID = iMyUserInfo(4)
			RegName = iMyUserInfo(5)
			Article = iMyUserInfo(8)
			UserLastLogin = iMyUserInfo(15)
			UserClass = iMyUserInfo(18)			
			GroupID = iMyUserInfo(19)
			TitlePic = iMyUserInfo(34)
			If Article<0 Then Article=0
	End If

	If ChkUserLogin Then
	REM 判断用户等级资料，当用户级别为跟随文章数增长则自动更新等级
	REM 自动更新用户数据
	set rsUser=Dvbbs.Execute("select MinArticle,IsSetting,ParentGID from Dv_UserGroups where usertitle='"&userclass&"'")
	If rsUser.eof and rsUser.bof Then
		'如果没有找到用户等级
		'先判断该组是否有按照文章升级的，也就是MinArticle不是-1的
		set UserGrade=Dvbbs.Execute("select top 1 usertitle,GroupPic,UserGroupID,IsSetting,ParentGID from Dv_UserGroups where (ParentGID="&GroupID&" Or UserGroupID="&GroupID&") and Minarticle<="&article&" and not Minarticle=-1 order by MinArticle desc")
		If not (UserGrade.eof and UserGrade.bof) Then
			userclass=UserGrade(0)
			titlepic=UserGrade(1)
			If UserGrade(3)=1 Then
				GroupID=UserGrade(2)
			Else
				GroupID=UserGrade(4)
			End If
			FoundGrade=True
		End If
		If not FoundGrade Then
			'该组在等级表中不按照文章升级
			set UserGrade=Dvbbs.Execute("select top 1 usertitle,GroupPic,UserGroupID,IsSetting,ParentGID from Dv_UserGroups where UserGroupID="&GroupID&" and Minarticle=-1 order by UserGroupID")
			If not (UserGrade.eof and UserGrade.bof) Then
				userclass=UserGrade(0)
				titlepic=UserGrade(1)
				If UserGrade(3)=1 Then
					GroupID=UserGrade(2)
				Else
					GroupID=UserGrade(4)
				End If
				FoundGrade=True
			End If
			If not FoundGrade Then
			'如果在等级表中未找到相关记录，则使用组名定义等级，采用最低等级用户的图片
			set UserGrade=Dvbbs.Execute("select top 1 GroupPic from Dv_UserGroups where ParentGID>0 And not Minarticle=-1 order by MinArticle")
			titlepic=UserGrade(0)
			set UserGrade=Dvbbs.Execute("select usertitle from Dv_UserGroups where UserGroupID="&GroupID)
			userclass=UserGrade(0)
			End If
		End If
	Else
		'找到用户等级
		'用户等级按照发布文章升级
		If rsUser(0)>-1 Then
			'如果为自定义等级，则取其父类GroupID做升级依据
			GroupID_Q=GroupID
			If RsUser(1)=1 And RsUser(2)>0 Then GroupID_Q=RsUser(2)
			set UserGrade=Dvbbs.Execute("select top 1 usertitle,GroupPic,UserGroupID,IsSetting,ParentGID from Dv_UserGroups where ParentGID="&GroupID_Q&" and Minarticle<="&article&" and not MinArticle=-1 order by MinArticle desc,UserGroupID")
			If not (UserGrade.eof and UserGrade.bof) Then
				userclass=UserGrade(0)
				titlepic=UserGrade(1)
				If UserGrade(3)=1 Then
					GroupID=UserGrade(2)
				Else
					GroupID=UserGrade(4)
				End If
				FoundGrade=True
			End If
			'如果没有相关用户组的等级记录，则采用用户组名称定义等级，采用最低等级用户的图片
			'该情况出现于认证用户组或者添加了用户组没有添加相关等级的用户组
			If not FoundGrade Then
			set UserGrade=Dvbbs.Execute("select top 1 GroupPic from Dv_UserGroups where ParentGID>0 And not Minarticle=-1 order by MinArticle")
			titlepic=UserGrade(0)
			set UserGrade=Dvbbs.Execute("select usertitle from Dv_UserGroups where UserGroupID="&GroupID)
			userclass=UserGrade(0)
			End If
		Else
		'用户等级不按照文章升级
			set UserGrade=Dvbbs.Execute("select usertitle,GroupPic,UserGroupID,IsSetting,ParentGID from Dv_UserGroups where usertitle='"&userclass&"'")
			If not (UserGrade.eof and UserGrade.bof) Then
				userclass=UserGrade(0)
				titlepic=UserGrade(1)
				If UserGrade(3)=1 Then
					GroupID=UserGrade(2)
				Else
					GroupID=UserGrade(4)
				End If
			End If
		End If
	End If
	set rsUser=nothing
	set UserGrade=nothing
	select case ctype
	case 1
		If datediff("d",UserLastLogin,Now())=0 Then
			sql="update [Dv_User] set LastLogin="&SqlNowString&",UserLogins=UserLogins+1,UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&",TruePassWord='"&TruePassWord&"' where userid="&dvbbs.UserID
		Else
			sql="update [Dv_User] set userWealth=userWealth+"&Dvbbs.Forum_user(4)&",userEP=userEP+"&Dvbbs.Forum_user(9)&",userCP=userCP+"&Dvbbs.Forum_user(14)&",LastLogin="&SqlNowString&",UserLogins=UserLogins+1,UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&",TruePassWord='"&TruePassWord&"' where userid="&dvbbs.UserID
		End If
	case 2
		sql="update [Dv_User] set UserPost=UserPost+1,UserTopic=UserTopic+1,userWealth=userWealth+"&Dvbbs.Forum_user(1)&",userEP=userEP+"&Dvbbs.Forum_user(6)&",userCP=userCP+"&Dvbbs.Forum_user(11)&",LastLogin="&SqlNowString&",UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&",TruePassWord='"&TruePassWord&"' where userid="&dvbbs.UserID
	case 3
		sql="update [Dv_User] set UserPost=UserPost+1,userWealth=userWealth+"&Dvbbs.Forum_user(2)&",userEP=userEP+"&Dvbbs.Forum_user(7)&",userCP=userCP+"&Dvbbs.Forum_user(12)&",LastLogin="&SqlNowString&",UserLastIP='"&lastip&"',userclass='"&userclass&"',titlepic='"&titlepic&"',UserGroupID="&GroupID&",TruePassWord='"&TruePassWord&"' where userid="&dvbbs.UserID
	end select
	Dvbbs.Execute(sql)
	Dim StatUserID,UserSessionID
		StatUserID = Dvbbs.checkStr(Trim(Request.Cookies(Dvbbs.Forum_sn)("StatUserID")))
		If IsNumeric(StatUserID) = 0 or StatUserID = "" Then
			StatUserID = Replace(Dvbbs.UserTrueIP,".","")
			UserSessionID = Replace(Startime,".","")
			If IsNumeric(StatUserID) = 0 or StatUserID = "" Then StatUserID = 0
			StatUserID = Ccur(StatUserID) + Ccur(UserSessionID)
		End If
	StatUserID = Ccur(StatUserID)
	Dvbbs.Execute("delete from dv_online where  id="&StatUserID&"")
	If trim(username)<>trim(Dvbbs.membername) Then
		Response.Cookies(Dvbbs.Forum_sn)("username")=""
		Response.Cookies(Dvbbs.Forum_sn)("password")=""
		Response.Cookies(Dvbbs.Forum_sn)("userclass")=""
		Response.Cookies(Dvbbs.Forum_sn)("userid")=""
		Response.Cookies(Dvbbs.Forum_sn)("userhidden")=""
		Response.Cookies(Dvbbs.Forum_sn)("usercookies")=""
		Dvbbs.Execute("delete from dv_online where username='"&Dvbbs.membername&"'")
	End If
	If isnull(usercookies) or usercookies="" Then usercookies="0"
	select case usercookies
	case "0"
		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	case 1
   		Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	case 2
		Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	case 3
		Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
		Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
	end select
	Response.Cookies(Dvbbs.Forum_sn).path = Dvbbs.cookiepath
	Response.Cookies(Dvbbs.Forum_sn)("username") = regname
	Response.Cookies(Dvbbs.Forum_sn)("userid") = Dvbbs.UserID
	Response.Cookies(Dvbbs.Forum_sn)("password") = TruePassWord
	Response.Cookies(Dvbbs.Forum_sn)("userclass") = userclass
	Response.Cookies(Dvbbs.Forum_sn)("userhidden") = userhidden
	rem 清除图片上传数的限制
	Response.Cookies("upNum")=0
	Dim iUserInfo
	iUserInfo = Session(Dvbbs.CacheName & "UserID")
	iUserInfo(35) = TruePassWord
	Session(Dvbbs.CacheName & "UserID") = iUserInfo
	End If
End Function

%>