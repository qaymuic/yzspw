<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/email.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dim Selectinfo(5)
Dvbbs.LoadTemplates("usermanager")
Selectinfo(0)=chk_select("",template.Strings(11))
Selectinfo(1)=chk_select("",template.Strings(12))
Selectinfo(2)=chk_select("",template.Strings(13))
Selectinfo(3)=chk_select("",template.Strings(14))
Selectinfo(4)=Chk_KidneyType("character","",template.Strings(15))
Selectinfo(5)=chk_select("",template.Strings(16))
Dvbbs.LoadTemplates("login")
Dim Stats,ErrCodes
Stats=split(template.Strings(25),"||")
Dvbbs.Stats=Stats(0)
Dvbbs.Nav()
If Cint(dvbbs.Forum_Setting(37))=0 Then
	ErrCodes=ErrCodes+"<li>"+template.Strings(26)
Else	
	If request("action")="apply" Then
		Dvbbs.stats=Stats(2)
		Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
		reg_2()
	ElseIf request("action")="save" Then
		Dvbbs.stats=Stats(3)
		Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
		reg_3()
	ElseIf request("action")="redir" Then
		Dvbbs.stats=Stats(3)
		Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
		redir()
	Else
		Dvbbs.stats=Stats(1)
		Dvbbs.Head_var 0,0,Stats(0),"reg.asp"
		reg_1()
	End If
End If
Dvbbs.Showerr()
If ErrCodes<>"" Then Response.redirect "showerr.asp?ErrCodes="&ErrCodes&"&action=OtherErr"
Dvbbs.ActiveOnline
Dvbbs.Footer()

Sub reg_1()
	Dim TempLateStr
	TempLateStr=template.html(12)
	TempLateStr=Replace(TempLateStr,"{$Forum_Name}",Dvbbs.Forum_Info(0))
	Response.Write TempLateStr
End Sub

Sub reg_2()
	Dim grouploopinfo,TempLateStr,Rs
	TempLateStr=template.html(13)
	If Dvbbs.forum_setting(78)="0" Then
		TempLateStr=Replace(TempLateStr,"{$getcode}","")
	Else
		template.html(24)=Replace(template.html(24),"{$codestr}",Dvbbs.GetCode())
		TempLateStr=Replace(TempLateStr,"{$getcode}",template.html(24))
	End If
	Set Rs=Dvbbs.Execute("select * from DV_GroupName")
	If Rs.eof and Rs.bof Then
		grouploopinfo="<option value=无门无派>无门无派</option>"
	Else
		do while not Rs.eof
		grouploopinfo=grouploopinfo & "<option value="&rs("Groupname")&">"&rs("GroupName")&"</option>"
		Rs.movenext
		loop
	End If
	Rs.close:Set Rs=Nothing
	Dim userregface,i,Forum_userface,FaceDefault
	Forum_userface = split(Dvbbs.Forum_userface,"|||")
	FaceDefault=Forum_userface(0)&Forum_userface(1)
	For i = 1 to Ubound(Forum_userface)-1
		userregface = userregface+"<option value="&Forum_userface(0)&Forum_userface(i)
		userregface = userregface+">"+Forum_userface(i)+"</option>"
	Next
	TempLateStr=Replace(TempLateStr,"{$color}",Dvbbs.mainsetting(1))
	TempLateStr=Replace(TempLateStr,"{$FaceDefault}",FaceDefault)
	TempLateStr=Replace(TempLateStr,"{$Face_select}",userregface)
	TempLateStr=Replace(TempLateStr,"{$FaceMaxWidth}",Dvbbs.Forum_Setting(38))
	TempLateStr=Replace(TempLateStr,"{$FaceMaxHeight}",Dvbbs.Forum_Setting(39))
	TempLateStr=Replace(TempLateStr,"{$ForumFaceMax}",Dvbbs.Forum_Setting(57))
	TempLateStr=Replace(TempLateStr,"{$NameLimLength}",Dvbbs.Forum_Setting(40))
	TempLateStr=Replace(TempLateStr,"{$NameMaxLength}",Dvbbs.Forum_Setting(41))
	TempLateStr=Replace(TempLateStr,"{$Forum_ChanSetting0}",Dvbbs.Forum_ChanSetting(0))
	TempLateStr=Replace(TempLateStr,"{$Forum_ChanSetting9}",Dvbbs.Forum_ChanSetting(9))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting7}",Dvbbs.Forum_Setting(7))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting23}",Dvbbs.Forum_Setting(23))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting32}",Dvbbs.Forum_Setting(32))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting54}",Dvbbs.Forum_Setting(54))
	TempLateStr=Replace(TempLateStr,"{$Forum_Setting42}",Dvbbs.Forum_Setting(42))
	TempLateStr=Replace(TempLateStr,"{$grouploopinfo}",grouploopinfo)
	TempLateStr=Replace(TempLateStr,"{$user_blood}",chk_select("","A,B,AB,O"))
	TempLateStr=Replace(TempLateStr,"{$user_shengxiao}",Selectinfo(0))
	TempLateStr=Replace(TempLateStr,"{$user_occupation}",Selectinfo(1))
	TempLateStr=Replace(TempLateStr,"{$user_marital}",Selectinfo(2))
	TempLateStr=Replace(TempLateStr,"{$user_education}",Selectinfo(3))
	TempLateStr=Replace(TempLateStr,"{$user_character}",Selectinfo(4))
	TempLateStr=Replace(TempLateStr,"{$user_belief}",Selectinfo(5))
	Response.Write TempLateStr
End Sub

'下拉菜单转换输出
Function Chk_select(str1,str2)
	Dim k
	str2=Split(str2,",")
	If IsEmpty(str1) Or str1="" Then chk_select="<option value='' selected>...</option>"
	For k=0 to ubound(str2)
		chk_select=chk_select+"<option value="+str2(k)
		If str2(k)=str1 Then chk_select=chk_select+" selected "
		chk_select=chk_select+" >"+str2(k)+"</option>"
	Next
End Function

'多项选取转换输出
Function Chk_KidneyType(str0,str1,str2)
	Dim k
	str2=split(str2,",")
	For k = 0 to ubound(str2)	
		chk_KidneyType=chk_KidneyType+"<input type=""checkbox"" name="""&str0&""" value="""&trim(str2(k))&""" "	 
		If instr(str1,trim(str2(k)))>0 Then '如果有此项性格
		chk_KidneyType=chk_KidneyType + "checked" 
		End If 
		chk_KidneyType=chk_KidneyType + ">"&trim(str2(k))&" "
	If ((k+1) mod 5)=0 Then chk_KidneyType=chk_KidneyType +  "<br>"  '每行显示六个性格进行换行
	Next
End Function

Sub reg_3()
	Dim username,sex,pass1,pass2,password
	Dim useremail,face,width,height
	Dim sign,showRe,birthday,UserIM
	Dim mailbody,sendmsg,rndnum,num1
	Dim quesion,answer,topic
	Dim userinfo,usersetting
	Dim userclass
	Dim rs,sql,i,TempLateStr
	If not isnull(session("regtime")) or cint(Dvbbs.Forum_Setting(22))>0 Then
		If DateDiff("s",session("regtime"),Now())<cint(Dvbbs.Forum_Setting(22)) Then
			ErrCodes=ErrCodes+"<li>"+Replace(template.Strings(27),"{$RegTime}",Dvbbs.Forum_Setting(22))
			Exit Sub
		End If
	End If
	If Dvbbs.chkpost=false Then
		Dvbbs.AddErrCode(16)
		Exit sub
	End If
	If Request.form("name")="" or strLength(Request.form("name"))>Cint(Dvbbs.Forum_Setting(41)) or strLength(Request.form("name"))<Cint(Dvbbs.Forum_Setting(40)) Then
		TempLateStr=template.Strings(28)
		TempLateStr=Replace(TempLateStr,"{$RegMaxLength}",Dvbbs.Forum_Setting(41))
		TempLateStr=Replace(TempLateStr,"{$RegLimLength}",Dvbbs.Forum_Setting(40))
		ErrCodes=ErrCodes+"<li>"+TempLateStr
		TempLateStr=""
		Exit Sub
	Else
		username=Dvbbs.CheckStr(Trim(Request.form("name")))
	End If

	If Instr(username,"=")>0 or Instr(username,"%")>0 or Instr(username,chr(32))>0 or Instr(username,"?")>0 or Instr(username,"&")>0 or Instr(username,";")>0 or Instr(username,",")>0 or Instr(username,"'")>0 or Instr(username,",")>0 or Instr(username,chr(34))>0 or Instr(username,chr(9))>0 or Instr(username,"")>0 or Instr(username,"$")>0 Then
		Dvbbs.AddErrCode(19)
		Exit sub
	End If
	If Dvbbs.forum_setting(78)="1" Then
		If Not Dvbbs.CodeIsTrue() Then
			 Response.redirect "showerr.asp?ErrCodes=<li>验证码校验失败，请返回刷新页面后再输入验证码。&action=OtherErr"
		End If
	End If
	Dim RegSplitWords
	If Trim(Dvbbs.cachedata(1,0))<>"" Then
	RegSplitWords=split(Dvbbs.cachedata(1,0),"|||")(4)
	RegSplitWords=split(RegSplitWords,",")
		For i = 0 to ubound(RegSplitWords)
			If Trim(RegSplitWords(i))<>"" Then
				If instr(username,RegSplitWords(i))>0 Then
					Dvbbs.AddErrCode(19)
					Exit sub
				End If
			End If 
		next
	End If 
	If Request.form("sex")=0 or Request.form("sex")=1 Then
		sex=Cint(Request.form("sex"))
	Else
		sex=1
	End If
	
	If Request.form("showRe")=0 or Request.form("showRe")=1 Then
		showRe=Request.form("showRe")
	Else
		showRe=1
	End If

	If Cint(Dvbbs.Forum_Setting(23))=1 Then
		Randomize
		Do While Len(rndnum)<8
		num1=CStr(Chr((57-48)*rnd+48))
		rndnum=rndnum&num1
		loop
		password=md5(rndnum,16)
	Else
		If Request.form("psw")="" or len(Request.form("psw"))>10 or len(Request.form("psw"))<6 Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(13)
		Else
			pass1=Request.form("psw")
		End If
		If Request.form("pswc")="" or strLength(Request.form("pswc"))>10 or len(Request.form("pswc"))<6 Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(13)
		Else
			pass2=Request.form("pswc")
		End If
		If pass1<>pass2 Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(29)
		Else
			password=md5(pass2,16)
		End If
	End If

	If Request.form("quesion")="" Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(11)
	Else
		quesion=Request.form("quesion")
	End If
	If Request.form("answer")="" Then
  		ErrCodes=ErrCodes+"<li>"+template.Strings(11)
	ElseIf Request.form("answer")=Request.form("oldanswer") Then
		answer=Request.form("answer")
	Else
		answer=md5(Request.form("answer"),16)
	End If

	If IsValidEmail(Trim(Request.form("e_mail")))=false Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(30)
	Else
		If not Isnull(Dvbbs.Forum_Setting(52)) and Dvbbs.Forum_Setting(52)<>"" and Dvbbs.Forum_Setting(52)<>"0" Then
			Dim SplitUserEmail
			SplitUserEmail=Split(Dvbbs.Forum_Setting(52),"|")
			For i=0 to Ubound(SplitUserEmail)
				If Instr(Request.form("e_mail"),SplitUserEmail(i))>0 Then
				ErrCodes=ErrCodes+"<li>"+template.Strings(31)
				Exit Sub
				End If
			Next
		End If
		useremail=Dvbbs.CheckStr((Request.form("e_mail")))
	End If

	If Request.form("myface")<>"" and Cint(Dvbbs.Forum_Setting(54))=0 Then
		If Request.form("width")="" or Request.form("height")="" Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(32)
		ElseIf Not IsNumeric(Request.form("width")) or not IsNumeric(Request.form("height")) Then
			Dvbbs.AddErrCode(18)
			Exit sub
		ElseIf Cint(Request.form("width"))>Cint(Dvbbs.Forum_Setting(57)) Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(33)
		ElseIf Cint(Request.form("height"))>Cint(Dvbbs.Forum_Setting(57)) Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(33)
		Else
			If Cint(Dvbbs.Forum_Setting(55))=0 Then
				If instr(lcase(Request.form("myface")),"http://")>0 or instr(lcase(Request.form("myface")),"www.")>0 Then
					ErrCodes=ErrCodes+"<li>"+template.Strings(34)
				End If
			End If
			face=Request.form("myface")
		End If
	Else
		If Request.form("face")<>"" Then
			face=Request.form("face")
		End If
	End If
	width=Request.form("width")
	height=Request.form("height")
	If width="" Or Not IsNumeric(width) Then width=CInt(Dvbbs.forum_setting(57))
	If height="" Or Not IsNumeric(height) Then height=CInt(Dvbbs.forum_setting(57))
	width=CInt(width)
	height=CInt(height)
	If Width > CInt(Dvbbs.forum_setting(57)) Then width=CInt(Dvbbs.forum_setting(57))
	If height > CInt(Dvbbs.forum_setting(57)) Then height=CInt(Dvbbs.forum_setting(57))
	birthday=Dvbbs.Checkstr(Trim(Request.Form("birthday")))
	If not Isdate(birthday) Then birthday=""
	userinfo=checkreal(Request.Form("realname")) & "|||" & checkreal(Request.Form("character")) & "|||" & checkreal(Request.Form("personal")) & "|||" & checkreal(Request.Form("country")) & "|||" & checkreal(Request.Form("province")) & "|||" & checkreal(Request.Form("city")) & "|||" & Request.Form("shengxiao") & "|||" & Request.Form("blood") & "|||" & Request.Form("belief") & "|||" & Request.Form("occupation") & "|||" & Request.Form("marital") & "|||" & Request.Form("education") & "|||" & checkreal(Request.Form("college")) & "|||" & checkreal(Request.Form("userphone")) & "|||" & checkreal(Request.Form("address"))
	usersetting=Request.Form("setuserinfo") & "|||" & Request.Form("setusertrue") & "|||" & showRe
	UserIM=checkreal(Request.form("homepage")) &"|||"& checkreal(Request.form("OICQ")) &"|||"& checkreal(Request.form("ICQ")) &"|||"& checkreal(Request.form("msn")) &"|||"& checkreal(Request.form("yahoo")) &"|||"& checkreal(Request.form("aim")) &"|||"& checkreal(Request.form("uc"))
	If ErrCodes<>"" Then Exit Sub
	If Dvbbs.ErrCodes<>"" Then Exit Sub
	Dim titlepic
	Dim TruePassWord
	TruePassWord=Dvbbs.Createpass
	Set Rs=Dvbbs.execute("select usertitle,grouppic,UserGroupID,IsSetting,ParentGID from Dv_UserGroups where not MinArticle=-1 order by MinArticle")
	userclass=rs(0)
	titlepic=rs(1)
	If Rs(3)=1 Then
		Dvbbs.UserGroupID = Rs(2)
	Else
		If Rs(4)=0 Then
			Dvbbs.UserGroupID = Rs(2)
		Else
			Dvbbs.UserGroupID = Rs(4)
		End If
	End If
	set rs=server.createobject("adodb.recordset")
	If request("ischallenge")="yes" and cint(Dvbbs.Forum_Setting(24))=1 Then
		sql="select * from [Dv_user] where username='"&username&"' or useremail='"&useremail&"' or usermobile='"&Dvbbs.CheckStr(request("mobile"))&"'"
	ElseIf request("ischallenge")="yes" Then
		sql="select * from [Dv_user] where username='"&username&"' or usermobile='"&Dvbbs.CheckStr(request("mobile"))&"'"
	ElseIf cint(Dvbbs.Forum_Setting(24))=1 Then
		sql="select * from [Dv_user] where username='"&username&"' or useremail='"&useremail&"'"
	Else
		sql="select * from [Dv_user] where username='"&username&"'"
	End If
	'Response.Write sql
	'response.end
	rs.open sql,conn,1,3
	If not rs.eof and not rs.bof Then
		If Dvbbs.Forum_Setting(24)="1" Then
			Dvbbs.AddErrCode(20)
			Exit sub
		Else
			Dvbbs.AddErrCode(21)
			Exit Sub
		End If
	Else
	rs.addnew
		rs("UserName")=username
		rs("UserPassword")=password
		rs("UserEmail")=useremail
		rs("Userclass")=userclass
		rs("TitlePic")=titlepic
		rs("UserQuesion")=quesion
		rs("UserAnswer")=answer
		rs("TruePassWord")=TruePassWord
		rs("UserIM")=UserIM
		If Request.Form("Signature")<>"" Then rs("UserSign")=Dvbbs.Htmlencode(Trim(Request.Form("Signature")))
		rs("UserPost")=0
		If Dvbbs.Forum_Setting(25)="1" Then
			rs("UserGroupID")=5
		Else
		   	rs("UserGroupID")=Dvbbs.UserGroupID
		End If
		rs("Lockuser")=0
		rs("UserSex")=sex
		If birthday<>"" Then rs("UserBirthday")=birthday
		rs("UserGroup")=Request.form("UserGroup")
		rs("JoinDate")=NOW()
		If Request.form("myface")<>"" Then
			
			rs("UserFace")=replace(face,"'","")
		Else
			rs("UserFace")=replace(face,"'","")
		End If
		rs("UserWidth")=width
		rs("UserHeight")=height
		rs("UserLogins")=1
		rs("LastLogin")=NOW()
		rs("userWealth")=dvbbs.Forum_user(0)
		rs("userEP")=dvbbs.Forum_user(5)
		rs("usercP")=dvbbs.Forum_user(10)
		rs("UserInfo")=userinfo
		rs("UserSetting")=usersetting
		rs("UserPower")=0
		rs("UserDel")=0
		rs("UserIsbest")=0
		rs("UserFav")="陌生人,我的好友,黑名单"
		rs("IsChallenge")=0
		rs("UserLastIP")=Request.ServerVariables("REMOTE_ADDR")
		rs.update
		Dvbbs.execute("UpDate Dv_Setup Set Forum_UserNum=Forum_UserNum+1,Forum_lastUser='"&username&"'")
		
	End If
	rs.close
	Dvbbs.ReloadSetupCache username,14
	Dvbbs.ReloadSetupCache (CLng(Dvbbs.CacheData(10,0))+1),10 
	Dim facename
	Set rs=Dvbbs.execute("select top 1 userid,UserFace from [Dv_user] order by userid desc")
		Dvbbs.userid=rs(0)
		facename=rs(1)
	rs.close
	set rs=nothing

	'******************
	'对上传头象进行过滤与改名
	If Cint(Dvbbs.Forum_Setting(7))=1 Then 
		on error resume next
		Dim objFSO,upface,newfilename
		facename=Replace(facename,"\","/")
		facename=Replace(facename,"//","/")
		facename=Replace(facename,"..","")
		facename=Replace(facename,"^","")
		facename=Replace(facename,"@","")
		facename=Replace(facename,"%","")
		If instr(Lcase(facename),"uploadface/") Then
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			facename=objFSO.GetFileName(facename)
			upface="uploadFace/"&facename
			newfilename="uploadFace/"&Dvbbs.userid&"_"&facename
			if	objFSO.fileExists(Server.MapPath(upface)) Then
				objFSO.movefile ""&Server.MapPath(upface)&"",""&Server.MapPath(newfilename)&""
				If Not Err Then
					Dvbbs.execute("update [Dv_user] set UserFace='"&replace(newfilename,"'","")&"' Where userid="&Dvbbs.userid)
				End If
			End If
			set objFSO=nothing
		End If
	End If
	'对上传头象进行过滤与改名结束
	'****************

	If Dvbbs.Forum_Setting(47)=1 Then
		'on error resume next
		'发送注册邮件
		Dim getpass
		topic=Replace(template.Strings(35),"{$Forumname}",Dvbbs.Forum_Info(0))
		If cint(Dvbbs.Forum_Setting(23))=1 Then
			getpass=Dvbbs.htmlencode(rndnum)
		Else
			getpass=Dvbbs.htmlencode(Request.form("psw"))
		End If
		mailbody = template.html(17)
		mailbody = Replace(mailbody,"{$username}",Dvbbs.HtmlEncode(username))
		mailbody = Replace(mailbody,"{$password}",getpass)
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
		'response.write mailbody
	End If

	If Dvbbs.Forum_Setting(46)="1" Then
		'发送注册短信
		Dim sender,title,body,UserMsg,MsgID
		sender=Dvbbs.Forum_Info(0)
		title=Dvbbs.lanstr(2)&Dvbbs.Forum_Info(0)
		body = template.html(18)
		body = Replace(body,"{$Forumname}",Dvbbs.Forum_Info(0))
		sql="insert into dv_message(incept,sender,title,content,sendtime,flag,issend) values('"&username&"','"&sender&"','"&title&"','"&body&"',"&SqlNowString&",0,1)"
		Dvbbs.Execute(sql)
		Set rs=Dvbbs.execute("select top 1 ID from [Dv_message] order by ID desc")
		MsgID=rs(0)
		Rs.close:Set Rs=Nothing
		UserMsg="1||"& MsgID &"||"& sender
		Dvbbs.execute("UPDATE [Dv_User] Set UserMsg='"&Dvbbs.CheckStr(UserMsg)&"' WHERE UserID="&Dvbbs.userid)
	End If

	If cint(Dvbbs.Forum_Setting(23))=1 or cint(Dvbbs.Forum_Setting(25))=1 Then

	Else
		Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
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
		Dvbbs.Execute("delete from dv_online where username='"&dvbbs.membername&"' Or id="&StatUserID&"")
		'客人=SessionID+活动时间+发贴时间+版面ID
		Session(Dvbbs.CacheName & "UserID") = Split(StatUserID & "_" & Now & "_" & Now & "_" & Dvbbs.BoardID,"_")
		Response.Cookies(Dvbbs.Forum_sn)("StatUserID") = StatUserID
		select case request("usercookies")
	 	case 0
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = request("usercookies")
		Case 1
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = request("usercookies")
		Case 2
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = request("usercookies")
	    case 3
			Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
			Response.Cookies(Dvbbs.Forum_sn)("usercookies") = request("usercookies")
		end select
		Response.Cookies(Dvbbs.Forum_sn)("username") = username
		Response.Cookies(Dvbbs.Forum_sn)("password") = TruePassWord
		Response.Cookies(Dvbbs.Forum_sn)("userclass") = userclass
		Response.Cookies(Dvbbs.Forum_sn)("userid") = Dvbbs.userid
		Response.Cookies(Dvbbs.Forum_sn)("userhidden") = 2
		Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
		Dvbbs.membername=username
		Dvbbs.userhidden=2
		Dvbbs.MemberClass=userclass
	End If
	session("regtime")=now()

	If request("ischallenge")="yes" and Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(9)=1 Then
		Get_ChallengeWord
		Session("challengeUserID")=Dvbbs.UserID
		If cint(request("sex"))=1 Then
			sex="F"
		Else
			sex="M"
		End If
		set rs=Dvbbs.Execute("select top 1 * from Dv_ChallengeInfo")
		Dim MyForumID
		MyForumID=rs("D_ForumID")
		Response.Write Replace(template.html(14),"{$Forumname}",Dvbbs.Forum_Info(0))
%>
<form name="redir" action="http://bbs.ray5198.com/user_register.jsp" method="post">
<INPUT type=hidden name="username" value="<%=username%>">
<INPUT type=hidden name="forumPwd" value="<%=Request.form("psw")%>">
<INPUT type=hidden name="forumUrl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<INPUT type=hidden name="mobile" value="<%=request("mobile")%>">
<INPUT type=hidden name="sex" value="<%=sex%>">
<INPUT type=hidden name="qq" value="<%=Request.form("Oicq")%>">
<INPUT type=hidden name="email" value="<%=useremail%>">
<INPUT type=hidden name="forumId" value="<%=MyForumID%>">
<input type=hidden value="<%=Session("challengeWord")%>" name="challengeWord">
<input type=hidden value="reg.asp?action=redir" name="dirPage">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
Else
	TempLateStr=template.html(15)
	TempLateStr=Replace(TempLateStr,"{$Forumname}",Dvbbs.Forum_Info(0))
	TempLateStr=Replace(TempLateStr,"{$sendmsg}",sendmsg)
	Response.Write TempLateStr
End If
End Sub

Function redir()
	Dim ErrorCode,ErrorMsg
	Dim remobile,rechallengeWord,retokerWord
	Dim challengeWord_key,rechallengeWord_key
	ErrorCode=trim(request("ErrorCode"))
	ErrorMsg=trim(request("ErrorMsg"))
	remobile=trim(Dvbbs.CheckStr(request("mobile")))
	rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
	retokerWord=trim(request("tokenWord"))
	select case ErrorCode
	case 100
		If challengeWord_key=retokerWord Then
			Dvbbs.Execute("update [Dv_user] set UserMobile='"&remobile&"',IsChallenge=1 where userid="&Session("challengeUserID"))
		Else
			ErrCodes=ErrCodes+"<li>"+template.Strings(40)
			ErrCodes=ErrCodes+"<li>"+template.Strings(41) & ErrorMsg
			Exit Function
		End If
	case 101
		ErrCodes=ErrCodes+"<li>"+template.Strings(40)
		ErrCodes=ErrCodes+"<li>"+template.Strings(42) & ErrorMsg
		Exit Function
	case 102
		ErrCodes=ErrCodes+"<li>"+template.Strings(40)
		ErrCodes=ErrCodes+"<li>"+template.Strings(42) & ErrorMsg
		Exit Function
	case Else
		ErrCodes=ErrCodes+"<li>"+template.Strings(40)
		ErrCodes=ErrCodes+"<li>高级用户注册失败，此手机已经在当前论坛注册过" & ErrorMsg
		Exit Function
	end select
	Response.Write Replace(Replace(template.html(15),"{$Forumname}",Dvbbs.Forum_Info(0)),"{$sendmsg}",template.Strings(47))
End Function

Function checkreal(v)
Dim w
If not isnull(v) Then
	w=replace(v,"|||","§§§")
	checkreal=w
End If
End Function
%>