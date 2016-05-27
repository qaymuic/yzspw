<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/chkinput.asp"-->
<%
'mymodify.asp
Dvbbs.LoadTemplates("usermanager")
Dvbbs.Stats=Dvbbs.MemberName&template.Strings(1)
Dvbbs.Nav()
Dvbbs.Head_var 0,0,template.Strings(0),"usermanager.asp"
Dim Sql,Rs,TempStr,ErrCodes

If Clng(Dvbbs.GroupSetting(16))=0 Then
	Dvbbs.AddErrCode(28)
	Dvbbs.Showerr()
End If

If Dvbbs.userid=0 Then
	Dvbbs.AddErrCode(6)
	Dvbbs.Showerr()
Else
	Response.Write Template.Html(0)
	If Request("action")="updat" Then
		update()
	Else
		Userinfo()
	End If
	If ErrCodes<>"" Then Response.redirect "showerr.asp?ErrCodes="&ErrCodes&"&action=OtherErr"
End If
Dvbbs.ActiveOnline()
Dvbbs.Footer()

Sub Userinfo()
	Dim CanUseTitle,CanUseTitle1,CanUseTitle2,i,CanUserInfo
	Dim My_info,My_infotemp,My_Cookies,ShowUserInfo
	Dim UseRsetting,SetUserInfo,SetUserTrue,ShowRe
	Dim signtrue
	My_infotemp=Template.Html(5)
	My_Cookies=Request.Cookies(Dvbbs.Forum_sn)("usercookies")
	CanUseTitle=False
	CanUseTitle1=False
	CanUseTitle2=False
	CanUserInfo=False
	'UserID=0,UserName=1,UserEmail=2,UserPost=3,UseRsign=4,UseRsex=5,UserFace=6,UserWidth=7,UserHeight=8,JoinDate=9,UserGroup=10,UserTitle=11,UserBirthday=12,UserPhoto=13,UserInfo=14,UseRsetting=15
	sql="Select UserID,UserName,UserEmail,UserPost,UseRsign,UseRsex,UserFace,UserWidth,UserHeight,JoinDate,UserGroup,UserTitle,UserBirthday,UserPhoto,UserInfo,UseRsetting from [DV_User] where userid="&Dvbbs.userid
	Set Rs=Dvbbs.Execute(Sql)
	If Rs.eof And Rs.bof Then
		Dvbbs.AddErrCode(32)
		Exit Sub
	Else
		Sql=Rs.GetString(,,"###","","")
	End If
	Rs.close :Set Rs=Nothing
	My_info= Split(Sql,"###")
	If Clng(Dvbbs.Forum_Setting(6))=1 Then CanUseTitle=True

	If  CanUseTitle and Clng(Dvbbs.Forum_Setting(60))>0 and Clng(My_info(3))>Clng(Dvbbs.Forum_Setting(60)) Then
		CanUseTitle1=True
	ElseIf CanUseTitle and Clng(Dvbbs.Forum_Setting(60))=0 Then
		CanUseTitle1=True
	Else 
		CanUseTitle1=False
	End If

	If CanUseTitle and Clng(Dvbbs.Forum_Setting(61))>0 And DateDiff("d",My_info(9),Now())>Clng(Dvbbs.Forum_Setting(61)) Then
		CanUseTitle2=True
	ElseIf CanUseTitle And Clng(Dvbbs.Forum_Setting(61))=0 Then
		CanUseTitle2=True
	Else
		CanUseTitle2=False
	End If

	If CanUseTitle And Clng(Dvbbs.Forum_Setting(62))=1 Then 
		If CanUseTitle1 And CanUseTitle2 Then 
			CanUseTitle=True 
		Else
			CanUseTitle=False
		End If 
	ElseIf CanUseTitle And (CanUseTitle1 or CanUseTitle2) Then
		CanUseTitle=True 
	Else
		CanUseTitle=False 
	End If
	signtrue=My_info(4)
	If My_info(14)<>"" Then
		ShowUserInfo=split(My_info(14),"|||")
		If ubound(ShowUserInfo)=14 Then
		CanUserInfo=True
		End If
	End If

	UseRsetting=split(My_info(15),"|||")
	If UBound(UseRsetting)=2 Then
		If isnumeric(UseRsetting(0)) Then Setuserinfo=Clng(UseRsetting(0)) Else Setuserinfo=1
		If isnumeric(UseRsetting(1)) Then Setusertrue=Clng(UseRsetting(1)) Else Setusertrue=0
		If isnumeric(UseRsetting(2)) Then ShowRe=Clng(UseRsetting(2)) Else ShowRe=0
	Else
		Setuserinfo=1
		Setusertrue=0
		ShowRe=0
	End If
	If (Clng(Dvbbs.Forum_Setting(54))>0 And Clng(My_info(3))>Clng(Dvbbs.Forum_Setting(54))) Or Clng(Dvbbs.Forum_Setting(54))=0 Then
		My_infotemp=Replace(My_infotemp,"{$SetFace_info}",SetUserFace(Clng(Dvbbs.Forum_Setting(7)),My_info(6)&"",My_info(7),My_info(8)))
	Else
		My_infotemp=Replace(My_infotemp,"{$SetFace_info}","")
	End If

	If Clng(Dvbbs.Forum_Setting(32))=1 Then
		My_infotemp=Replace(My_infotemp,"{$SetGroup_info}",SetUserGroup(My_info(10)))
	Else
		My_infotemp=Replace(My_infotemp,"{$SetGroup_info}","")
	End If
	My_infotemp=Replace(My_infotemp,"{$user_Id}",My_info(0))
	My_infotemp=Replace(My_infotemp,"'{$Dvbbs.FoundIsChallenge}'",Lcase(Dvbbs.FoundIsChallenge))
	If CanUseTitle Then
		My_Infotemp = Replace(My_Infotemp, "{$SetTitle_info}", SetUserTitle(Dvbbs.Htmlencode(My_info(11))))
	Else
		My_Infotemp = Replace(My_infotemp, "{$SetTitle_info}", "")
	End If
	My_infotemp=Replace(My_infotemp,"{$checked_sex}",My_info(5))
	My_infotemp=Replace(My_infotemp,"{$user_Birthday}",My_info(12))
	My_infotemp=Replace(My_infotemp,"{$user_Photo}",Dvbbs.htmlencode(Trim(My_info(13))))
	My_infotemp=Replace(My_infotemp,"{$user_Signature}",signtrue)
	My_infotemp=Replace(My_infotemp,"{$showRe}",ShowRe)
	My_infotemp=Replace(My_infotemp,"{$user_Cookies}",My_Cookies)
	My_infotemp=Replace(My_infotemp,"{$user_Setuserinfo}",Setuserinfo)
	My_infotemp=Replace(My_infotemp,"{$user_Setusertrue}",Setusertrue)

	If CanUserInfo=True Then
		My_infotemp=Replace(My_infotemp,"{$user_Realname}",ShowUserInfo(0))
		My_infotemp=Replace(My_infotemp,"{$user_character}",Chk_KidneyType("character",ShowUserInfo(1),template.Strings(15)))
		My_infotemp=Replace(My_infotemp,"{$user_Personal}",ShowUserInfo(2))
		My_infotemp=Replace(My_infotemp,"{$user_Country}",ShowUserInfo(3))
		My_infotemp=Replace(My_infotemp,"{$user_Province}",ShowUserInfo(4))
		My_infotemp=Replace(My_infotemp,"{$user_City}",ShowUserInfo(5))
		My_infotemp=Replace(My_infotemp,"{$user_College}",ShowUserInfo(12))
		My_infotemp=Replace(My_infotemp,"{$user_Phone}",ShowUserInfo(13))
		My_infotemp=Replace(My_infotemp,"{$user_Address}",ShowUserInfo(14))
		My_infotemp=Replace(My_infotemp,"{$user_shengxiao}",chk_select(ShowUserInfo(6),template.Strings(11)))
		My_infotemp=Replace(My_infotemp,"{$user_blood}",chk_select(ShowUserInfo(7),"A,B,AB,O"))
		My_infotemp=Replace(My_infotemp,"{$user_belief}",chk_select(ShowUserInfo(8),template.Strings(16)))
		My_infotemp=Replace(My_infotemp,"{$user_occupation}",chk_select(ShowUserInfo(9),template.Strings(12)))
		My_infotemp=Replace(My_infotemp,"{$user_marital}",chk_select(ShowUserInfo(10),template.Strings(13)))
		My_infotemp=Replace(My_infotemp,"{$user_education}",chk_select(ShowUserInfo(11),template.Strings(14)))
	Else
		My_infotemp=Replace(My_infotemp,"{$user_Realname}","")
		My_infotemp=Replace(My_infotemp,"{$user_character}",Chk_KidneyType("character","",template.Strings(15)))
		My_infotemp=Replace(My_infotemp,"{$user_Personal}","")
		My_infotemp=Replace(My_infotemp,"{$user_Country}","")
		My_infotemp=Replace(My_infotemp,"{$user_Phone}","")
		My_infotemp=Replace(My_infotemp,"{$user_Address}","")
		My_infotemp=Replace(My_infotemp,"{$user_Province}","")
		My_infotemp=Replace(My_infotemp,"{$user_City}","")
		My_infotemp=Replace(My_infotemp,"{$user_Cartype}","")	
		My_infotemp=Replace(My_infotemp,"{$user_College}","")
		My_infotemp=Replace(My_infotemp,"{$user_shengxiao}",chk_select("",template.Strings(11)))
		My_infotemp=Replace(My_infotemp,"{$user_blood}",chk_select("","A,B,AB,O"))
		My_infotemp=Replace(My_infotemp,"{$user_belief}",chk_select("",template.Strings(16)))
		My_infotemp=Replace(My_infotemp,"{$user_occupation}",chk_select("",template.Strings(12)))
		My_infotemp=Replace(My_infotemp,"{$user_marital}",chk_select("",template.Strings(13)))
		My_infotemp=Replace(My_infotemp,"{$user_education}",chk_select("",template.Strings(14)))
	End If
	Response.write My_infotemp
End Sub

Sub update()
	If Dvbbs.chkpost=False Then
		Dvbbs.AddErrCode(16)
		Exit Sub
	End If
	Dim CanUseTitle,CanUseTitle1,CanUseTitle2
	Dim SplitUserTitle,i,sex,showRe,face,width,height,birthday,usercookies,usertitle
	CanUseTitle=false
	CanUseTitle1=false
	CanUseTitle2=false
	If Not Dvbbs.FoundIsChallenge Then
		If Request.Form("sex")="" Then
			Dvbbs.AddErrCode(18)
		ElseIf isInteger(Request.Form("sex")) Then
			sex=Request.Form("sex")
		Else
			Dvbbs.AddErrCode(18)
		End If
	End If
	
	If Request.Form("showRe")="" Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(17)
	ElseIf isInteger(Request.Form("showRe")) Then
		showRe=Clng(Request.Form("showRe"))
	Else
		Dvbbs.AddErrCode(18)
	End If
			
	If Request.Form("myface")<>"" and ((Clng(Dvbbs.Forum_Setting(54))>0 and Clng(Dvbbs.MyUserInfo(8))>Clng(Dvbbs.Forum_Setting(54))) or Clng(Dvbbs.Forum_Setting(54))=0) Then
		If Request.Form("width")="" or Request.Form("height")="" Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(18)
		ElseIf not isInteger(Request.Form("width")) or not isInteger(Request.Form("height")) Then
			Dvbbs.AddErrCode(18)
		ElseIf Clng(Request.Form("width"))>Clng(Dvbbs.Forum_Setting(57)) Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(19)
		ElseIf Clng(Request.Form("height"))>Clng(Dvbbs.Forum_Setting(57)) Then
			ErrCodes=ErrCodes+"<li>"+template.Strings(20)
		Else
			If Clng(Dvbbs.Forum_Setting(55))=0 Then
				If InStr(lcase(Request.Form("myface")),"http://")>0 or instr(lcase(Request.Form("myface")),"www.")>0 Then
					ErrCodes=ErrCodes+"<li>"+template.Strings(21)
				End If
			End If
			Face=Request.Form("myface")
			width=Request.Form("width")
			height=Request.Form("height")
		End If
	Else
		Dvbbs.Forum_userface = Split(Dvbbs.Forum_userface,"|||")
		If Request.Form("face")="" Then
			Face=Dvbbs.Forum_userface(0)&Dvbbs.Forum_userface(1)
		Else
			Face=Request.Form("face")
		End If
	End If
	face=Dv_FilterJS(Replace(face,"'",""))
	face=Replace(face,"..","")
	face=Replace(face,"\","/")
	face=Replace(face,"^","")
	face=Replace(face,"#","")
	face=Replace(face,"%","")
	If width="" or height="" Then
		width=Dvbbs.Forum_Setting(38)
		height=Dvbbs.Forum_Setting(39)
	End If
	If Dvbbs.StrLength(Request.Form("Signature"))>250 Then
		ErrCodes=ErrCodes+"<li>"+template.Strings(23)
	End If
	birthday=trim(Request.Form("birthday"))
	If Not IsDate(birthday) Then birthday=""
	Dim userinfo,useRsetting
	userinfo=checkreal(Request.Form("realname")) & "|||" & checkreal(Request.Form("character")) & "|||" & checkreal(Request.Form("peRsonal")) & "|||" & checkreal(Request.Form("country")) & "|||" & checkreal(Request.Form("province")) & "|||" & checkreal(Request.Form("city")) & "|||" & Request.Form("shengxiao") & "|||" & Request.Form("blood") & "|||" & Request.Form("belief") & "|||" & Request.Form("occupation") & "|||" & Request.Form("marital") & "|||" & Request.Form("education") & "|||" & checkreal(Request.Form("college")) & "|||" & checkreal(Request.Form("userphone")) & "|||" & checkreal(Request.Form("address"))
	usersetting=Request.Form("setuserinfo") & "|||" & Request.Form("setusertrue") & "|||" & showRe
	Dim UpSessionID
	UpSessionID=Session(Dvbbs.CacheName & "UserID")
	UpSessionID(11)=Trim(face)
	UpSessionID(12)=Trim(Width)
	UpSessionID(13)=Trim(height)
	UpSessionID(25)=Trim(birthday)
	If ErrCodes<>"" Then Exit Sub
	Set Rs=server.createobject("adodb.recordset")
	If Not IsObject(Conn) Then ConnectionDatabase
	sql="Select * from [Dv_User] where userid="&Dvbbs.UserID
	Rs.open sql,conn,1,3
	If Rs.EOF And Rs.BOF Then
		Dvbbs.AddErrCode(12)
	Else
		Rs("UserFace")=face
		Rs("UserWidth")=width
		Rs("UserHeight")=height
		If Not Dvbbs.FoundIsChallenge Then Rs("UseRsex")=sex
		Rs("UserSign")=Request.Form("Signature")
		Rs("UserPhoto")=Dv_FilterJS(Request.Form("userphoto"))
		If Dvbbs.Forum_Setting(32)="1" And IsTrueGroupName(Request.Form("groupname")) Then 
			Rs("UserGroup")=Dvbbs.iHtmlEncode(Request.Form("groupname"))
		Else
			Rs("UserGroup")=""
		End If
		'判断是否允许提交头衔
		If Clng(Dvbbs.Forum_Setting(6))=1 Then
			CanUseTitle=True 
		End If
		If CanUseTitle and Clng(Dvbbs.Forum_Setting(60))>0 and Rs("UserPost")>Clng(Dvbbs.Forum_Setting(60)) Then
			CanUseTitle1=True 
		ElseIf CanUseTitle and Clng(Dvbbs.Forum_Setting(60))=0 Then
			CanUseTitle1=True 
		Else
			CanUseTitle1=False 
		End If
		If CanUseTitle and Clng(Dvbbs.Forum_Setting(61))>0 and DateDiff("d",Rs("JoinDate"),Now())>Clng(Dvbbs.Forum_Setting(61)) Then
			CanUseTitle2=True 
		ElseIf CanUseTitle and Clng(Dvbbs.Forum_Setting(61))=0 Then
			CanUseTitle2=True 
		Else
			CanUseTitle2=False 
		End If
		If CanUseTitle and Clng(Dvbbs.Forum_Setting(62))=1 Then
			If CanUseTitle1 and CanUseTitle2 Then
				CanUseTitle=True
			Else
				CanUseTitle=False
			End If
		ElseIf CanUseTitle and (CanUseTitle1 or CanUseTitle2) Then
			CanUseTitle=True 
		Else
			CanUseTitle=False
		End If
		usertitle = Dvbbs.iHtmlencode(Request.Form("title"))
		If CanUseTitle Then
			If Trim(Dvbbs.Forum_Setting(63))<>"" Then
				SplitUserTitle=split(Dvbbs.Forum_Setting(63),"|")
				For i=0 to ubound(SplitUserTitle)
					If InStr(lcase(usertitle),lcase(SplitUserTitle(i)))>0 Then
						ErrCodes=ErrCodes+"<li>"+template.Strings(24)
						Exit sub
					End If
				Next
			End If
			If Len(usertitle)>Clng(Dvbbs.Forum_Setting(59)) Then
				ErrCodes=ErrCodes+"<li>"+Replace(template.Strings(25),"{$MaxTitleLen}",Dvbbs.Forum_Setting(59))
				Exit Sub
			End If
			Rs("UserTitle")=usertitle
			UpSessionID(34)=usertitle
		End If


		If birthday<>"" Then Rs("UserBirthday")=birthday
		Rs("Userinfo")=trim(Userinfo)
		Rs("UseRsetting")=trim(UseRsetting)
		Rs.Update
		usercookies=Request.Form("usercookies")
		If IsNumeric(usercookies) Then usercookies=Clng(usercookies) Else usercookies=0
		Select Case usercookies
			Case 0
				Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
			Case 1
				Response.Cookies(Dvbbs.Forum_sn).Expires=Date+1
				Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
			Case 2
				Response.Cookies(Dvbbs.Forum_sn).Expires=Date+31
				Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
			Case 3
				Response.Cookies(Dvbbs.Forum_sn).Expires=Date+365
				Response.Cookies(Dvbbs.Forum_sn)("usercookies") = usercookies
		End Select
		Response.Cookies(Dvbbs.Forum_sn).path=Dvbbs.cookiepath
	End If
	Rs.Close
	Set Rs=Nothing
	Session(Dvbbs.CacheName & "UserID")=UpSessionID
	Dvbbs.Dvbbs_Suc("<li>"+template.Strings(26))
End sub

Function checkreal(v)
	Dim w
	If not isnull(v) Then
		w=replace(v,"|||","§§§")
		checkreal=w
	End If
End Function

Function IsTrueGroupName(GroupName)
	IsTrueGroupName=False
	If GroupName="" Then Exit Function
	Dim tRs
	Set tRs=dvbbs.Execute("Select GroupName From [Dv_GroupName]")
	Do While Not tRs.EOF
		If GroupName=tRs(0) Then
			IsTrueGroupName=True
			Exit Do 
		End If 
	tRs.MoveNext
	Loop
	tRs.close:Set tRs=Nothing 
End Function 

'用户头衔输出
Function SetUserTitle(str)
	SetUserTitle=template.html(6)
	SetUserTitle=Replace(SetUserTitle,"{$user_Title}",str)
End Function

'str=0 关闭显示上传头像表单
Function SetUserFace(str,face,wid,hig)
Dim tempstr,facetemp,userregface,i
	tempstr = Split(template.html(7),"||")
	Dvbbs.Forum_userface = split(Dvbbs.Forum_userface,"|||")
	For i = 1 to Ubound(Dvbbs.Forum_userface)-1
		userregface = userregface+"<option value="+Dvbbs.Forum_userface(0)&Dvbbs.Forum_userface(i)
		If trim(lcase(userregface)) = trim(lcase(face)) then userregface = userregface+" selected "
		userregface = userregface+">"+Dvbbs.Forum_userface(i)+"</option>"
	Next
	If str = 0 Then
	SetUserFace = tempstr(0)+tempstr(2)
	Else
	SetUserFace = tempstr(0)+tempstr(1)+tempstr(2)
	End If
	SetUserFace=Replace(SetUserFace,"{$Face_select}",userregface)
	SetUserFace=Replace(SetUserFace,"{$color}",Dvbbs.mainsetting(1))
	SetUserFace=Replace(SetUserFace,"{$user_Face}",Dv_FilterJS(face))
	SetUserFace=Replace(SetUserFace,"{$user_FaceWidth}",wid)
	SetUserFace=Replace(SetUserFace,"{$user_FaceHeight}",hig)
	SetUserFace=Replace(SetUserFace,"{$forum_Mwidth}",Dvbbs.Forum_Setting(57))
	SetUserFace=Replace(SetUserFace,"{$forum_Mheight}",Dvbbs.Forum_Setting(57))
End Function

'用户门派输出
Function SetUserGroup(str)
	Dim tempstr
	Set Rs=Dvbbs.Execute("Select GroupName From [Dv_GroupName]")
	Do While Not Rs.EOF
		tempstr=tempstr+"<option value="&Rs(0)
		If trim(Rs(0))=trim(str) Then tempstr=tempstr+" selected"
		tempstr=tempstr+" > "&Rs(0)&" </option>"
		Rs.MoveNext
	Loop
	Rs.Close
	SetUserGroup=Replace(template.html(8),"{$user_GroupName}",tempstr)
End Function

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

Rem 判断数字是否整形
Function isInteger(Para)
	isInteger=False
	If Not (IsNull(Para) Or Trim(Para)="" Or Not IsNumeric(Para)) Then
		isInteger=True
	End If
End Function

Function Dv_FilterJS(v)
	If  Not Isnull(V) Then
		Dim t
		Dim re
		Dim reContent
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(%)"
		t=re.Replace(v,"<I>%</I>")
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
		Dv_FilterJS=Trim(t)
		Set Re=Nothing
	End If 
End Function
%>