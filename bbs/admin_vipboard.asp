<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!--#include file="inc/md5.asp"-->
<!-- #include file="inc/DvADChar.asp" -->
<!--#include file="inc/chan_const.asp"-->
<%
Head()

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(8)=1) Then
	Errmsg=Errmsg+"<br>"+"<li>本论坛没有开启VIP收费论坛功能。"
	call dvbbs_error()
Else
	If Not dvbbs.master or instr(","&session("flag")&",",",9,")=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	Else
		Dim action
		action=Request("action")
		Select Case action
			Case "reg1"
				Reg1()
			Case "result"
				result()
			Case "result1"
				result1()
			Case "showvipuser"
				showvipuser()
			Case "reinstall"
				Reg2()
			Case Else
				Main()
		End Select
	end if
end if
Sub Reg1()
	If Dvbbs.BoardID=0 Then 
		Errmsg=ErrMsg + "<BR><li>错误的板面参数。"
		dvbbs_error()
	End If
	'校验表单数据
	Dim vipshow,vipintegral,vipusetime,vipid
	Dim NowStr
	vipshow=Request("vipshow")
	vipintegral=Request("vipintegral")
	vipusetime=Request("vipusetime")
	session("vipintegral")=vipintegral
	Session("vipusetime")=vipusetime
	NowStr=datediff("s","1970-1-1",Now)
	vipid=Dvbbs.BoardID
	vipid=cstr(NowStr) & cstr(vipid)
	vipid=md5(vipid,32)
	Session("NowStr")=NowStr
	If vipshow="" Then 
		Errmsg=ErrMsg + "<BR><li>请填写VIP申请说明。"
		dvbbs_error()
	ElseIf len(vipshow) > 1024 Then 
		Errmsg=ErrMsg + "<BR><li>VIP申请说明的长度不能大于1024字节。"
		dvbbs_error()
	End If
	If Not IsNumeric(vipintegral) or InStr(vipintegral,".")>0  Then
		 Errmsg=ErrMsg + "<BR><li>VIP论坛积分只能是整数。"
		dvbbs_error()
	ElseIf  Len(Trim(vipintegral))>6 Then 
		Errmsg=ErrMsg + "<BR><li>VIP论坛魔力水晶球最多只能设置6位数字。"
		dvbbs_error()
	ElseIf Not InStr("500,600,800,900,1000,1200,1600,2000,2500,2600,2800,2900,3000,3200,3500,3600,4000,4500,4800,5000,6000,6600,8000,8800,9000,9600,9800,9900,10000",vipintegral)>0 Then
		Errmsg=ErrMsg + "<BR><li>错误的VIP论坛魔力水晶球数值。"
		dvbbs_error()
	End If
	If  Not IsNumeric(vipusetime) or InStr(vipusetime,".")>0 Then
		Errmsg=ErrMsg + "<BR><li>VIP使用时间必须是大于零的整数."
		dvbbs_error()
	ElseIf vipusetime=0 Then 
		Errmsg=ErrMsg + "<BR><li>VIP使用时间不能等于0。"
		dvbbs_error()
	End If
	If len(Dvbbs.BoardType)>128 Then
		Errmsg=ErrMsg + "<BR><li>VIP您要申请的VIP论坛名称过长，不能大于128字节(中文字64个),请修改论坛名称后再试。"
		dvbbs_error()
	End If
	Dim Rs,SQL
	SQL="select D_ForumID from [Dv_ChallengeInfo]"
	Set RS=Dvbbs.Execute(SQL)
	If Rs.eof Then
		Errmsg=ErrMsg + "<BR><li>您的阳光论坛信息不存在，请重新确认资料"
		dvbbs_error()
	End If
	If Not IsNumeric(Rs(0)) Or Len(Rs(0))>20 Then
		Errmsg=ErrMsg + "<BR><li>您的阳光论坛ID不合法，如果您未安装，请安装，如果已安装，请重新确认资料."
		dvbbs_error()
	End If
	vipintegral=CLng(vipintegral)
	vipusetime=CLng(vipusetime)
	Get_ChallengeWord()
	session("challengeWord_key")=md5(Session("challengeWord") & ":" & Dvbbs.CacheData(21,0),32)

	Response.Write "数据校验完成，正在向阳光服务器提交数据，请稍候……"
	Response.Write "<form name=""redir"" action=""http://bbs.ray5198.com/rayvipforum_magicgarden/vipforum/vipapply.jsp"" method=""post"">"
	Response.Write "<INPUT type=""hidden"" name=""forumid"" value="""&Rs(0)&""">"
	Response.Write "<INPUT type=""hidden"" name=""vipid"" value="""&vipid&""">"
	Response.Write "<INPUT type=""hidden"" name=""vipshow"" value="""&vipshow&""">"
	Response.Write "<INPUT type=""hidden"" name=""vipname"" value="""&Dvbbs.BoardType&""">"
	Response.Write "<INPUT type=""hidden"" name=""vipintegral"" value="""&vipintegral/100&""">"
	Response.Write "<INPUT type=""hidden"" name=""vipusetime"" value="""&vipusetime&""">"
	Response.Write "<INPUT type=""hidden"" name=""challengword"" value="""&Session("challengeWord")&""">"
	Response.Write "<INPUT type=""hidden"" name=""viptransurl"" value="""&Dvbbs.Get_ScriptNameUrl()&""">"
	Response.Write "<INPUT type=""hidden"" name=""viptranspage"" value=""admin_vipboard.asp?action=result&Boardid="&Dvbbs.BoardID&""">"
	Response.Write "</form>"
	Response.Write Chr(10)
	Response.Write "<script language=""javascript"">redir.submit();</script>"
	Response.Write Chr(10)
	Set Rs = Nothing
End Sub 
'激活VIP论坛
Sub Reg2()
	If Dvbbs.BoardID=0 Then 
		Errmsg=ErrMsg + "<BR><li>错误的板面参数。"
		dvbbs_error()
	End If
	'校验表单数据
	Dim vipid
	Dim NowStr
	Dim trs
	NowStr=datediff("s","1970-1-1",Now)
	vipid=Dvbbs.BoardID
	vipid=cstr(NowStr) & cstr(vipid)
	vipid=md5(vipid,32)
	Session("NowStr")=NowStr

	Dim Rs,SQL
	SQL="select D_ForumID from [Dv_ChallengeInfo]"
	Set RS=Dvbbs.Execute(SQL)
	If Rs.eof Then
		Errmsg=ErrMsg + "<BR><li>您的阳光论坛信息不存在，请重新确认资料"
		dvbbs_error()
	End If
	If Not IsNumeric(Rs(0)) Or Len(Rs(0))>20 Then
		Errmsg=ErrMsg + "<BR><li>您的阳光论坛ID不合法，如果您未安装，请安装，如果已安装，请重新确认资料."
		dvbbs_error()
	End If
	vipintegral=CLng(vipintegral)
	vipusetime=CLng(vipusetime)
	Get_ChallengeWord()
	session("challengeWord_key")=md5(Session("challengeWord") & ":" & Dvbbs.CacheData(21,0),32)

	dim oldvipid,tboard_setting
	set trs=Dvbbs.Execute("select board_setting from dv_board where boardid="&dvbbs.boardid)
	tboard_setting=split(rs(0),",")
	oldvipid=cstr(tboard_setting(47)) & cstr(dvbbs.boardid)
	oldvipid=md5(oldvipid,32)


	Response.Write "数据校验完成，正在向阳光服务器提交数据，请稍候……"
	Response.Write "<form name=""redir"" action=""http://bbs.ray5198.com/rayvipforum_magicgarden/vipforum/vipapply_update.jsp"" method=""post"">"
	Response.Write "<INPUT type=""hidden"" name=""forumid"" value="""&Rs("D_challengePassWord")&""">"
	Response.Write "<INPUT type=""hidden"" name=""vipid"" value="""&oldvipid&""">"
	Response.Write "<INPUT type=""hidden"" name=""newforumid"" value="""&Rs(0)&""">"
	Response.Write "<INPUT type=""hidden"" name=""newvipid"" value="""&vipid&""">"
	Response.Write "<INPUT type=""hidden"" name=""challengword"" value="""&Session("challengeWord")&""">"
	Response.Write "<INPUT type=""hidden"" name=""viptransurl"" value="""&Dvbbs.Get_ScriptNameUrl()&""">"
	Response.Write "<INPUT type=""hidden"" name=""viptranspage"" value=""admin_vipboard.asp?action=result1&Boardid="&Dvbbs.BoardID&""">"
	Response.Write "</form>"
	Response.Write Chr(10)
	Response.Write "<script language=""javascript"">redir.submit();</script>"
	Response.Write Chr(10)
	Set Rs = Nothing
	Set Trs=nothing
End Sub 
Sub Main()
	If Dvbbs.BoardID=0 Then 
		Errmsg=ErrMsg + "<BR><li>错误的板面参数。"
		dvbbs_error()
	End If
	'Response.Write "这部分要添加一些关于VIP论坛的介绍，让人了解。"
	Response.Write "<br>"
	Response.Write "<form action=""?action=reg1"" method=post name=dvform>"
	Response.Write "<table cellspacing=""1"" cellpadding=""0"" border=""0"" align=""center"" width="""&Dvbbs.mainsetting(0)&""" class=""forumRowHighlight"">"
	Response.Write "<tr><th height=""25"" colspan=""2"">"
	Response.Write "VIP论坛申请"
	Response.Write "</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write "<b>您要申请的版面是：</b>"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"
	Response.Write Dvbbs.BoardType 
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write "<b>您的VIP论坛ID：</b>"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"
	Response.Write Dvbbs.BoardID
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" valign=""top"" align=""right""  width=""30%"">"
	Response.Write "<br><b>VIP申请说明：</b><p align=""left"">请仔细填写，管理员将根据您站点的信息以及提交的申请信息来确定是否通过该申请。</p>"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight""align=""left""  width=""70%"" valign=""top"" >"
	Response.Write "<textarea name=""vipshow"" cols=""36"" rows=""10""></textarea>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write "<b>VIP论坛魔力水晶球 ：</b>"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%""><select name=""vipintegral"" size=1>"
	Dim PointList
	PointList="500,600,800,900,1000,1200,1600,2000,2500,2600,2800,2900,3000,3200,3500,3600,4000,4500,4800,5000,6000,6600,8000,8800,9000,9600,9800,9900,10000"
	PointList=split(PointList,",")
	For i=0 To Ubound(PointList)
		Response.Write "<option value="""&PointList(i)&""">"&PointList(i)/100&"</option>"
	Next
	'Response.Write "<input name=""vipintegral"" type=""text"" value=""3000"" size=""6"" maxlength=""6""> "
	Response.Write "</select>用户登录VIP论坛若干天（和下面的天数对应）所需<b>魔力水晶球</b>数值。</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write "<b>VIP使用时间：</b>"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%""><select name=""vipusetime"" size=1>"
	'For i=1 To 12
	i=1
	Response.Write "<option value="""&i*30&""">"&i*30&"</option>"
	'Next
	'Response.Write "<input name=""vipusetime"" type=""text"" value=""30"" size=""6"">"
	Response.Write "</select> 天&nbsp;&nbsp;*请选择相应天数，此为用户支付相应魔力水晶球所能使用VIP论坛的天数。</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write ""
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"
	Response.Write "<input type=checkbox name=""iapply"" value=""on"">&nbsp;我同意《<a href=# onclick=""javascript:onclick_apply()"">VIP论坛申请须知</a>》中的全部条款"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write "<input name=""Boardid"" type=""hidden"" value="""&Dvbbs.BoardID&""">"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"
	Response.Write ""
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write "<input type=""reset"" name=""Submit"" value=""重 填"">"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" name=""B1"" value=""下一步"" Onclick=""submitclick();"">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr><td height=""35"" colspan=""2"" Class=""forumHeaderBackgroundAlternate"">"
	Response.Write "</td></tr>"
	Response.Write "</table>"
	Response.Write "</form>"

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function submitclick()
{
	if(dvform.iapply.checked==false)
	{
		alert('申请VIP论坛，请仔细阅读VIP论坛申请须知，并在我同意《VIP论坛申请须知》中的全部条款前打勾。'); 
	}
	else
	{
		dvform.submit()
	}
}
function onclick_apply()
{
	alert('申请VIP论坛须知\n\n　　VIP论坛是阳光论坛系列软件为所有论坛站长提供的重要互动服务，由此带来的相关服务收益，将由阳光加信和论坛站长分享。\n\n　　为了感谢广大阳光会员和阳光论坛站长对阳光加信的长期支持，北京阳光加信科技有限公司近日推出了多种回馈奖励阳光魔力水晶球活动。用户可以通过使用阳光加信魔力花园中的各项服务，不但可以获得免费信息服务、手机、电脑等奖品，更可以进入各大论坛的VIP论坛享受阳光会员所专用的VIP服务。\n\n　　阳光论坛站长可以通过在论坛中开设各类VIP论坛享受自已的收益。每100魔力水晶球相当于人民币1元，初始收益比例为25%。如果您当月的收益超过2000元，便可在当月按照超级论坛进行结算，超出部分享受30%的高比例收益。\n\n使用方式：\n　　1.请在论坛的管理中心建立VIP论坛。在建立的过程中，请认真填写各项信息。其中VIP论坛的名称和VIP论坛的使用说明两项请务必详细填写。您的VIP论坛是否申请成功很可能就取决与此；魔力水晶球的个数，表明最终用户要进入此VIP论坛需要使用的魔力水晶球的个数。用户使用一次水晶球，将可以拥有访问VIP论坛的权限30天。\n　　2.阳光加信的工作人员将在48小时内审批您的VIP论坛是否能够开通。如果申请被通过的话，您将会收到阳光加信给您发送的成功邮件，同时您的VIP论坛即时就可以正常使用。\n\n注意：\n　　您必须遵守《全国人大常委会关于维护互联网安全的决定》及中华人民共和国其他各项法律法规，确保申请建立的VIP论坛中不含有任何境内境外色情或反动内容。否则，阳光加信有权立即关闭此VIP论坛，并没收您在使用阳光论坛系列软件过程中所取得所有收益。同时，您将承担由此产生的一切后果。\n\n　　您使用阳光论坛系列软件中的VIP论坛功能， 即表示您已经阅读并接受如上所有条款以及阳光论坛系列软件站长须知中的所有条款。'); 
}
//-->
</SCRIPT>
<%
End Sub 
Sub result()
	Dim Sv
	Response.Write "<br>"
	Response.Write "<table cellspacing=""1"" cellpadding=""0"" border=""0"" align=""center"" width="""&Dvbbs.mainsetting(0)&""" class=""forumRowHighlight"">"
	Response.Write "<tr><th height=""25"" colspan=""2"">"
	Response.Write "VIP论坛申请返回信息"
	Response.Write "</td></tr>"
	'Response.Write "<tr>"
	'For Each Sv in Request.form
	'	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	'	Response.Write sv & "</td><td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"&Request.form(Sv)
	'	Response.Write "</td></tr>"
	'Next
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""left""  width=""100%"" colspan=""2"">"
	Dim errcode,errmessage,retokerWord,challengword,challengeWord_key
	errcode=Trim(Request("errcode"))
	errmessage=Trim(Request("errmessage"))
	retokerWord=Trim(Request("tokenword"))
	challengword=Trim(Request("challengword"))
	Dim vipboardsetting,vipboardslist,vipintegral,vipusetime
	vipusetime = Session("vipusetime")
	vipintegral=Session("vipintegral")
	Select Case errcode
		Case 1000
			challengeWord_key=session("challengeWord_key")
			If challengeWord_key=retokerWord Then
				'校验成功，更新版面数据
				Set Rs=Dvbbs.Execute("Select Board_Setting from dv_Board where BoardID="&Dvbbs.BoardID)
				vipboardsetting=split(rs("Board_Setting"),",")
				For i=0 to UBound(vipboardsetting)
					If i=0 Then
						vipboardslist=vipboardsetting(i)
					Else
						'认证论坛
						If i=2 Then
							vipboardslist=vipboardslist & ",1"
							'积分
						ElseIf i=20 Then
							vipboardslist=vipboardslist & "," & vipintegral
							'时间
						ElseIf i=46 Then
							vipboardslist=vipboardslist & "," & vipusetime
							'标识
						Elseif i=47 Then
							vipboardslist=vipboardslist & "," & Session("NowStr")
						Else
							vipboardslist=vipboardslist & "," & vipboardsetting(i)
						End If
					End If
				Next
				Set Rs = Nothing 
				Dvbbs.Execute("update dv_Board Set Board_Setting='"&vipboardslist&"' Where BoardId="&Dvbbs.BoardID)
				Dvbbs.ReloadBoardInfo(Dvbbs.BoardID)
				Response.Write "<li>您的VIP论坛已经成功提交申请。"
			Else
				Errmsg=ErrMsg + "<BR><li>非法的提交过程，"&errmessage&"。"
				dvbbs_error()
			End If
		Case 1001
			Errmsg=ErrMsg + "<BR><li>VIP没有审核通过，"&errmessage&"。"
			dvbbs_error()
		Case 1002
			Errmsg=ErrMsg + "<BR><li>用户积分不够，"&errmessage&"。"
			dvbbs_error()
		Case 1004
			Errmsg=ErrMsg + "<BR><li>论坛数据不合法，"&errmessage&"。"
			dvbbs_error()
		Case 1005
			Errmsg=ErrMsg + "<BR><li>论坛ID不存在，"&errmessage&"。"
			dvbbs_error()
		Case 1007
			Errmsg=ErrMsg + "<BR><li>VIP论坛已经申请正处于使用状态，"&errmessage&"。"
			dvbbs_error()
		Case 1008
			Errmsg=ErrMsg + "<BR><li>不是有效的论坛，"&errmessage&"。"
			dvbbs_error()
		Case 1009
			Errmsg=ErrMsg + "<BR><li>VIP论坛申请失败，"&errmessage&"。"
			dvbbs_error()
		Case 1010
			Errmsg=ErrMsg + "<BR><li>数据操作失败，"&errmessage&"。"
			dvbbs_error()
		Case 1011
			Errmsg=ErrMsg + "<BR><li>不明原因与管理员联系，"&errmessage&"。"
			dvbbs_error()
		Case 1012
			Errmsg=ErrMsg + "<BR><li>积分超过上限，"&errmessage&"。"
			dvbbs_error()
		Case 1013
			Errmsg=ErrMsg + "<BR><li>提供的挑战随机数是空，"&errmessage&"。"
			dvbbs_error()
		Case 1014
			Set Rs=Dvbbs.Execute("Select Board_Setting from dv_Board where BoardID="&Dvbbs.BoardID)
				vipboardsetting=rs("Board_Setting")
				vipboardsetting=split(vipboardsetting,",")
				For i=0 to UBound(vipboardsetting)
					If i=0 Then
						vipboardslist=vipboardsetting(i)
					Else
						'认证论坛
						If i=2 Then
							vipboardslist=vipboardslist & ",0"
							'积分
						ElseIf i=20 Then
							vipboardslist=vipboardslist & "," & vipintegral
							'时间
						ElseIf i=46 Then
							vipboardslist=vipboardslist & "," & vipusetime
						Else
							vipboardslist=vipboardslist & "," & vipboardsetting(i)
						End If
					End If
				Next
				Set rs=Nothing  
				Dvbbs.Execute("update dv_Board Set Board_Setting='"&vipboardslist&"' Where BoardId="&Dvbbs.BoardID)
				Dvbbs.ReloadBoardInfo(Dvbbs.BoardID)
			Errmsg=ErrMsg + "<BR><li>你已经申请过VIP论坛,但是VIP论坛没有被激活，请等待。"
			dvbbs_error()
		Case Else
			Errmsg=ErrMsg + "<BR><li>非法的提交过程，"&errmessage&"。"
			dvbbs_error()
	End Select
	Response.Write "</td></tr>"
	Response.Write "</table>"
			
End Sub 
Sub result1()
	Dim Sv
	Response.Write "<br>"
	Response.Write "<table cellspacing=""1"" cellpadding=""0"" border=""0"" align=""center"" width="""&Dvbbs.mainsetting(0)&""" class=""forumRowHighlight"">"
	Response.Write "<tr><th height=""25"" colspan=""2"">"
	Response.Write "VIP论坛申请返回信息"
	Response.Write "</td></tr>"
	'Response.Write "<tr>"
	'For Each Sv in Request.form
	'	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	'	Response.Write sv & "</td><td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"&Request.form(Sv)
	'	Response.Write "</td></tr>"
	'Next
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""left""  width=""100%"" colspan=""2"">"
	Dim errcode,errmessage,retokerWord,challengword,challengeWord_key
	errcode=Trim(Request("errcode"))
	errmessage=Trim(Request("errmessage"))
	retokerWord=Trim(Request("tokenword"))
	challengword=Trim(Request("challengword"))
	Dim vipboardsetting,vipboardslist,vipintegral,vipusetime
	vipusetime = Session("vipusetime")
	vipintegral=Session("vipintegral")
	Select Case errcode
		Case 1000
			challengeWord_key=session("challengeWord_key")
			If challengeWord_key=retokerWord Then
				'校验成功，更新版面数据
				Set Rs=Dvbbs.Execute("Select Board_Setting from dv_Board where BoardID="&Dvbbs.BoardID)
				vipboardsetting=split(rs("Board_Setting"),",")
				For i=0 to UBound(vipboardsetting)
					If i=0 Then
						vipboardslist=vipboardsetting(i)
					Else
						'认证论坛
						If i=2 Then
							vipboardslist=vipboardslist & ",1"
							'标识
						Elseif i=47 Then
							vipboardslist=vipboardslist & "," & Session("NowStr")
						Else
							vipboardslist=vipboardslist & "," & vipboardsetting(i)
						End If
					End If
				Next
				Set Rs = Nothing 
				Dvbbs.Execute("update dv_Board Set Board_Setting='"&vipboardslist&"' Where BoardId="&Dvbbs.BoardID)
				Dvbbs.ReloadBoardInfo(Dvbbs.BoardID)
				Response.Write "<li>您的VIP论坛已经成功提交申请。"
			Else
				Errmsg=ErrMsg + "<BR><li>非法的提交过程。"
				dvbbs_error()
			End If
		Case 1001
			Errmsg=ErrMsg + "<BR><li>VIP没有审核通过。"
			dvbbs_error()
		Case 1002
			Errmsg=ErrMsg + "<BR><li>用户积分不够。"
			dvbbs_error()
		Case 1004
			Errmsg=ErrMsg + "<BR><li>论坛数据不合法"
			dvbbs_error()
		Case 1005
			Errmsg=ErrMsg + "<BR><li>论坛ID不存在"
			dvbbs_error()
		Case 1007
			Errmsg=ErrMsg + "<BR><li>VIP论坛已经申请正处于使用状态"
			dvbbs_error()
		Case 1008
			Errmsg=ErrMsg + "<BR><li>不是有效的论坛"
			dvbbs_error()
		Case 1009
			Errmsg=ErrMsg + "<BR><li>VIP论坛申请失败"
			dvbbs_error()
		Case 1010
			Errmsg=ErrMsg + "<BR><li>数据操作失败"
			dvbbs_error()
		Case 1011
			Errmsg=ErrMsg + "<BR><li>不明原因与管理员联系"
			dvbbs_error()
		Case 1012
			Errmsg=ErrMsg + "<BR><li>积分超过上限"
			dvbbs_error()
		Case 1013
			Errmsg=ErrMsg + "<BR><li>提供的挑战随机数是空"
			dvbbs_error()
		Case 1014
			Set Rs=Dvbbs.Execute("Select Board_Setting from dv_Board where BoardID="&Dvbbs.BoardID)
				vipboardsetting=rs("Board_Setting")
				vipboardsetting=split(vipboardsetting,",")
				For i=0 to UBound(vipboardsetting)
					If i=0 Then
						vipboardslist=vipboardsetting(i)
					Else
						'认证论坛
						If i=2 Then
							vipboardslist=vipboardslist & ",0"
							'积分
						ElseIf i=20 Then
							vipboardslist=vipboardslist & "," & vipintegral
							'时间
						ElseIf i=46 Then
							vipboardslist=vipboardslist & "," & vipusetime
						Else
							vipboardslist=vipboardslist & "," & vipboardsetting(i)
						End If
					End If
				Next
				Set rs=Nothing  
				Dvbbs.Execute("update dv_Board Set Board_Setting='"&vipboardslist&"' Where BoardId="&Dvbbs.BoardID)
				Dvbbs.ReloadBoardInfo(Dvbbs.BoardID)
			Errmsg=ErrMsg + "<BR><li>你已经申请过VIP论坛,但是VIP论坛没有被激活，请等待。"
			dvbbs_error()
		Case Else
			Errmsg=ErrMsg + "<BR><li>非法的提交过程。"
			dvbbs_error()
	End Select
	Response.Write "</td></tr>"
	Response.Write "</table>"
			
End Sub 
Sub showvipuser()'查看版面VIP用户
	If Dvbbs.BoardID=0 Then 
		Errmsg=ErrMsg + "<BR><li>错误的板面参数。"
		dvbbs_error()
	End If
	Dim Rs,SQL
	dim currentPage,page_count,totalrec,Pcount,endpage,i
	currentPage=request("page")
	If currentpage="" or Not IsNumeric(currentpage) Then
		currentpage=1
	Else
		currentpage=CLng(currentpage)
	End If
	Set Rs=server.createobject("adodb.recordset")
	SQL="Select *  from [DV_ChanOrders] where O_type=2 and O_BoardID="&Dvbbs.BoardID&" and O_issuc=1"
	Rs.Open SQL,Conn,1,1
	Response.Write "<br>"
	Response.Write "<table cellspacing=""1"" cellpadding=""0"" border=""0"" align=""center"" width="""&Dvbbs.mainsetting(0)&""" class=""forumRowHighlight"">"
	Response.Write "<tr><th height=""25"" colspan=""2"">"
	Response.Write "VIP论坛:"&Dvbbs.BoardType&"用户列表"
	Response.Write "</td></tr>"
	If Not (Rs.eof And Rs.BOF) Then
		rs.PageSize = Dvbbs.Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=Rs.RecordCount 
		Response.Write "<tr>"
		Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""center""  width=""30%"">"
		Response.Write "<b>用户名</b>"
		Response.Write "</td>"
		Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""30%"">"
		Response.Write "<b>所支付的魔力晶球</b>"
		Response.Write "</td>"
		Response.Write "</tr>"
		Do While Not Rs.EOF
			Response.Write "<tr>"
			Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""center""  width=""30%"">"
			Response.Write Rs("O_Username")
			Response.Write "</td>"
			Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""30%"">"
			Response.Write Rs("O_PayMoney")
			Response.Write "</td>"
			Response.Write "</tr>"
			Rs.MoveNext
		Loop
		If totalrec mod Dvbbs.Forum_Setting(11)=0 Then
			Pcount= totalrec \ Dvbbs.Forum_Setting(11)
		Else
			Pcount= totalrec \ Dvbbs.Forum_Setting(11)+1
		End If
		Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""left"">"
		Response.Write "页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"
		Response.Write "&nbsp;每页<b>"&Dvbbs.Forum_Setting(11)&"</b> 总数<b>"&totalrec&"</b></td>"
		Response.Write "<td Class=""forumRowHighlight"" valign=middle nowrap align=right>分页："
		If currentpage > 4 Then
			Response.Write "<a href=""?page=1&action="&request("action")&""">[1]</a> ..."
		End If
		If Pcount>currentpage+3 Then
			endpage=currentpage+3
		Else
			endpage=Pcount
		End If
		for i=currentpage-3 to endpage
			If Not i<1 Then
				If i = CLng(currentpage) Then
					Response.Write " <font color="&Dvbbs.mainsetting(1)&">["&i&"]</font>"
				Else
					Response.Write " <a href=""?page="&i&"&action="&request("action")&""">["&i&"]</a>"
				End If
			End If
		Next
		If currentpage+3 < Pcount Then
			Response.Write "... <a href=""?page="&Pcount&"&action="&request("action")&""">["&Pcount&"]</a>"
		End If
		Response.Write "</td></tr>"	
	Else
		Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" colspan=""2"" align=""center"">"
		Response.Write "本论坛尚无VIP用户。"
		Response.Write "</td></tr>"		
	End If
	Response.Write "<tr><td class=""forumHeaderBackgroundAlternate"" height=""25"" colspan=""2"" align=""center"">"
	Response.Write "</td></tr>"		
	Response.Write "</table>"
End Sub
Footer
%>