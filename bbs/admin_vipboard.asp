<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!--#include file="inc/md5.asp"-->
<!-- #include file="inc/DvADChar.asp" -->
<!--#include file="inc/chan_const.asp"-->
<%
Head()

If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(8)=1) Then
	Errmsg=Errmsg+"<br>"+"<li>����̳û�п���VIP�շ���̳���ܡ�"
	call dvbbs_error()
Else
	If Not dvbbs.master or instr(","&session("flag")&",",",9,")=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
		Errmsg=ErrMsg + "<BR><li>����İ��������"
		dvbbs_error()
	End If
	'У�������
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
		Errmsg=ErrMsg + "<BR><li>����дVIP����˵����"
		dvbbs_error()
	ElseIf len(vipshow) > 1024 Then 
		Errmsg=ErrMsg + "<BR><li>VIP����˵���ĳ��Ȳ��ܴ���1024�ֽڡ�"
		dvbbs_error()
	End If
	If Not IsNumeric(vipintegral) or InStr(vipintegral,".")>0  Then
		 Errmsg=ErrMsg + "<BR><li>VIP��̳����ֻ����������"
		dvbbs_error()
	ElseIf  Len(Trim(vipintegral))>6 Then 
		Errmsg=ErrMsg + "<BR><li>VIP��̳ħ��ˮ�������ֻ������6λ���֡�"
		dvbbs_error()
	ElseIf Not InStr("500,600,800,900,1000,1200,1600,2000,2500,2600,2800,2900,3000,3200,3500,3600,4000,4500,4800,5000,6000,6600,8000,8800,9000,9600,9800,9900,10000",vipintegral)>0 Then
		Errmsg=ErrMsg + "<BR><li>�����VIP��̳ħ��ˮ������ֵ��"
		dvbbs_error()
	End If
	If  Not IsNumeric(vipusetime) or InStr(vipusetime,".")>0 Then
		Errmsg=ErrMsg + "<BR><li>VIPʹ��ʱ������Ǵ����������."
		dvbbs_error()
	ElseIf vipusetime=0 Then 
		Errmsg=ErrMsg + "<BR><li>VIPʹ��ʱ�䲻�ܵ���0��"
		dvbbs_error()
	End If
	If len(Dvbbs.BoardType)>128 Then
		Errmsg=ErrMsg + "<BR><li>VIP��Ҫ�����VIP��̳���ƹ��������ܴ���128�ֽ�(������64��),���޸���̳���ƺ����ԡ�"
		dvbbs_error()
	End If
	Dim Rs,SQL
	SQL="select D_ForumID from [Dv_ChallengeInfo]"
	Set RS=Dvbbs.Execute(SQL)
	If Rs.eof Then
		Errmsg=ErrMsg + "<BR><li>����������̳��Ϣ�����ڣ�������ȷ������"
		dvbbs_error()
	End If
	If Not IsNumeric(Rs(0)) Or Len(Rs(0))>20 Then
		Errmsg=ErrMsg + "<BR><li>����������̳ID���Ϸ��������δ��װ���밲װ������Ѱ�װ��������ȷ������."
		dvbbs_error()
	End If
	vipintegral=CLng(vipintegral)
	vipusetime=CLng(vipusetime)
	Get_ChallengeWord()
	session("challengeWord_key")=md5(Session("challengeWord") & ":" & Dvbbs.CacheData(21,0),32)

	Response.Write "����У����ɣ�����������������ύ���ݣ����Ժ򡭡�"
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
'����VIP��̳
Sub Reg2()
	If Dvbbs.BoardID=0 Then 
		Errmsg=ErrMsg + "<BR><li>����İ��������"
		dvbbs_error()
	End If
	'У�������
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
		Errmsg=ErrMsg + "<BR><li>����������̳��Ϣ�����ڣ�������ȷ������"
		dvbbs_error()
	End If
	If Not IsNumeric(Rs(0)) Or Len(Rs(0))>20 Then
		Errmsg=ErrMsg + "<BR><li>����������̳ID���Ϸ��������δ��װ���밲װ������Ѱ�װ��������ȷ������."
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


	Response.Write "����У����ɣ�����������������ύ���ݣ����Ժ򡭡�"
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
		Errmsg=ErrMsg + "<BR><li>����İ��������"
		dvbbs_error()
	End If
	'Response.Write "�ⲿ��Ҫ���һЩ����VIP��̳�Ľ��ܣ������˽⡣"
	Response.Write "<br>"
	Response.Write "<form action=""?action=reg1"" method=post name=dvform>"
	Response.Write "<table cellspacing=""1"" cellpadding=""0"" border=""0"" align=""center"" width="""&Dvbbs.mainsetting(0)&""" class=""forumRowHighlight"">"
	Response.Write "<tr><th height=""25"" colspan=""2"">"
	Response.Write "VIP��̳����"
	Response.Write "</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write "<b>��Ҫ����İ����ǣ�</b>"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"
	Response.Write Dvbbs.BoardType 
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write "<b>����VIP��̳ID��</b>"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"
	Response.Write Dvbbs.BoardID
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" valign=""top"" align=""right""  width=""30%"">"
	Response.Write "<br><b>VIP����˵����</b><p align=""left"">����ϸ��д������Ա��������վ�����Ϣ�Լ��ύ��������Ϣ��ȷ���Ƿ�ͨ�������롣</p>"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight""align=""left""  width=""70%"" valign=""top"" >"
	Response.Write "<textarea name=""vipshow"" cols=""36"" rows=""10""></textarea>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write "<b>VIP��̳ħ��ˮ���� ��</b>"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%""><select name=""vipintegral"" size=1>"
	Dim PointList
	PointList="500,600,800,900,1000,1200,1600,2000,2500,2600,2800,2900,3000,3200,3500,3600,4000,4500,4800,5000,6000,6600,8000,8800,9000,9600,9800,9900,10000"
	PointList=split(PointList,",")
	For i=0 To Ubound(PointList)
		Response.Write "<option value="""&PointList(i)&""">"&PointList(i)/100&"</option>"
	Next
	'Response.Write "<input name=""vipintegral"" type=""text"" value=""3000"" size=""6"" maxlength=""6""> "
	Response.Write "</select>�û���¼VIP��̳�����죨�������������Ӧ������<b>ħ��ˮ����</b>��ֵ��</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write "<b>VIPʹ��ʱ�䣺</b>"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%""><select name=""vipusetime"" size=1>"
	'For i=1 To 12
	i=1
	Response.Write "<option value="""&i*30&""">"&i*30&"</option>"
	'Next
	'Response.Write "<input name=""vipusetime"" type=""text"" value=""30"" size=""6"">"
	Response.Write "</select> ��&nbsp;&nbsp;*��ѡ����Ӧ��������Ϊ�û�֧����Ӧħ��ˮ��������ʹ��VIP��̳��������</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""right""  width=""30%"">"
	Response.Write ""
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"
	Response.Write "<input type=checkbox name=""iapply"" value=""on"">&nbsp;��ͬ�⡶<a href=# onclick=""javascript:onclick_apply()"">VIP��̳������֪</a>���е�ȫ������"
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
	Response.Write "<input type=""reset"" name=""Submit"" value=""�� ��"">"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""70%"">"
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" name=""B1"" value=""��һ��"" Onclick=""submitclick();"">"
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
		alert('����VIP��̳������ϸ�Ķ�VIP��̳������֪��������ͬ�⡶VIP��̳������֪���е�ȫ������ǰ�򹴡�'); 
	}
	else
	{
		dvform.submit()
	}
}
function onclick_apply()
{
	alert('����VIP��̳��֪\n\n����VIP��̳��������̳ϵ�����Ϊ������̳վ���ṩ����Ҫ���������ɴ˴�������ط������棬����������ź���̳վ������\n\n����Ϊ�˸�л��������Ա��������̳վ����������ŵĳ���֧�֣�����������ſƼ����޹�˾�����Ƴ��˶��ֻ�����������ħ��ˮ�������û�����ͨ��ʹ���������ħ����԰�еĸ�����񣬲������Ի�������Ϣ�����ֻ������ԵȽ�Ʒ�������Խ��������̳��VIP��̳���������Ա��ר�õ�VIP����\n\n����������̳վ������ͨ������̳�п������VIP��̳�������ѵ����档ÿ100ħ��ˮ�����൱�������1Ԫ����ʼ�������Ϊ25%����������µ����泬��2000Ԫ������ڵ��°��ճ�����̳���н��㣬������������30%�ĸ߱������档\n\nʹ�÷�ʽ��\n����1.������̳�Ĺ������Ľ���VIP��̳���ڽ����Ĺ����У���������д������Ϣ������VIP��̳�����ƺ�VIP��̳��ʹ��˵�������������ϸ��д������VIP��̳�Ƿ�����ɹ��ܿ��ܾ�ȡ����ˣ�ħ��ˮ����ĸ��������������û�Ҫ�����VIP��̳��Ҫʹ�õ�ħ��ˮ����ĸ������û�ʹ��һ��ˮ���򣬽�����ӵ�з���VIP��̳��Ȩ��30�졣\n����2.������ŵĹ�����Ա����48Сʱ����������VIP��̳�Ƿ��ܹ���ͨ��������뱻ͨ���Ļ����������յ�������Ÿ������͵ĳɹ��ʼ���ͬʱ����VIP��̳��ʱ�Ϳ�������ʹ�á�\n\nע�⣺\n�������������ء�ȫ���˴�ί�����ά����������ȫ�ľ��������л����񹲺͹���������ɷ��棬ȷ�����뽨����VIP��̳�в������κξ��ھ���ɫ��򷴶����ݡ��������������Ȩ�����رմ�VIP��̳����û������ʹ��������̳ϵ�������������ȡ���������档ͬʱ�������е��ɴ˲�����һ�к����\n\n������ʹ��������̳ϵ������е�VIP��̳���ܣ� ����ʾ���Ѿ��Ķ��������������������Լ�������̳ϵ�����վ����֪�е��������'); 
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
	Response.Write "VIP��̳���뷵����Ϣ"
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
				'У��ɹ������°�������
				Set Rs=Dvbbs.Execute("Select Board_Setting from dv_Board where BoardID="&Dvbbs.BoardID)
				vipboardsetting=split(rs("Board_Setting"),",")
				For i=0 to UBound(vipboardsetting)
					If i=0 Then
						vipboardslist=vipboardsetting(i)
					Else
						'��֤��̳
						If i=2 Then
							vipboardslist=vipboardslist & ",1"
							'����
						ElseIf i=20 Then
							vipboardslist=vipboardslist & "," & vipintegral
							'ʱ��
						ElseIf i=46 Then
							vipboardslist=vipboardslist & "," & vipusetime
							'��ʶ
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
				Response.Write "<li>����VIP��̳�Ѿ��ɹ��ύ���롣"
			Else
				Errmsg=ErrMsg + "<BR><li>�Ƿ����ύ���̣�"&errmessage&"��"
				dvbbs_error()
			End If
		Case 1001
			Errmsg=ErrMsg + "<BR><li>VIPû�����ͨ����"&errmessage&"��"
			dvbbs_error()
		Case 1002
			Errmsg=ErrMsg + "<BR><li>�û����ֲ�����"&errmessage&"��"
			dvbbs_error()
		Case 1004
			Errmsg=ErrMsg + "<BR><li>��̳���ݲ��Ϸ���"&errmessage&"��"
			dvbbs_error()
		Case 1005
			Errmsg=ErrMsg + "<BR><li>��̳ID�����ڣ�"&errmessage&"��"
			dvbbs_error()
		Case 1007
			Errmsg=ErrMsg + "<BR><li>VIP��̳�Ѿ�����������ʹ��״̬��"&errmessage&"��"
			dvbbs_error()
		Case 1008
			Errmsg=ErrMsg + "<BR><li>������Ч����̳��"&errmessage&"��"
			dvbbs_error()
		Case 1009
			Errmsg=ErrMsg + "<BR><li>VIP��̳����ʧ�ܣ�"&errmessage&"��"
			dvbbs_error()
		Case 1010
			Errmsg=ErrMsg + "<BR><li>���ݲ���ʧ�ܣ�"&errmessage&"��"
			dvbbs_error()
		Case 1011
			Errmsg=ErrMsg + "<BR><li>����ԭ�������Ա��ϵ��"&errmessage&"��"
			dvbbs_error()
		Case 1012
			Errmsg=ErrMsg + "<BR><li>���ֳ������ޣ�"&errmessage&"��"
			dvbbs_error()
		Case 1013
			Errmsg=ErrMsg + "<BR><li>�ṩ����ս������ǿգ�"&errmessage&"��"
			dvbbs_error()
		Case 1014
			Set Rs=Dvbbs.Execute("Select Board_Setting from dv_Board where BoardID="&Dvbbs.BoardID)
				vipboardsetting=rs("Board_Setting")
				vipboardsetting=split(vipboardsetting,",")
				For i=0 to UBound(vipboardsetting)
					If i=0 Then
						vipboardslist=vipboardsetting(i)
					Else
						'��֤��̳
						If i=2 Then
							vipboardslist=vipboardslist & ",0"
							'����
						ElseIf i=20 Then
							vipboardslist=vipboardslist & "," & vipintegral
							'ʱ��
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
			Errmsg=ErrMsg + "<BR><li>���Ѿ������VIP��̳,����VIP��̳û�б������ȴ���"
			dvbbs_error()
		Case Else
			Errmsg=ErrMsg + "<BR><li>�Ƿ����ύ���̣�"&errmessage&"��"
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
	Response.Write "VIP��̳���뷵����Ϣ"
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
				'У��ɹ������°�������
				Set Rs=Dvbbs.Execute("Select Board_Setting from dv_Board where BoardID="&Dvbbs.BoardID)
				vipboardsetting=split(rs("Board_Setting"),",")
				For i=0 to UBound(vipboardsetting)
					If i=0 Then
						vipboardslist=vipboardsetting(i)
					Else
						'��֤��̳
						If i=2 Then
							vipboardslist=vipboardslist & ",1"
							'��ʶ
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
				Response.Write "<li>����VIP��̳�Ѿ��ɹ��ύ���롣"
			Else
				Errmsg=ErrMsg + "<BR><li>�Ƿ����ύ���̡�"
				dvbbs_error()
			End If
		Case 1001
			Errmsg=ErrMsg + "<BR><li>VIPû�����ͨ����"
			dvbbs_error()
		Case 1002
			Errmsg=ErrMsg + "<BR><li>�û����ֲ�����"
			dvbbs_error()
		Case 1004
			Errmsg=ErrMsg + "<BR><li>��̳���ݲ��Ϸ�"
			dvbbs_error()
		Case 1005
			Errmsg=ErrMsg + "<BR><li>��̳ID������"
			dvbbs_error()
		Case 1007
			Errmsg=ErrMsg + "<BR><li>VIP��̳�Ѿ�����������ʹ��״̬"
			dvbbs_error()
		Case 1008
			Errmsg=ErrMsg + "<BR><li>������Ч����̳"
			dvbbs_error()
		Case 1009
			Errmsg=ErrMsg + "<BR><li>VIP��̳����ʧ��"
			dvbbs_error()
		Case 1010
			Errmsg=ErrMsg + "<BR><li>���ݲ���ʧ��"
			dvbbs_error()
		Case 1011
			Errmsg=ErrMsg + "<BR><li>����ԭ�������Ա��ϵ"
			dvbbs_error()
		Case 1012
			Errmsg=ErrMsg + "<BR><li>���ֳ�������"
			dvbbs_error()
		Case 1013
			Errmsg=ErrMsg + "<BR><li>�ṩ����ս������ǿ�"
			dvbbs_error()
		Case 1014
			Set Rs=Dvbbs.Execute("Select Board_Setting from dv_Board where BoardID="&Dvbbs.BoardID)
				vipboardsetting=rs("Board_Setting")
				vipboardsetting=split(vipboardsetting,",")
				For i=0 to UBound(vipboardsetting)
					If i=0 Then
						vipboardslist=vipboardsetting(i)
					Else
						'��֤��̳
						If i=2 Then
							vipboardslist=vipboardslist & ",0"
							'����
						ElseIf i=20 Then
							vipboardslist=vipboardslist & "," & vipintegral
							'ʱ��
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
			Errmsg=ErrMsg + "<BR><li>���Ѿ������VIP��̳,����VIP��̳û�б������ȴ���"
			dvbbs_error()
		Case Else
			Errmsg=ErrMsg + "<BR><li>�Ƿ����ύ���̡�"
			dvbbs_error()
	End Select
	Response.Write "</td></tr>"
	Response.Write "</table>"
			
End Sub 
Sub showvipuser()'�鿴����VIP�û�
	If Dvbbs.BoardID=0 Then 
		Errmsg=ErrMsg + "<BR><li>����İ��������"
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
	Response.Write "VIP��̳:"&Dvbbs.BoardType&"�û��б�"
	Response.Write "</td></tr>"
	If Not (Rs.eof And Rs.BOF) Then
		rs.PageSize = Dvbbs.Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=Rs.RecordCount 
		Response.Write "<tr>"
		Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""center""  width=""30%"">"
		Response.Write "<b>�û���</b>"
		Response.Write "</td>"
		Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""left""  width=""30%"">"
		Response.Write "<b>��֧����ħ������</b>"
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
		Response.Write "ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"
		Response.Write "&nbsp;ÿҳ<b>"&Dvbbs.Forum_Setting(11)&"</b> ����<b>"&totalrec&"</b></td>"
		Response.Write "<td Class=""forumRowHighlight"" valign=middle nowrap align=right>��ҳ��"
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
		Response.Write "����̳����VIP�û���"
		Response.Write "</td></tr>"		
	End If
	Response.Write "<tr><td class=""forumHeaderBackgroundAlternate"" height=""25"" colspan=""2"" align=""center"">"
	Response.Write "</td></tr>"		
	Response.Write "</table>"
End Sub
Footer
%>