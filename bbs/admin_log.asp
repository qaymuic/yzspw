<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
Head()
Dim admin_flag
Dim action,actiontype
Dim sqlstr,l_type
action=request("action")
admin_flag=",3,"
If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	dvbbs_error()
Else 
Select Case action
Case "topic"
	actiontype="���ӹ�����־���������̶���صĶ����ӵ����в�����"
	sqlstr=" where l_type=3 "
	l_type=3
	main
Case "istop"
	actiontype="�̶�������־"
	sqlstr=" where l_type=4 "
	l_type=4
	main
Case "wealth"
	actiontype="�û�������־"
	sqlstr=" where l_type=5 "
	l_type=5
	main
Case "users"
	actiontype="�û�������־���������Ρ���������IP�ͽ����"
	sqlstr=" where l_type=6 "
	l_type=6
	main
Case "admin0"
	actiontype="��̨��־0"
	sqlstr=" where l_type=0 "
	l_type=0
	main
Case "admin1"
	actiontype="��̨��־1"
	sqlstr=" where l_type=1 "
	l_type=1
	main
Case "admin2"
	actiontype="��̨��־2"
	sqlstr=" where l_type=2 "
	l_type=2
	main
Case "dellog"
	batch()
Case Else
	actiontype="ȫ����־"
	sqlstr=" "
	l_type=""
	main
End Select

If founderr then call dvbbs_error()
footer()
End If 
Sub main()
'��־���ࣺ��̨һ���¼,l_type=1,��̨��Ҫ��¼,l_type=2,����һ�����:l_type=3�����ӹ̶����l_type=4,�����ͷ���l_type=5,�û����� l_type=6
Response.Write "<table width=""95%"" border=""0"" cellspacing=""0"" cellpadding=""0""  align=center class=""tableBorder"" >"
Response.Write "<tr>"
Response.Write "<th width=""100%"" colspan=""6"" class=""tableHeaderText""  height=25>��̳��־����"
Response.Write "</th>"
Response.Write "</tr>"
Response.Write "<tr>"
Response.Write "<td align=""center"" width=""100%"" colspan=""6"" class=""tableHeaderText""  height=25>��ǰ��ʾ��"
Response.Write actiontype
Response.Write "</td>"
Response.Write "</tr>"
Response.Write "<th width=""100%"" colspan=""6"" class=""tableHeaderText""  height=25 id=tabletitlelink >ѡ��鿴��"
Response.Write " <a href=""?action="">ȫ����־</a> |"
Response.Write " <a href=""?action=topic"">���ӹ���</a> |"
Response.Write " <a href=""?action=istop"">�̶�����</a> |"
Response.Write " <a href=""?action=wealth"">���Ͳ���</a> |"
Response.Write " <a href=""?action=users"">�û�����</a> |"
Response.Write " <a href=""?action=admin0"">��̨�¼�0</a> |"
Response.Write " <a href=""?action=admin1"">��̨�¼�1</a> |"
Response.Write " <a href=""?action=admin2"">��̨�¼�2</a> |"
Response.Write "</th>"
Response.Write "</tr>"
Response.Write "</table><br>"
Dim currentpage,page_count,Pcount,endpage
Dim sql,Rs,totalrec
currentPage=request("page")
If currentpage="" or not IsNumeric(currentpage) Then
	currentpage=1
Else
	currentpage=clng(currentpage)
End If
Dvbbs.Forum_Setting(11)=50
sql="select * from [dv_log] "&sqlstr&" order by l_addtime desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1

Response.Write "<table width=""95%"" border=""0"" cellspacing=""0"" cellpadding=""0""  align=center class=""tableBorder"" style=""word-break:break-all"" >"
Response.Write "<form action=admin_log.asp?action=dellog&l_type="&l_type&" method=post name=even>"
Response.Write "<tr align=left>"
Response.Write "<th height=25 width=""10%"" >"
Response.Write "����"
Response.Write "</td>"
Response.Write "<th height=25 width=""55%"" >"
Response.Write "�¼�����"
Response.Write "</td>"
Response.Write "<th height=25 width=""20%"">"
Response.Write "����ʱ��/IP"
Response.Write "</td>"
Response.Write "<th height=25 width=""10%"" >"
Response.Write "������"
Response.Write "</td>"
Response.Write "<th height=25 width=""5%"" >"
Response.Write "����"
Response.Write "</th>"
Response.Write "</tr>"
If Not(Rs.eof or Rs.bof) Then
	rs.PageSize = Dvbbs.Forum_Setting(11)
	rs.AbsolutePage=currentpage
	page_count=0
    	totalrec=rs.recordcount
	While (Not Rs.EOF) And (Not page_count = Rs.PageSize)
	Response.Write "<tr align=left>"
	Response.Write "<td class=""forumrow""  width=""10%"" >"
	Response.Write "<a href=dispuser.asp?name="
	Response.Write Dvbbs.HTMLEncode(rs("l_touser"))
	Response.Write " target=_blank>"
	Response.Write Dvbbs.HTMLEncode(rs("l_touser"))
	Response.Write "</a>"
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""55%"" >"
	Response.Write Dvbbs.HTMLEncode(Rs("l_content"))
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""20%"">"
	Response.Write rs("l_addtime")
	Response.Write "<br>"
	Response.Write Rs("l_ip")
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""10%"">"
	Response.Write "<a href=dispuser.asp?name="&Dvbbs.HTMLEncode(rs("l_username"))&" target=_blank>"&Dvbbs.HTMLEncode(rs("l_username"))&"</a>"
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""5%"">"
	If Rs("l_type")<>2 Then
		Response.Write  "<input type=checkbox name=lid value="&rs("l_id")&">"
	End If
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td height=2></td></tr>"
	
	page_count = page_count + 1
	Rs.MoveNext
	Wend
	Response.Write "<tr><td class=forumrowHighLight colspan=6>��ѡ��Ҫɾ�����¼���<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">ȫѡ <input type=submit name=act value=ɾ��  onclick=""{if(confirm('��ȷ��ִ�еĲ�����?')){this.document.even.submit();return true;}return false;}"">"
	Response.Write "��<input type=submit name=act onclick=""{if(confirm('ȷ���������վ���еļ�¼��?')){this.document.even.submit();return true;}return false;}"" value=�����־></td></tr>"
	If totalrec mod Dvbbs.Forum_Setting(11)=0 Then
		Pcount= totalrec \ Dvbbs.Forum_Setting(11)
  	Else
  		Pcount= totalrec \ Dvbbs.Forum_Setting(11)+1
  	End If
  	Response.Write "<table border=0 cellpadding=0 cellspacing=3 width="""&Dvbbs.mainsetting(0)&""" align=center>"
  	Response.Write "<tr><td valign=middle nowrap>"
	Response.Write "ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"
	Response.Write "&nbsp;ÿҳ<b>"&Dvbbs.Forum_Setting(11)&"</b> ����<b>"&totalrec&"</b></td>"
	Response.Write "<td valign=middle nowrap align=right>��ҳ��"
	If currentpage > 4 Then
		Response.Write "<a href=""?page=1&action="&action&""">[1]</a> ..."
	End If
	If Pcount>currentpage+3 Then
		endpage=currentpage+3
	Else
		endpage=Pcount
	End If
	For i=currentpage-3 to endpage
	If Not i<1 Then
		If i = clng(currentpage) Then
			response.write " <font color="&Dvbbs.mainsetting(1)&">["&i&"]</font>"
		Else
			Response.Write " <a href=""?page="&i&"&action="&action&""">["&i&"]</a>"
		End If
	End If
	Next
	If currentpage+3 < Pcount Then   
		Response.Write "... <a href=""?page="&Pcount&"&action="&action&""">["&Pcount&"]</a>"
	End If
	Response.Write "</td></tr></table>"
Else
	Response.Write "<tr align=center>"
	Response.Write "<td class=""forumrow"" width=""100%"" colspan=""6"" >"
	Response.Write "����ؼ�¼��"
	Response.Write "</td>"
	Response.Write "</tr>"
End If
Response.Write "</form>"
Response.Write "</table>"
Rs.close
Set rs=Nothing
End Sub

Sub batch()
	Dim lid
	If request("act")="ɾ��" Then
		If request.form("lid")="" Then
			DVbbs.AddErrmsg "��ָ������¼���"
		Else
			lid=replace(request.Form("lid"),"'","")
			lid=replace(lid,";","")
			lid=replace(lid,"--","")
			lid=replace(lid,")","")
		End If
	End if
	If request("act")="ɾ��" Then
		Dvbbs.Execute("delete from dv_log where Datediff(""D"",l_addtime, "&SqlNowString&") > 2 and l_id in ("&lid&")")
	ElseIf request("act")="�����־" Then
		If request("l_type")="" or IsNull(request("l_type")) Then 
			If IsSqlDataBase = 1 Then
			Dvbbs.Execute("delete from dv_log Where Datediff(D,l_addtime, "&SqlNowString&") > 2")
			else
			Dvbbs.Execute("delete from dv_log Where Datediff('D',l_addtime, "&SqlNowString&") > 2")
			end if
		Else
			If IsSqlDataBase = 1 Then
			Dvbbs.Execute("delete from dv_log where  Datediff(D,l_addtime, "&SqlNowString&") > 2 and l_type="&CInt(request("l_type"))&"")
			else
			Dvbbs.Execute("delete from dv_log where  Datediff('D',l_addtime, "&SqlNowString&") > 2 and l_type="&CInt(request("l_type"))&"")
			end if
		End If
	End If
	Dv_suc("�ɹ�ɾ����־��ע�⣺�����ڵ���־�ᱻϵͳ������")
End Sub
%>
<script language="javascript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = form.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form.chkall.checked;  
    }  
  }  
</script>