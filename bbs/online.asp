<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<%
Response.Expires=0
Dvbbs.LoadTemplates("online")
Response.Write Dvbbs.mainhtml(1)
Response.Write vbNewLine
Response.Write "<title>"
Response.Write Dvbbs.Forum_Info(0)
Response.Write "-"
Response.Write template.Strings(0)						
Response.Write "</title>"		
Response.Write vbNewLine
Response.Write template.html(0)
Response.Write vbNewLine
Response.Write "</head>"
Response.Write vbNewLine
Response.Write "<body>"
Response.Write vbNewLine
Response.Write "<script language=""javascript"">"
Response.Write vbNewLine

'���͵ȼ�ͼƬ������JS
Dim i,GroupTitlePic,TempGroupInfo
'ȡ���û���ͼ��ָ�����
GroupTitlePic=GetGroupTitlePic
GroupTitlePic=Split(GroupTitlePic,"@@@")
For i=0 to UBound(GroupTitlePic)-1
	TempGroupInfo=Split(GroupTitlePic(i),"|||")
	Response.Write "piclist["&TempGroupInfo(0)&"]='"&TempGroupInfo(1)&"';"
	Response.Write vbNewLine		
Next
'�����ַ���������JS

For i=0 to UBound(template.Strings)-1
	Response.Write "Strings[Strings.length]='"& template.Strings(i)&"';"		
Next 
If Request("action")<>"3" Then  
	Response.Write "reshow("&Dvbbs.boardid&");"
End If 
Response.Write vbNewLine
If Request("action")="1" Or Request("action")="3" Then
	Getonline()
End If 
Response.Write "</script>"
Response.Write "</body></html>"
Sub Getonline()
	Response.Write "nowbodarid="&Dvbbs.boardid&";"
	If Dvbbs.userid<>0 Then
		Response.Write "username='"&Dvbbs.MemberName&"';"		
	Else
		Response.Write "myid='"&Session(Dvbbs.CacheName & "UserID")(0)&"';"		
	End If
	If Dvbbs.master Then Response.Write "var master=1;"
	Dim Rs,SQL,page,tmpdata
	page=Request("page")
	If page="" Then page=1
	page=CLng(page)
	Dim Selectlist
	Selectlist=""
	'���������б���ʾ��¼�ͻʱ��
	If CInt(Dvbbs.forum_setting(33))=1  Then 
		Selectlist=Selectlist&",stats"	
	End If
	If CInt(Dvbbs.forum_setting(34))=1  Then 
		Selectlist=Selectlist&",startime,lastimebk"	
	End If
	
	'��ʾ������Ͳ���ϵͳ
	If CInt(Dvbbs.forum_setting(35))=1 Then 
		Selectlist=Selectlist&",browser"	
	End If
	'���������б���ʾ��Դ
	If CInt(Dvbbs.forum_setting(36))=1 Then
		Selectlist=Selectlist&",actCome"
	End If
	'���Բ鿴����IP����Դ  0���� 1����
	If (Dvbbs.master Or Dvbbs.Superboardmaster) And CInt(Dvbbs.GroupSetting(30)) =1 Then
		Selectlist=Selectlist&",IP"
	End If
	If Not IsObject(Conn) Then ConnectionDatabase
	Set Rs = Server.CreateObject("adodb.recordset")
	If Dvbbs.boardid=0 Then
		SQL="Select id,username,UserGroupID,userhidden,userid,userclass"&Selectlist&" from Dv_online Order by userid desc,lastimebk Desc"	
	Else
		SQL="Select id,username,UserGroupID,userhidden,userid,userclass"&Selectlist&" from Dv_online where boardid="&Dvbbs.Boardid&" Order by userid desc,lastimebk Desc"	
	End If
	'Response.Write "SQL="&sql
	Dim j
	j=0
	'Dvbbs.Forum_setting(58)=30
	rs.open sql,conn,1,1
	If Not(Rs.BOF And Rs.EOF) Then
		If Dvbbs.BoardID=0 Then
			Dvbbs.Name="Forum_Online"
			Dvbbs.Value=Rs.recordcount
		End If
		Response.Write "Count="&Rs.recordcount&";"
		Rs.PageSize= CInt(Dvbbs.Forum_setting(58))
		Rs.AbsolutePage=page
		Response.Write "pageCount="&Rs.pageCount&";"
		Response.Write "PageSize="&Dvbbs.Forum_setting(58)&";"		
		Response.Write "page="&page&";"	
		Do while Not Rs.EOF
			For i=0 to Rs.Fields.count-1
				tmpdata=tmpdata & Rs(i)& "^&%&"
			Next 
			tmpdata=tmpdata&"%#!&"
			Rs.MoveNext
			j=j+1
			If j=CInt(Dvbbs.Forum_setting(58)) Then Exit Do
		Loop		
	End If
	tmpdata=Dvbbs.HTMLEncode(tmpdata)
	tmpdata=Replace(Replace(Replace(Replace(tmpdata&"","\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")
	Response.Write "Selectlist='"&Selectlist&"';"
	Response.Write "showonlist('"&tmpdata&"');"	
	Rs.Close
	Set Rs= Nothing 
End Sub

'�û���ͼ�껺�溯��,����״̬�б�ɵ��ã��û���ID|||�û���ͼ�꣩
Function GetGroupTitlePic()
	Dvbbs.Name="GetGroupTitlePic"
	If Dvbbs.ObjIsEmpty() Then
		Dim Rs,SQl
		SQL="select UserGroupID,TitlePic from [Dv_UserGroups] Order by UserGroupID "
		Set Rs=Dvbbs.Execute(SQL)
		'������Ĭ��Ϊ����
		SQL=Rs.GetString(,, "|||", "@@@", "Skins/Default/messages2.gif")
		Rs.close:Set Rs=Nothing
		'��������Աͼ�꣬����Ϊ0
		SQL="0|||" & template.pic(0) &"@@@"& SQL
		Dvbbs.Value = SQL
	End If
	GetGroupTitlePic=Dvbbs.Value
End Function
%>