<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/Dv_ClsOther.asp" -->
<%
	Dim Str
	Dvbbs.Stats="�鿴�ļ�"
	Dim Downid,Rs
	If CInt(Dvbbs.GroupSetting(49))=0 Then Dvbbs.AddErrCode(54)
	If request("id")="" Then
		Dvbbs.AddErrCode(35)
	ElseIf Not IsNumeric(request("id")) Then
		Dvbbs.AddErrCode(35)
	Else
		DownID=Clng(request("id"))
	End If
	Dvbbs.ShowErr()

	'��̳��������(�������¡����֡���Ǯ����������������������ɾ����ע��ʱ��)
	Dim BoardUserLimited
	BoardUserLimited = Split(Dvbbs.Board_Setting(55),"|")
	If Ubound(BoardUserLimited)=8 Then
		'����
		If Trim(BoardUserLimited(0))<>"0" And IsNumeric(BoardUserLimited(0)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(0)&"</B> ��������&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(8))<Clng(BoardUserLimited(0)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(0)&"</B> ��������&action=OtherErr"
		End If
		'����
		If Trim(BoardUserLimited(1))<>"0" And IsNumeric(BoardUserLimited(1)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(1)&"</B> ��������&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(22))<Clng(BoardUserLimited(1)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(1)&"</B> ��������&action=OtherErr"
		End If
		'��Ǯ
		If Trim(BoardUserLimited(2))<>"0" And IsNumeric(BoardUserLimited(2)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���Ǯ����Ϊ <B>"&BoardUserLimited(2)&"</B> ��������&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(21))<Clng(BoardUserLimited(2)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���Ǯ����Ϊ <B>"&BoardUserLimited(2)&"</B> ��������&action=OtherErr"
		End If
		'����
		If Trim(BoardUserLimited(3))<>"0" And IsNumeric(BoardUserLimited(3)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(3)&"</B> ��������&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(23))<Clng(BoardUserLimited(3)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(3)&"</B> ��������&action=OtherErr"
		End If
		'����
		If Trim(BoardUserLimited(4))<>"0" And IsNumeric(BoardUserLimited(4)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(4)&"</B> ��������&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(24))<Clng(BoardUserLimited(4)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(4)&"</B> ��������&action=OtherErr"
		End If
		'����
		If Trim(BoardUserLimited(5))<>"0" And IsNumeric(BoardUserLimited(5)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(5)&"</B> ��������&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(28))<Clng(BoardUserLimited(5)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���������Ϊ <B>"&BoardUserLimited(5)&"</B> ��������&action=OtherErr"
		End If
		'ɾ��
		If Trim(BoardUserLimited(6))<>"0" And IsNumeric(BoardUserLimited(6)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���ɾ������ <B>"&BoardUserLimited(6)&"</B> ��������&action=OtherErr"
			If Clng(Dvbbs.MyUserInfo(27))>Clng(BoardUserLimited(6)) Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û���ɾ������ <B>"&BoardUserLimited(6)&"</B> ��������&action=OtherErr"
		End If
		'ע��ʱ��
		If Trim(BoardUserLimited(7))<>"0" And IsNumeric(BoardUserLimited(7)) Then
			If Dvbbs.UserID = 0 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û�ע��ʱ����� <B>"&BoardUserLimited(7)&"</B> ���Ӳ�������&action=OtherErr"
			If DateDiff("s",Dvbbs.MyUserInfo(14),Now)<Clng(BoardUserLimited(7))*60 Then Response.redirect "showerr.asp?ErrCodes=<li>�������������û�ע��ʱ����� <B>"&BoardUserLimited(7)&"</B> ���Ӳ�������&action=OtherErr"
		End If
	End If
	If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
	If right(Dvbbs.Forum_Setting(76),1)<>"/" Then Dvbbs.Forum_Setting(76)=Dvbbs.Forum_Setting(76)&"/"
	Dim uploadpath,filename
	uploadpath=Dvbbs.Forum_Setting(76)
	Set Rs=Dvbbs.Execute("Select * From dv_upfile Where F_id="&downid)
	If Rs.Eof And Rs.Bof Then
		Dvbbs.AddErrCode(32)
	Else
		If Dvbbs.Forum_Setting(75)="0" Then
			Dvbbs.Execute("Update dv_upfile Set F_DownNum=F_DownNum+1 Where F_ID="&DownID)
			Response.Redirect uploadpath&rs("F_filename")
		Else
			filename=Replace(rs("F_filename"),"..","")&""
			If Request.ServerVariables("HTTP_REFERER")="" Or InStr(Request.ServerVariables("HTTP_REFERER"),Request.ServerVariables("SERVER_NAME"))=0 Or filename="" Then
				Response.Redirect "index.asp"
			Else
				Call downloadFile(Server.MapPath(Dvbbs.Forum_Setting(76)&filename))
				
			End If
		End If
	End If 
	Rs.close
	Set Rs=Nothing
	Dvbbs.ShowErr()
Sub downloadFile(strFile)
	On error resume next
	Server.ScriptTimeOut=999999
	Dim S,fso,f,intFilelength,strFilename
	strFilename = strFile
	Response.Clear
	Set s = Server.CreateObject("ADODB.Stream") 
	s.Open
	s.Type = 1 
	Set fso = Server.CreateObject("Scripting.FileSystemObject") 
	If Not fso.FileExists(strFilename) Then
		Response.Write("<h1>����: </h1><br>ϵͳ�Ҳ���ָ���ļ�")
		Exit Sub		
	End If
	Set f = fso.GetFile(strFilename)
		intFilelength = f.size
		s.LoadFromFile(strFilename)
		If err Then
		 	Response.Write("<h1>����: </h1>" & err.Description & "<p>")
			Response.End 
		End If
		Set fso=Nothing
		Dim Data
		Data=s.Read
		s.Close
		Set s=Nothing
		If Response.IsClientConnected Then 
			Response.AddHeader "Content-Disposition", "attachment; filename=" & f.name 
			Response.AddHeader "Content-Length", intFilelength 
 			Response.CharSet = "UTF-8" 
			Response.ContentType = "application/octet-stream"
			Response.BinaryWrite Data
			Response.Flush
		End If
End Sub
%>