<!--#include FILE="conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="Upload.inc"-->
<%
If Not Dvbbs.ChkPost Then
	Response.End
End If
Dvbbs.LoadTemplates("usermanager")
Dvbbs.Stats = Dvbbs.MemberName&template.Strings(1)
Dvbbs.Head()
%>
<table width="100%" height="100%" border=0  cellspacing="0" cellpadding="0"><tr><td class=tablebody1 width="100%" height="100%" >
<script>
parent.document.theForm.Submit.disabled=false;
parent.document.theForm.Submit2.disabled=false;
</script>
<%
If Session("upface")="done" Then
	Response.Write "���Ѿ��ϴ���ͷ��"
	Response.End
End If

If SysSetting(Dvbbs.Forum_Setting(7)) = False or Clng(Dvbbs.Forum_Setting(53)) = 0 Then
	Response.Write "��ϵͳδ�����ϴ���ͷ����"
	Response.End
End If

'ϵͳ����
Function SysSetting(Setting)
	SysSetting = False
	Select Case Clng(Setting)
		Case 1 : SysSetting = True
		Case 2 :
			If Dvbbs.UserID > 0 Then SysSetting = True
	End Select
End Function

If Dvbbs.UserID>0 Then
	If Clng(Dvbbs.MyUserInfo(8))>Clng(Dvbbs.Forum_Setting(54)) Then
		UpUserFace()	'ɾ���ɵ�ͷ���ļ�
	Else
		Response.Write "ֻ������������"& Dvbbs.Forum_Setting(54) &"ƪ�ſ����Զ���ͷ��"
		Response.End
	End If
End If

'---------------------------------------------------------------
'ͷ���ϴ���ʼ
'---------------------------------------------------------------
Dim Upload_type
Dim Upload,File,FormName,FormPath,FileName,FileExt
FormPath="UploadFace/"
'If Right(FormPath,1)<>"/" then FormPath=FormPath&"/" 
'---------------------------------------------------------------
'�ϴ����ѡ��:Upload_type=����
'����˵��:0���������1��LyfUpload��2��AspUpload3.0��3��SA-FileUp 4.0��4=DvFile.Upload V1.0
Upload_type=Cstr(Dvbbs.Forum_setting(43))	'Ĭ������Ϊ������ϴ�
'---------------------------------------------------------------
'ͷ���ϴ����ѡȡ
Select Case Upload_type
	Case 0
		Call Upload_0()
	Case 1
		Call Upload_1()
	Case 2
		Call Upload_2()
	Case 3
		Call Upload_3()
	Case 4
		Call Upload_4()
	Case Else
		Response.Write "��ϵͳδ�����ϴ���ͷ����"
		Response.Write "</body></html>"
		Response.End
End Select


'===========������ϴ�(Upload_0)====================
Sub Upload_0()
	Set Upload = New UpFile_Class						''�����ϴ�����
	Upload.InceptFileType = "gif,jpg,bmp,jpeg,png"		'�ϴ���������
	Upload.MaxSize = Int(Dvbbs.Forum_Setting(56))*1024	'���ƴ�С
	Upload.GetDate()	'ȡ���ϴ�����
	If Upload.Err > 0 Then
		Select Case Upload.Err
			Case 1 : Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			Case 2 : Response.Write "ͼƬ��С���������� "&Dvbbs.Forum_Setting(56)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			Case 3 : Response.Write "���ϴ����Ͳ���ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		End Select
		Exit Sub
	Else
		'FormPath=Upload.Form("filepath")
		 For Each FormName in Upload.file		''�г������ϴ��˵��ļ�
			 Set File = Upload.File(FormName)	''����һ���ļ�����
			 If File.Filesize<10 Then
		 		Response.Write "����ѡ����Ҫ�ϴ���ͼƬ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				Exit Sub
	 		End If
			FileExt	= FixName(File.FileExt)
 			If Not ( CheckFileExt(FileExt) and CheckFileType(File.FileType) ) Then
 				Response.Write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				Exit Sub
			End If
 			FileName=FormPath&UserFaceName(FileExt)
 			If File.FileSize>0 Then   ''��� FileSize > 0 ˵�����ļ�����
				File.SaveToFile Server.mappath(FileName)   ''�����ļ�
				Response.Write "<script>parent.document.images['face'].src='" &FileName& "';parent.document.theForm.myface.value='"&FileName&"';</script>"
				Response.Write "<script>parent.document.images['face'].width='" &File.FileWidth& "';parent.document.images['face'].height='"&File.FileHeight&"';</script>"
				Response.Write "<script>parent.document.theForm.height.value='" &File.FileHeight& "';parent.document.theForm.width.value='"&File.FileWidth&"';</script>"
				Session("upface")="done"
				Response.Write "ͼƬ�ϴ��ɹ�!"
 			End If
 			Set File=Nothing
		Next
	End If
	Set Upload=Nothing
End Sub

'===========LyfUpload����ϴ�(Upload_1)=========================
Sub Upload_1()
	Dim obj,FileName,FileExt_a
	Dim ss
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'��С
    	obj.maxsize = Int(Dvbbs.Forum_Setting(56))*1024
	'����
    	obj.extname = "gif,jpg,bmp,jpeg,png"
	'������
	'��Ŀ¼���(/)
	'if right(FormPath,1)<>"/" then FormPath=FormPath&"/" 
	If obj.request("fname") = "" Or IsNull(obj.request("fname")) then
		Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		Exit Sub
	End If
	FileExt		= Mid(obj.Request("fname"), InStrRev(obj.Request("fname"), ".")+1)
	FileExt		= FixName(FileExt)
	FileName	= UserFaceName(FileExt)
	If Not ( CheckFileExt(FileExt) and CheckFileType(obj.FileType("file1")) ) Then
 		Response.Write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		Exit Sub
	End If
	ss=obj.SaveFile("file1",Server.MapPath(FormPath), true,FileName)
	If ss = "3" Then
		Response.Write ("�ļ����ظ�![ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		Response.Write "</body></html>"
		Response.End
	ElseIf ss = "0" Then
   		Response.Write ("�ļ��ߴ����![ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		Response.Write "</body></html>"
		Response.End
	ElseIf ss = "1" Then
		Response.Write ("�ļ�����ָ�������ļ�![ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		Response.Write "</body></html>"
		Response.End
	ElseIf ss = "" Then
		Response.Write ("�ļ��ϴ�ʧ��![ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		Response.Write "</body></html>"
		response.end
	Else
		Response.Write "ͼƬ�ϴ��ɹ�!" 
		Response.Write "<script>parent.document.images['face'].src='" &FormPath&FileName& "';parent.document.theForm.myface.value='" &FormPath&FileName & "'</script>"
		session("upface")="done"
		Response.Write "</body></html>"
	End if
	Set obj=nothing
End Sub

''===========================AspUpload3.0����ϴ�============================
Sub Upload_2()
	Dim Count
	on Error Resume Next
	Set Upload = Server.CreateObject("Persits.Upload") 
	Upload.OverwriteFiles = False   '���ܸ���
	Upload.IgnoreNoPost = True
	Upload.SetMaxSize int(Dvbbs.Forum_Setting(56))*1024, True	 '���ƴ�С
	Count = Upload.Save
	If Err.Number = 8 Then 
  		 Response.Write "�ļ���С���������� "&Dvbbs.Forum_Setting(56)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]" 
	Else 
		If Err <> 0 Then 
      			Response.Write "������Ϣ: " & Err.Description 
		Else
			If Count < 1 Then 
				Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				Exit Sub
			End If
			For Each file in Upload.Files	'�г������ϴ��ļ�
				FileExt = Replace(File.ext,".","")
				FileExt	= FixName(FileExt)
				'�ж��ļ�����
				If Not ( CheckFileExt(FileExt) and File.ImageType <> "UNKNOWN" ) Then
					Response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
					Exit Sub
				End If
				'�ļ�������ֵ
				FileName=UserFaceName(FileExt)
				FileName=FormPath&FileName
				File.saveas Server.MapPath(FileName)	'�ϴ������ļ�
				Response.Write "ͼƬ�ϴ��ɹ�!"
				Response.Write "<script>parent.document.images['face'].src='" &FileName& "';parent.document.theForm.myface.value='" &FileName& "';"
				Response.Write "parent.document.images['face'].width='" &File.ImageWidth& "';parent.document.images['face'].height='"&File.ImageHeight&"';"
				Response.Write "parent.document.theForm.height.value='" &File.ImageHeight& "';parent.document.theForm.width.value='"&File.ImageWidth&"';</script>"
				Session("upface")="done"
			Next
		End If 
	End If
	Set Upload =Nothing
End Sub

''===========================SA-FileUp 4.0����ϴ�============================
Sub Upload_3()
	Dim oFileUp
	Dim FileExt_a,Filesize,file_name
	Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")
	If Not oFileUp.Form("file1").IsEmpty Then
		FileName	= oFileUp.Form("file1").ShortFileName	 'ԭ�ļ���
		FileExt		= Mid(Filename, InStrRev(Filename, ".")+1)
		FileExt		= FixName(FileExt)
		Filesize	= oFileUp.Form("file1").TotalBytes 
		If Filesize>Int(Dvbbs.Forum_Setting(56))*1024 Then
			Response.Write "�ļ���С���������� "&Dvbbs.Forum_Setting(56)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]" 
			Exit Sub
		End If
		'�ж��ļ�����
		If Not ( CheckFileExt(FileExt) and CheckFileType(oFileUp.Form("file1").ContentType) ) Then
			Response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			Exit Sub
		End If
		'�ļ�������ֵ
		FileName=UserFaceName(FileExt)
		FileName=FormPath&FileName
		'�����ļ�
		oFileUp.Form("file1").Saveas Server.MapPath(FileName)
		Response.Write "ͼƬ�ϴ��ɹ�!"
		Response.Write "<script>parent.document.images['face'].src='" &FileName& "';parent.document.theForm.myface.value='" &FileName& "'</script>"
		Session("upface")="done"
	Else
		Response.Write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	End If
	Set oFileUp = Nothing
End Sub

'===========DvFile-Up V1.0����ϴ�====================
Sub Upload_4()
	Set Upload = Server.CreateObject("DvFile.Upload")	''�����ϴ�����
	Upload.InceptFileType = "gif,jpg,bmp,jpeg,png"		''�ϴ���������
	Upload.MaxSize = Int(Dvbbs.Forum_Setting(56))*1024	''���ƴ�С
	Upload.Install										''ȡ���ϴ�����
	If Upload.Err > 0 Then
		Select Case Upload.Err
			Case 2 : Response.Write Upload.Description & Dvbbs.Forum_Setting(56) &"KB"	''ͼƬ��С���������� "&Upload.MaxSize&"K��
			Case Else
			Response.Write Upload.Description
		End Select
		Response.Write "��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		Exit Sub
	Else
		If Upload.Count>1 Then Response.Write "�ϴ�������������" : Exit Sub
		For Each FormName in Upload.File		''�г������ϴ��˵��ļ�
			Set File = Upload.File(FormName)	''����һ���ļ�����
				If File.Filesize<10 Then
					Response.Write "����ѡ����Ҫ�ϴ���ͼƬ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
					Exit Sub
				End If
				FileExt	= FixName(File.FileExt)
 				If Not ( CheckFileExt(FileExt) and CheckFileType(File.FileType) ) Then
 					Response.Write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
					Exit Sub
				End If
 				FileName = FormPath & UserFaceName(FileExt)
 				If File.FileSize>0 Then   ''��� FileSize > 0 ˵�����ļ�����
					Session("upface")="done"
					Upload.SaveToFile Server.Mappath(FileName),FormName		''�����ļ�
					Response.Write "<script>parent.document.images['face'].src='" &FileName& "';parent.document.theForm.myface.value='"&FileName&"';</script>"
					Response.Write "<script>parent.document.images['face'].width='" &File.FileWidth& "';parent.document.images['face'].height='"&File.FileHeight&"';</script>"
					Response.Write "<script>parent.document.theForm.height.value='" &File.FileHeight& "';parent.document.theForm.width.value='"&File.FileWidth&"';</script>"
					Response.Write "ͼƬ�ϴ��ɹ�!"
 				End If
 			Set File=Nothing
		Next
	End If
	Set Upload=Nothing
End Sub

'�ж��ļ������Ƿ�ϸ�
Private Function CheckFileExt(FileExt)
	Dim ForumUpload,i
	ForumUpload="gif,jpg,bmp,jpeg,png"
	ForumUpload=Split(ForumUpload,",")
	CheckFileExt=False
	For i=0 to UBound(ForumUpload)
		If LCase(FileExt)=Lcase(Trim(ForumUpload(i))) Then
			CheckFileExt=True
			Exit Function
		End If
	Next
End Function
'��ʽ��׺
Function FixName(UpFileExt)
	If IsEmpty(UpFileExt) Then Exit Function
	FixName = Lcase(UpFileExt)
	FixName = Replace(FixName,Chr(0),"")
	FixName = Replace(FixName,".","")
	FixName = Replace(FixName,"asp","")
	FixName = Replace(FixName,"asa","")
	FixName = Replace(FixName,"aspx","")
	FixName = Replace(FixName,"cer","")
	FixName = Replace(FixName,"cdx","")
	FixName = Replace(FixName,"htr","")
End Function
'�ļ�Content-Type�ж�
Private Function CheckFileType(FileType)
	CheckFileType = False
	If Left(Cstr(Lcase(Trim(FileType))),6)="image/" Then CheckFileType = True
End Function
'�ļ�������
Private Function UserFaceName(FileExt)
	Dim UserID,RanNum
	UserID = ""
	If Dvbbs.UserID>0 Then UserID = Dvbbs.UserID&"_"
	Randomize
	RanNum = Int(90000*rnd)+10000
 	UserFaceName = UserID&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&RanNum&"."&FileExt
End Function
'ɾ����ͷ��
Sub UpUserFace()
	on Error Resume Next
	Dim objFSO,OldUserFace
	OldUserFace = Server.MapPath(FormPath&Dvbbs.UserID&"_")&"*.*"
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	'If objFSO.FileExists(OldUserFace) Then
		objFSO.DeleteFile OldUserFace
		If Err<>0 Then Err.Clear
	'End If
	Set objFSO = Nothing
End Sub
%>
</td></tr></table>
</body>
</html>