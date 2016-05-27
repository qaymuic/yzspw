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
	Response.Write "您已经上传了头像"
	Response.End
End If

If SysSetting(Dvbbs.Forum_Setting(7)) = False or Clng(Dvbbs.Forum_Setting(53)) = 0 Then
	Response.Write "本系统未开放上传了头像功能"
	Response.End
End If

'系统设置
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
		UpUserFace()	'删除旧的头像文件
	Else
		Response.Write "只有文章数多于"& Dvbbs.Forum_Setting(54) &"篇才可以自定义头像！"
		Response.End
	End If
End If

'---------------------------------------------------------------
'头像上传开始
'---------------------------------------------------------------
Dim Upload_type
Dim Upload,File,FormName,FormPath,FileName,FileExt
FormPath="UploadFace/"
'If Right(FormPath,1)<>"/" then FormPath=FormPath&"/" 
'---------------------------------------------------------------
'上传组件选择:Upload_type=参数
'参数说明:0＝无组件，1＝LyfUpload，2＝AspUpload3.0，3＝SA-FileUp 4.0，4=DvFile.Upload V1.0
Upload_type=Cstr(Dvbbs.Forum_setting(43))	'默认设置为无组件上传
'---------------------------------------------------------------
'头像上传组件选取
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
		Response.Write "本系统未开放上传了头像功能"
		Response.Write "</body></html>"
		Response.End
End Select


'===========无组件上传(Upload_0)====================
Sub Upload_0()
	Set Upload = New UpFile_Class						''建立上传对象
	Upload.InceptFileType = "gif,jpg,bmp,jpeg,png"		'上传类型限制
	Upload.MaxSize = Int(Dvbbs.Forum_Setting(56))*1024	'限制大小
	Upload.GetDate()	'取得上传数据
	If Upload.Err > 0 Then
		Select Case Upload.Err
			Case 1 : Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			Case 2 : Response.Write "图片大小超过了限制 "&Dvbbs.Forum_Setting(56)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			Case 3 : Response.Write "所上传类型不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		End Select
		Exit Sub
	Else
		'FormPath=Upload.Form("filepath")
		 For Each FormName in Upload.file		''列出所有上传了的文件
			 Set File = Upload.File(FormName)	''生成一个文件对象
			 If File.Filesize<10 Then
		 		Response.Write "请先选择你要上传的图片　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				Exit Sub
	 		End If
			FileExt	= FixName(File.FileExt)
 			If Not ( CheckFileExt(FileExt) and CheckFileType(File.FileType) ) Then
 				Response.Write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				Exit Sub
			End If
 			FileName=FormPath&UserFaceName(FileExt)
 			If File.FileSize>0 Then   ''如果 FileSize > 0 说明有文件数据
				File.SaveToFile Server.mappath(FileName)   ''保存文件
				Response.Write "<script>parent.document.images['face'].src='" &FileName& "';parent.document.theForm.myface.value='"&FileName&"';</script>"
				Response.Write "<script>parent.document.images['face'].width='" &File.FileWidth& "';parent.document.images['face'].height='"&File.FileHeight&"';</script>"
				Response.Write "<script>parent.document.theForm.height.value='" &File.FileHeight& "';parent.document.theForm.width.value='"&File.FileWidth&"';</script>"
				Session("upface")="done"
				Response.Write "图片上传成功!"
 			End If
 			Set File=Nothing
		Next
	End If
	Set Upload=Nothing
End Sub

'===========LyfUpload组件上传(Upload_1)=========================
Sub Upload_1()
	Dim obj,FileName,FileExt_a
	Dim ss
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'大小
    	obj.maxsize = Int(Dvbbs.Forum_Setting(56))*1024
	'类型
    	obj.extname = "gif,jpg,bmp,jpeg,png"
	'重命名
	'在目录后加(/)
	'if right(FormPath,1)<>"/" then FormPath=FormPath&"/" 
	If obj.request("fname") = "" Or IsNull(obj.request("fname")) then
		Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		Exit Sub
	End If
	FileExt		= Mid(obj.Request("fname"), InStrRev(obj.Request("fname"), ".")+1)
	FileExt		= FixName(FileExt)
	FileName	= UserFaceName(FileExt)
	If Not ( CheckFileExt(FileExt) and CheckFileType(obj.FileType("file1")) ) Then
 		Response.Write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		Exit Sub
	End If
	ss=obj.SaveFile("file1",Server.MapPath(FormPath), true,FileName)
	If ss = "3" Then
		Response.Write ("文件名重复![ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		Response.Write "</body></html>"
		Response.End
	ElseIf ss = "0" Then
   		Response.Write ("文件尺寸过大![ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		Response.Write "</body></html>"
		Response.End
	ElseIf ss = "1" Then
		Response.Write ("文件不是指定类型文件![ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		Response.Write "</body></html>"
		Response.End
	ElseIf ss = "" Then
		Response.Write ("文件上传失败![ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		Response.Write "</body></html>"
		response.end
	Else
		Response.Write "图片上传成功!" 
		Response.Write "<script>parent.document.images['face'].src='" &FormPath&FileName& "';parent.document.theForm.myface.value='" &FormPath&FileName & "'</script>"
		session("upface")="done"
		Response.Write "</body></html>"
	End if
	Set obj=nothing
End Sub

''===========================AspUpload3.0组件上传============================
Sub Upload_2()
	Dim Count
	on Error Resume Next
	Set Upload = Server.CreateObject("Persits.Upload") 
	Upload.OverwriteFiles = False   '不能复盖
	Upload.IgnoreNoPost = True
	Upload.SetMaxSize int(Dvbbs.Forum_Setting(56))*1024, True	 '限制大小
	Count = Upload.Save
	If Err.Number = 8 Then 
  		 Response.Write "文件大小超过了限制 "&Dvbbs.Forum_Setting(56)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]" 
	Else 
		If Err <> 0 Then 
      			Response.Write "错误信息: " & Err.Description 
		Else
			If Count < 1 Then 
				Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				Exit Sub
			End If
			For Each file in Upload.Files	'列出所有上传文件
				FileExt = Replace(File.ext,".","")
				FileExt	= FixName(FileExt)
				'判断文件类型
				If Not ( CheckFileExt(FileExt) and File.ImageType <> "UNKNOWN" ) Then
					Response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
					Exit Sub
				End If
				'文件变量付值
				FileName=UserFaceName(FileExt)
				FileName=FormPath&FileName
				File.saveas Server.MapPath(FileName)	'上传保存文件
				Response.Write "图片上传成功!"
				Response.Write "<script>parent.document.images['face'].src='" &FileName& "';parent.document.theForm.myface.value='" &FileName& "';"
				Response.Write "parent.document.images['face'].width='" &File.ImageWidth& "';parent.document.images['face'].height='"&File.ImageHeight&"';"
				Response.Write "parent.document.theForm.height.value='" &File.ImageHeight& "';parent.document.theForm.width.value='"&File.ImageWidth&"';</script>"
				Session("upface")="done"
			Next
		End If 
	End If
	Set Upload =Nothing
End Sub

''===========================SA-FileUp 4.0组件上传============================
Sub Upload_3()
	Dim oFileUp
	Dim FileExt_a,Filesize,file_name
	Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")
	If Not oFileUp.Form("file1").IsEmpty Then
		FileName	= oFileUp.Form("file1").ShortFileName	 '原文件名
		FileExt		= Mid(Filename, InStrRev(Filename, ".")+1)
		FileExt		= FixName(FileExt)
		Filesize	= oFileUp.Form("file1").TotalBytes 
		If Filesize>Int(Dvbbs.Forum_Setting(56))*1024 Then
			Response.Write "文件大小超过了限制 "&Dvbbs.Forum_Setting(56)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]" 
			Exit Sub
		End If
		'判断文件类型
		If Not ( CheckFileExt(FileExt) and CheckFileType(oFileUp.Form("file1").ContentType) ) Then
			Response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			Exit Sub
		End If
		'文件变量付值
		FileName=UserFaceName(FileExt)
		FileName=FormPath&FileName
		'保存文件
		oFileUp.Form("file1").Saveas Server.MapPath(FileName)
		Response.Write "图片上传成功!"
		Response.Write "<script>parent.document.images['face'].src='" &FileName& "';parent.document.theForm.myface.value='" &FileName& "'</script>"
		Session("upface")="done"
	Else
		Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	End If
	Set oFileUp = Nothing
End Sub

'===========DvFile-Up V1.0组件上传====================
Sub Upload_4()
	Set Upload = Server.CreateObject("DvFile.Upload")	''建立上传对象
	Upload.InceptFileType = "gif,jpg,bmp,jpeg,png"		''上传类型限制
	Upload.MaxSize = Int(Dvbbs.Forum_Setting(56))*1024	''限制大小
	Upload.Install										''取得上传数据
	If Upload.Err > 0 Then
		Select Case Upload.Err
			Case 2 : Response.Write Upload.Description & Dvbbs.Forum_Setting(56) &"KB"	''图片大小超过了限制 "&Upload.MaxSize&"K　
			Case Else
			Response.Write Upload.Description
		End Select
		Response.Write "　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		Exit Sub
	Else
		If Upload.Count>1 Then Response.Write "上传个数超过限制" : Exit Sub
		For Each FormName in Upload.File		''列出所有上传了的文件
			Set File = Upload.File(FormName)	''生成一个文件对象
				If File.Filesize<10 Then
					Response.Write "请先选择你要上传的图片　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
					Exit Sub
				End If
				FileExt	= FixName(File.FileExt)
 				If Not ( CheckFileExt(FileExt) and CheckFileType(File.FileType) ) Then
 					Response.Write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
					Exit Sub
				End If
 				FileName = FormPath & UserFaceName(FileExt)
 				If File.FileSize>0 Then   ''如果 FileSize > 0 说明有文件数据
					Session("upface")="done"
					Upload.SaveToFile Server.Mappath(FileName),FormName		''保存文件
					Response.Write "<script>parent.document.images['face'].src='" &FileName& "';parent.document.theForm.myface.value='"&FileName&"';</script>"
					Response.Write "<script>parent.document.images['face'].width='" &File.FileWidth& "';parent.document.images['face'].height='"&File.FileHeight&"';</script>"
					Response.Write "<script>parent.document.theForm.height.value='" &File.FileHeight& "';parent.document.theForm.width.value='"&File.FileWidth&"';</script>"
					Response.Write "图片上传成功!"
 				End If
 			Set File=Nothing
		Next
	End If
	Set Upload=Nothing
End Sub

'判断文件类型是否合格
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
'格式后缀
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
'文件Content-Type判断
Private Function CheckFileType(FileType)
	CheckFileType = False
	If Left(Cstr(Lcase(Trim(FileType))),6)="image/" Then CheckFileType = True
End Function
'文件名明名
Private Function UserFaceName(FileExt)
	Dim UserID,RanNum
	UserID = ""
	If Dvbbs.UserID>0 Then UserID = Dvbbs.UserID&"_"
	Randomize
	RanNum = Int(90000*rnd)+10000
 	UserFaceName = UserID&Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&RanNum&"."&FileExt
End Function
'删除旧头像
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