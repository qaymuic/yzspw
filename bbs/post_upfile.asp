<!--#include File="conn.asp"-->
<!--#include File="upload.inc"-->
<!-- #include File="inc/const.asp" -->
<!-- #include File="inc/dv_clsother.asp" -->
<script>
parent.document.Dvform.Submit.disabled=false;
parent.document.Dvform.Submit2.disabled=false;
</script>
<table width="100%" height="100%" border=0 cellspacing=0 cellpadding=0>
<tr><td class=tablebody2 valign=top height=40>
<%
Dvbbs.Loadtemplates("")
Dvbbs.Head()
Call Dvbbs.ShowErr()
Server.ScriptTimeOut=999999'要是你的论坛支持上传的文件比较大，就必须设置。
Dim upload_type,upload_ViewType
'-----------------------------------------------------------------------------
'设置上传方式upload_type值： 0＝无组件，1＝lyfupload 1.2版，2＝Aspupload3.0，3＝SA-FileUp 4.0，4=DvFile.Upload V1.0
'-----------------------------------------------------------------------------
upload_type=Cint(Dvbbs.Forum_Setting(43))
'-----------------------------------------------------------------------------
'创建生成预览图片,需要图片读写组件支持.(根目录下要有PreviewImage文件夹存放文件)
'设置支持组件upload_ViewType值： 
'0＝ CreatePreviewImage， 1＝ AspJpeg ，2=SoftArtisans ImgWriter V1.21，3=SJCatSoft V2.6 
Dim previewpath,F_Viewname
F_Viewname=""
previewpath="PreviewImage/"
upload_ViewType=Cint(Dvbbs.Forum_Setting(45))
'-----------------------------------------------------------------------------
'提交验证
If Not Dvbbs.ChkPost Then
	Response.End
End If
If Dvbbs.Userid=0 Then
	Response.write "你还未登陆！"
	Response.End
End If

DateUpNum=Clng(Dvbbs.UserToday(2))
UpNum=request.cookies("upNum")
If UpNum ="" then UpNum=0
UpNum=int(UpNum)

If Cint(Dvbbs.GroupSetting(7))=0 then
	Response.write "您没有在本论坛上传文件的权限"
	Response.End
End If
If upNum >= Clng(Dvbbs.GroupSetting(40)) then
 	Response.write "一次只能上传"&Dvbbs.GroupSetting(40)&"个文件！"
	Response.End
End If
If dateupnum+upNum >= Clng(Dvbbs.GroupSetting(50)) then
 	Response.write "您今天上传的文件已超出了"&Dvbbs.GroupSetting(50)&"个！"
	Response.End
End If
'-----------------------------------------------------------------------------
'定义变量
Dim Forumupload
Dim FormName,FormPath,Filename,File_name,FileExt,Filesize,F_Type,rename
Dim upNum,dateupnum
Dim TempSize,ImageWidth,ImageHeight
Dim ImageMode
ImageMode=Dvbbs.Forum_Setting(73)
If ImageMode="0" Then ImageMode=""
ImageWidth=80
ImageHeight=80
TempSize=Split(Dvbbs.Forum_Setting(72),"|")
If Ubound(TempSize)=1 Then
	ImageWidth=TempSize(0)
	ImageHeight=TempSize(1)
End If
FormPath=CheckFolder&CreatePath()	'上传目录路径

'On Error Resume Next 
Select case upload_type
	Case 0
		Call upload_0()
	Case 1
		Call upload_1()
	Case 2
		Call upload_2()
	Case 3
		Call upload_3()
	Case 4
		Call upload_4()
	Case Else
		Response.write "本系统未开放插件功能"
		Response.End
End Select

'===========================无组件上传============================
Sub upload_0()
	Dim upload,File,UpCount
	UpCount=0
	Set upload = new UpFile_Class						''建立上传对象
	Upload.InceptFileType = Replace(Dvbbs.Board_Setting(19),"|",",")
	Upload.MaxSize = Int(Dvbbs.GroupSetting(44))*1024
	Upload.GetDate ()	'取得上传数据
	If upload.err > 0 then
		Select Case upload.err
		Case 1
			Response.write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		Case 2
			Response.write "文件大小超过了限制 "&Dvbbs.GroupSetting(44)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		Case 3
			Response.write "文件类型不正确 "&Dvbbs.GroupSetting(44)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		End Select
		Exit Sub
	Else
		For Each FormName In upload.File		''列出所有上传了的文件
			If upNum >= Int(Dvbbs.GroupSetting(40)) or dateupnum+upNum >= clng(Dvbbs.GroupSetting(50)) then
				Response.write "已达到上传数的上限。"
				EXIT SUB
			End If
			Set File=upload.File(FormName)		''生成一个文件对象
			FileExt=FixName(File.FileExt)
			'判断文件类型
			If CheckFileExt(FileExt)=false then
				Response.write "文件格式不正确,或不能为空　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				EXIT SUB
			End If
			'付值变量
			F_Type		=	CheckFiletype(FileExt)
			File_name	=	CreateName()
			Filename	=	File_name&"."&FileExt
			rename		=	CreatePath&Filename&"|"
			Filename	=	FormPath&Filename
			Filesize	=	File.FileSize
			'记录文件
			If Filesize>0 then								'如果 FileSize > 0 说明有文件数据
			File.SaveToFile Server.mappath(FileName)		''执行上传文件
			'创建生成预览图片
				If upload_ViewType<>999 and F_Type=1 then
					F_Viewname=previewpath&"pre"&File_name&".jpg"
					Call CreateView(FileName,F_Viewname)
				End If
			'记录文件
			Call checksave()
			UpCount=UpCount+1
			End If
			Set File=Nothing
		Next
	End If
	Set upload=Nothing
	Call Suc_upload(UpCount,upNum)
End Sub

''===========================lyfupload组件上传1.2版============================
Sub upload_1()
	Dim obj,Filepath,FileExt_a,UpCount
	Dim ss,i
	Dim TempExt,TempFileValue
	UpCount=0
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'限制大小
	obj.maxsize=Int(Dvbbs.GroupSetting(44))*1024
	'限制类型
	obj.extname=Replace(Dvbbs.Board_Setting(19),"|",",")
	Filepath=Server.MapPath(FormPath)
	'在目录后加(/)
	If Right(Filepath,1)<>"\" Then Filepath=Filepath&"\"
	For i=1 to obj.Request("upcount")
	 TempFileValue	=	"file"&i
		FileExt_a	=	Split(obj.Request(TempFileValue),"""")
		TempExt		=	FileExt_a(1)
		If TempExt="" or isnull(TempExt) then
			Response.write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			Exit Sub
		End If
		FileExt		=	Mid(TempExt, InStrRev(TempExt, ".")+1)
		FileExt		=	FixName(FileExt)
		File_name	=	CreateName()
		Filename	=	File_name&"."&FileExt
		rename		=	CreatePath&Filename & "|"
		Filesize	=	obj.Filesize
		'判断文件类型及付值变量
		F_Type		=CheckFiletype(FileExt)

		'判断文件类型
		If CheckFileExt(FileExt)=False then
			Response.write "文件格式不正确,或不能为空　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
			EXIT SUB
		End If
		'保存文件
		ss=obj.SaveFile(TempFileValue,Server.MapPath(FormPath), false,Filename)
		If ss= "3" then
			Response.write ("文件名重复!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
			EXIT SUB
		ElseIf ss= "0" then
			Response.write ("文件大小超过了限制 "&Dvbbs.GroupSetting(44)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
			EXIT SUB
		ElseIf ss = "1" then
			Response.write ("文件不是指定类型文件!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
			EXIT SUB
		ElseIf ss = "" then
			Response.write ("文件上传失败!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
			EXIT SUB
		Else
			Filename=FormPath&Filename
			'创建生成预览图片
			If upload_ViewType<>999 and F_Type=1 then
				F_Viewname=previewpath&"pre"&File_name&".jpg"
				call CreateView(FileName,F_Viewname)
			End If
			'记录文件
			call checksave()
			UpCount=UpCount+1
		End If
	Next
	set obj=nothing
	Call Suc_upload(UpCount,upNum)
End Sub

''===========================Aspupload3.0组件上传============================
sub upload_2()
	On Error Resume Next
	Dim Upload,File
	Dim FilePath
	Dim Count,UpCount
	UpCount=0
	Set Upload = Server.CreateObject("Persits.Upload") 
	Upload.OverwriteFiles = false								'不能复盖
	Upload.IgnoreNoPost = True
	Upload.SetMaxSize Int(Dvbbs.GroupSetting(44))*1024, True	'限制大小
	Count = Upload.Save
	If Err.Number = 8 Then 
	   Response.write "文件大小超过了限制 "&Dvbbs.GroupSetting(44)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]" 
	Else 
		If Err <> 0 Then 
			Response.write "错误信息: " & Err.Description
			EXIT SUB
		Else
			If Count < 1 Then 
				Response.write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				EXIT SUB
			End If
		For Each File in Upload.Files	'列出所有上传文件
			If upNum >= Int(Dvbbs.GroupSetting(40)) or dateupnum+upNum >= Clng(Dvbbs.GroupSetting(50)) Then
				Response.write "已达到上传数的上限。"
				Exit Sub
			End If
			FileExt = Replace(File.ext,".","")
			FileExt = FixName(FileExt)
			'判断文件类型
			If CheckFileExt(FileExt)=False then
				Response.write "文件格式不正确,或不能为空　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				EXIT SUB
			End If
			'文件变量付值
			File_name	=	CreateName()
			Filename	=	File_name&"."&FileExt
			rename		=	CreatePath&Filename & "|"
			Filename	=	FormPath&Filename
			Filesize	=	File.Size
			F_Type		=	CheckFiletype(FileExt)
			File.saveas Server.MapPath(Filename)	'上传保存文件
			'创建生成预览图片
			If upload_ViewType<>999 and F_Type=1 then
				F_Viewname=previewpath&"pre"&File_name&".jpg"
				Call CreateView(FileName,F_Viewname)
			End If
			'记录文件
			Call checksave()			'记录文件
			UpCount=UpCount+1
		Next
		Call Suc_upload(UpCount,upNum)
		End If 
	End If
	Set Upload = Nothing
End Sub

''===========================SA-FileUp 4.0组件上传FileUpSE V4.09============================
sub upload_3()
	Dim oFileUp,UpCount,FileExt_a
	UpCount=0
	Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")
	'oFileUp.Path = Server.MapPath(FormPath)
	For Each FormName In oFileUp.Form
		If IsObject(oFileUp.Form(FormName)) Then
			If Not oFileUp.Form(FormName).IsEmpty Then
				oFileUp.Form(FormName).Maxbytes=int(Dvbbs.GroupSetting(44))*1024	'限制大小
				Filesize=oFileUp.Form(FormName).TotalBytes
				If Filesize>int(Dvbbs.GroupSetting(44))*1024 then
					Response.write "文件大小超过了限制 "&Dvbbs.GroupSetting(44)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]" 
					Exit sub
				End If
				Filename	= oFileUp.Form(FormName).ShortFileName	 '原文件名
				FileExt		= Mid(Filename, InStrRev(Filename, ".")+1)
				FileExt		= FixName(FileExt)
				'判断文件类型
				If CheckFileExt(FileExt)=false then
					Response.write "文件格式不正确,或不能为空　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
					Exit Sub
				End If
				'文件变量付值
				File_name	=	CreateName()
				Filename	=	File_name&"."&FileExt
				rename		=	CreatePath&Filename & "|"
				Filename	=	FormPath&Filename
				F_Type		=	CheckFiletype(FileExt)
				
				'保存文件
				oFileUp.Form(FormName).Saveas Server.MapPath(Filename)
				'创建生成预览图片
				If upload_ViewType<>999 and F_Type=1 then
					F_Viewname=previewpath&"pre"&File_name&".jpg"
					Call CreateView(FileName,F_Viewname)
				End If
				'记录文件
				Call checksave()			'记录文件
				UpCount	= UpCount+1
			Else
				Response.write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				EXIT SUB
			End If
		End If
	Next
	Set oFileUp = Nothing
	Call Suc_upload(UpCount,upNum)
End Sub

'===========================DvFile.Upload V1.0组件上传============================
Sub upload_4()
	Dim upload,File,UpCount
	UpCount=0
	Set upload = Server.CreateObject("DvFile.Upload")			''建立上传对象
	Upload.InceptFileType = Replace(Dvbbs.Board_Setting(19),"|",",")
	Upload.MaxSize = Int(Dvbbs.GroupSetting(44))*1024
	Upload.Install		'取得上传数据
	If upload.err > 0 Then
		Select Case Upload.Err
			Case 1 : Response.Write Upload.Description	''请先选择你要上传的文件
			Case 2 : Response.Write Upload.Description & Dvbbs.GroupSetting(44) &"KB"	''图片大小超过了限制 "&Upload.MaxSize&"K　
			Case 3 : Response.Write Upload.Description	''非法的上传类型
			Case 4 : Response.Write Upload.Description	''所上传的类型受系统限制
			Case 5 : Response.Write Upload.Description	''参数有误，上传意外中止
		End Select
		Response.Write "　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		Exit Sub
	Else
		For Each FormName In upload.File		''列出所有上传了的文件
			If upNum >= Int(Dvbbs.GroupSetting(40)) or dateupnum+upNum >= Clng(Dvbbs.GroupSetting(50)) then
				Response.write "已达到上传数的上限。"
				EXIT SUB
			End If
			Set File = upload.File(FormName)		''生成一个文件对象
			FileExt = FixName(File.FileExt)
			'判断文件类型
			If CheckFileExt(FileExt)=False then
				Response.write "文件格式不正确,或不能为空　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
				EXIT SUB
			End If
			'付值变量
			F_Type		=	CheckFiletype(FileExt)
			File_name	=	CreateName()
			Filename	=	File_name&"."&FileExt
			Rename		=	CreatePath&Filename&"|"
			Filename	=	FormPath&Filename
			Filesize	=	File.FileSize
			'记录文件
			If Filesize>0 then								''如果 FileSize > 0 说明有文件数据
			Upload.SaveToFile Server.Mappath(FileName),FormName		''保存文件
			'创建生成预览图片
				If upload_ViewType<>999 And F_Type=1 Then
					F_Viewname = previewpath & "pre" & File_name & ".jpg"
					Call CreateView(FileName,F_Viewname)
				End If
			'记录文件
			Call Checksave()
			UpCount=UpCount+1
			End If
			Set File=Nothing
		Next
	End If
	Set upload=Nothing
	Call Suc_upload(UpCount,upNum)
End Sub

'保存上传数据并返回附件ID
Private sub checksave()
	Dim Rs,DownloadID,UpFileID,shwofilename
	shwofilename=Replace(Filename,CheckFolder,"UploadFile/")
	If upload_ViewType<>999 and F_Type=1 then
		Dvbbs.execute("insert into dv_upFile (F_BoardID,F_UserID,F_Username,F_Filename,F_Viewname,F_FileType,F_Type,F_FileSize,F_Flag) values ("&Dvbbs.BoardID&","&Dvbbs.UserID&",'"&Dvbbs.membername&"','"&replace(rename,"|","")&"','"&F_Viewname&"','"&replace(FileExt,".","")&"',"&F_Type&","&Filesize&",4)")
	Else
		Dvbbs.execute("insert into dv_upFile (F_BoardID,F_UserID,F_Username,F_Filename,F_FileType,F_Type,F_FileSize,F_Flag) values ("&Dvbbs.BoardID&","&Dvbbs.UserID&",'"&Dvbbs.membername&"','"&replace(rename,"|","")&"','"&replace(FileExt,".","")&"',"&F_Type&","&Filesize&",4)")
	End If
	Set Rs=Dvbbs.execute("Select top 1 F_ID from dv_upFile order by F_ID desc")
		DownloadID=rs(0)
		UpFileID=DownloadID & ","
	Set Rs=nothing

	If F_Type=1 or F_Type=2 then
		Response.write "<script>parent.Dvbbs_Composition.document.body.innerHTML+='[upload="&FileExt&"]"&shwofilename&"[/upload]<br>'</script>"
	Else
		Response.write "<script>parent.Dvbbs_Composition.document.body.innerHTML+='[upload="&FileExt&"]viewFile.asp?ID="&DownloadID&"[/upload]<br>'</script>"
	End If
	Response.write "<script>parent.Dvform.upfilerename.value+='"&UpFileID&"'</script>"
	upNum	=	upNum+1
	Response.cookies("upNum")=upNum
End sub

Private Sub Suc_upload(UpCount,upNum)
	REM 修正批量上传累加个数错误 2004-5-14 Dv.Yz
	If upNum < Clng(Dvbbs.GroupSetting(40)) And Dateupnum+UpCount < Clng(Dvbbs.GroupSetting(50)) Then
		Response.Write UpCount & "个文件上传成功,目前今天总共上传了" & Dateupnum+UpCount & "个附件 [ <a href=post_upload.asp?boardid=" & Dvbbs.BoardID & ">继续上传</a> ]"
	Else
		Response.write UpCount & "个文件上传成功!本次已达到上传数上限。"
	End If
	Dvbbs.Execute("UPDATE [Dv_user] SET UserToday = '" & Dvbbs.UserToday(0) & "|" & Dvbbs.UserToday(1) & "|" & Dvbbs.UserToday(2)+UpCount & "' WHERE UserID = " & Dvbbs.UserID & "")
	Dim iUserInfo
	iUserInfo = Session(Dvbbs.CacheName & "UserID")
	iUserInfo(36) = Dvbbs.UserToday(0) & "|" & Dvbbs.UserToday(1) & "|" & Dvbbs.UserToday(2)+UpCount
	Session(Dvbbs.CacheName & "UserID") = iUserInfo
End Sub


'创建预览图片:call CreateView(原始文件的路径,预览文件名及路径)
Sub CreateView(imagename,tempFilename)
	'定义变量
	Dim PreviewImageFolderName
	Dim ogvbox,objFont
	Dim Logobox,LogoPath
	LogoPath = Server.MapPath("images") & "\logo.gif"  '//加入图片所在路径及文件名

	Select Case upload_ViewType
	Case 0
	'---------------------CreatePreviewImage---------------
		set ogvbox = Server.CreateObject("CreatePreviewImage.cGvbox")
		ogvbox.SetSavePreviewImagePath=Server.MapPath(tempFilename)			'预览图存放路径
		ogvbox.SetPreviewImageSize =SetPreviewImageSize						'预览图宽度
		ogvbox.SetImageFile = trim(Server.MapPath(imagename))				'imagename原始文件的物理路径
		'创建预览图的文件
		If ogvbox.DoImageProcess=false Then
		Response.write "生成预览图错误:"& ogvbox.GetErrString
		End If
	Case 1
	'---------------------AspJpegV1.2---------------
		
		'Set Logobox = Server.CreateObject("Persits.Jpeg")
		'*添加水印图片	添加时请关闭水印字体*
		'//读取添加的图片
		'Logobox.Open LogoPath
		'//重新设置图片的大小
		'Logobox.Width = 180		'// 加入图片的原宽度
		'Logobox.Height = 60		'// 加入图片的原高度
		'*添加水印图片*

		Set ogvbox = Server.CreateObject("Persits.Jpeg")
		' 读取要处理的原文件
		ogvbox.Open Trim(Server.MapPath(imagename))
		If ogvbox.OriginalWidth<Cint(ImageWidth) or ogvbox.Originalheight<Cint(ImageHeight) Then
			F_Viewname=""
			Set ogvbox = Nothing
			Exit Sub
		Else
			IF ImageMode<>"" and FileExt<>"gif" Then
				'//关于修改字体及文字颜色的
				ogvbox.Canvas.Font.Color	= &HFF0000		'// 文字的颜色
				ogvbox.Canvas.Font.Family	= "monospace"	'// 文字的字体
				'ogvbox.Canvas.Font.Bold = True
				' Draw frame: black, 2-pixel width
				ogvbox.Canvas.Print 10, 10, ImageMode		'// 加入文字的位置坐标
				ogvbox.Canvas.Pen.Color		= &H000000		'// 边框的颜色
				ogvbox.Canvas.Pen.Width		= 1				'// 边框的粗细
				ogvbox.Canvas.Brush.Solid	= False			'// 图片边框内是否填充颜色
				'ogvbox.DrawImage 0, 0, Logobox				'// 加入图片的位置价坐标（添加水印图片）
				ogvbox.Canvas.Bar 0, 0, ogvbox.Width, ogvbox.Height	'// 图片边框线的位置坐标
				ogvbox.Save Server.MapPath(imagename)		'// 生成文件
			End If
			ogvbox.Width	= ImageWidth
			ogvbox.height	= ImageHeight
			'ogvbox.height	= ogvbox.Originalheight*ImageWidth\ogvbox.OriginalWidth
			ogvbox.Sharpen 1, 120
			ogvbox.Save Server.MapPath(tempFilename)		'// 生成预览文件
		End If
		Set Logobox=Nothing
	Case 2
	'---------------------SoftArtisans ImgWriter V1.21---------------
		Set ogvbox = Server.CreateObject("SoftArtisans.ImageGen")
		' 读取要处理的原文件
		ogvbox.LoadImage Trim(Server.MapPath(imagename))
		If ogvbox.ErrorDescription <> "" Then
			Response.Write ogvbox.ErrorDescription
		End If
		If ogvbox.Width<Cint(ImageWidth) or ogvbox.Height<Cint(ImageHeight) Then
			F_Viewname=""
			Set ogvbox = Nothing
			Exit Sub
		Else
			IF ImageMode<>"" and FileExt<>"gif" Then
				ogvbox.Font.Italic	= True
				ogvbox.Font.height	= 15
				ogvbox.Font.name	= "monospace"
				ogvbox.Font.Color	= vbred
				ogvbox.Text			=ImageMode
				ogvbox.DrawTextOnImage 10, 10, ogvbox.TextWidth, ogvbox.TextHeight
				ogvbox.SaveImage 0, ogvbox.ImageFormat, Server.MapPath(imagename) 
				'ogvbox.AddWatermark Server.MapPath(Request.QueryString("mimg")), 0, 0.3
			End If
			'ogvbox.SharpenImage 100
			ogvbox.ColorResolution = 24	'24色保存
			ogvbox.ResizeImage ImageWidth,ImageHeight,0,0
			'0=saiFile,1=saiMemory,2=saiBrowser,4=saiDatabaseBlob
			'saiBMP=1,saiGIF=2,saiJPG=3,saiPNG=4,saiPCX=5,saiTIFF=6,saiWMF=7,saiEMF=8,saiPSD=9 
			ogvbox.SaveImage 0, 3, Server.MapPath(tempFilename)
			Response.Write Server.MapPath(tempFilename)
		End If
	Case 3
	'---------------------三角猫生成缩略图组件 SJCatSoft V2.6---------------
		Set ogvbox = Server.CreateObject("sjCatSoft.Thumbnail")
		ogvbox.SourceFile = Trim(Server.MapPath(imagename))
		IF ogvbox.OriginalWidth<Cint(ImageWidth) or ogvbox.OriginalHeight<Cint(ImageHeight) Then
			F_Viewname=""
			Set ogvbox = Nothing
			Exit Sub
		Else
			ogvbox.ByRatio			= False
			ogvbox.OutFileType		= 1
			ogvbox.OutPicWidth		= ImageWidth
			ogvbox.OutPicHeight		= ImageHeight
			ogvbox.DestFile			= Server.MapPath(tempFilename)
			ogvbox.Execute
			IF ImageMode<>"" and FileExt<>"gif" Then
			ogvbox.WaterMaskText	= ImageMode
			ogvbox.FontName			= "monospace"
			ogvbox.FontSize			= 12
			ogvbox.FontColor		= 13
			ogvbox.FontType			= 5
			ogvbox.ByRatio			= True
			ogvbox.Rate				= 100
			ogvbox.DestFile			= Server.MapPath(imagename)
			ogvbox.Execute
			End If
		End If
	End Select
	Set ogvbox = Nothing
End Sub

'按月份自动明名上传文件夹,需要ＦＳＯ组件支持。
Private Function CreatePath()
	Dim objFSO,Fsofolder,uploadpath
	uploadpath=year(now)&"-"&month(now)	'以年月创建上传文件夹，格式：2003－8
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Server.MapPath(CheckFolder&uploadpath))=False Then
			objFSO.CreateFolder Server.MapPath(CheckFolder&uploadpath)
		End If
		If Err.Number = 0 Then
			CreatePath=uploadpath&"/"
		Else
			CreatePath=""
		End If
	Set objFSO = Nothing
End Function

'读取上传目录
Function CheckFolder()
	If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
	CheckFolder = Replace(Replace(Dvbbs.Forum_Setting(76),Chr(0),""),".","")
	'在目录后加(/)
	If Right(CheckFolder,1)<>"/" Then CheckFolder=CheckFolder&"/"
End Function

'文件明名
Private Function CreateName()
	Dim ranNum
	Randomize
	ranNum=int(999*rnd)
	CreateName=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
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

'判断文件类型是否合格
Private Function CheckFileExt(FileExt)
	Dim Forumupload,i
	CheckFileExt=False
	If FileExt="" or IsEmpty(FileExt) Then
		CheckFileExt=False
		Exit Function
	End If
	If FileExt="asp" or FileExt="asa" or FileExt="aspx" Then
		CheckFileExt=False
		Exit Function
	End If
	Forumupload=Split(Dvbbs.Board_Setting(19),"|")
	For i=0 To ubound(Forumupload)
		If FileExt=Lcase(Trim(Forumupload(i))) Then
			CheckFileExt=True
			Exit Function
		Else
			CheckFileExt=False
		End If
	Next
End Function

'判断文件类型:0=其它,1=图片,2=FLASH,3=音乐,4=电影
Private Function CheckFiletype(FileExt)
	Dim upFiletype
	Dim FilePic,FileVedio,FileSoft,FileFlash,FileMusic
	FileExt=Lcase(Replace(FileExt,".",""))
	Select Case Lcase(FileExt)
			Case "gif", "jpg", "jpeg","png","bmp","tif","iff"
				CheckFiletype=1
			Case "swf", "swi"
				CheckFiletype=2
			Case "mid", "wav", "mp3","rmi","cda"
				CheckFiletype=3
			Case "avi", "mpg", "mpeg","ra","ram","wov","asf"
				CheckFiletype=4
			Case Else
				CheckFiletype=0
	End Select
End Function

'常见文件的MIME类型
'GIF文件  "image/gif"
'BMP文件 "image/bmp"
'JPG文件 "image/jpeg"
'PNG文件 "IMAGE/X-PNG"
'zip文件 "application/x-zip-compressed"
'DOC文件 "application/msword"
'文本文件 "text/plain"
'HTML文件 "text/html"
'一般文件 "application/octet-stream"

'SoftArtisans.ImageGen
'ogvbox.AddWatermark Watermark,Position,Opacity,TransitionColor,ShrinkToFit
'Position:
'saiTopMiddle 0  
'saiCenterMiddle 1  
'saiBottomMiddle 2  
'saiTopLeft 3  
'saiCenterLeft 4  
'saiBottomLeft 5  
'saiTopRight 6  
'saiCenterRight 7  
'saiBottomRight 8 
'ShrinkToFit:自动适中（默认为：TRUE）
%>
</td></tr>
</table>
</body>
</html>