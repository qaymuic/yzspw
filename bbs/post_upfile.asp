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
Server.ScriptTimeOut=999999'Ҫ�������̳֧���ϴ����ļ��Ƚϴ󣬾ͱ������á�
Dim upload_type,upload_ViewType
'-----------------------------------------------------------------------------
'�����ϴ���ʽupload_typeֵ�� 0���������1��lyfupload 1.2�棬2��Aspupload3.0��3��SA-FileUp 4.0��4=DvFile.Upload V1.0
'-----------------------------------------------------------------------------
upload_type=Cint(Dvbbs.Forum_Setting(43))
'-----------------------------------------------------------------------------
'��������Ԥ��ͼƬ,��ҪͼƬ��д���֧��.(��Ŀ¼��Ҫ��PreviewImage�ļ��д���ļ�)
'����֧�����upload_ViewTypeֵ�� 
'0�� CreatePreviewImage�� 1�� AspJpeg ��2=SoftArtisans ImgWriter V1.21��3=SJCatSoft V2.6 
Dim previewpath,F_Viewname
F_Viewname=""
previewpath="PreviewImage/"
upload_ViewType=Cint(Dvbbs.Forum_Setting(45))
'-----------------------------------------------------------------------------
'�ύ��֤
If Not Dvbbs.ChkPost Then
	Response.End
End If
If Dvbbs.Userid=0 Then
	Response.write "�㻹δ��½��"
	Response.End
End If

DateUpNum=Clng(Dvbbs.UserToday(2))
UpNum=request.cookies("upNum")
If UpNum ="" then UpNum=0
UpNum=int(UpNum)

If Cint(Dvbbs.GroupSetting(7))=0 then
	Response.write "��û���ڱ���̳�ϴ��ļ���Ȩ��"
	Response.End
End If
If upNum >= Clng(Dvbbs.GroupSetting(40)) then
 	Response.write "һ��ֻ���ϴ�"&Dvbbs.GroupSetting(40)&"���ļ���"
	Response.End
End If
If dateupnum+upNum >= Clng(Dvbbs.GroupSetting(50)) then
 	Response.write "�������ϴ����ļ��ѳ�����"&Dvbbs.GroupSetting(50)&"����"
	Response.End
End If
'-----------------------------------------------------------------------------
'�������
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
FormPath=CheckFolder&CreatePath()	'�ϴ�Ŀ¼·��

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
		Response.write "��ϵͳδ���Ų������"
		Response.End
End Select

'===========================������ϴ�============================
Sub upload_0()
	Dim upload,File,UpCount
	UpCount=0
	Set upload = new UpFile_Class						''�����ϴ�����
	Upload.InceptFileType = Replace(Dvbbs.Board_Setting(19),"|",",")
	Upload.MaxSize = Int(Dvbbs.GroupSetting(44))*1024
	Upload.GetDate ()	'ȡ���ϴ�����
	If upload.err > 0 then
		Select Case upload.err
		Case 1
			Response.write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		Case 2
			Response.write "�ļ���С���������� "&Dvbbs.GroupSetting(44)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		Case 3
			Response.write "�ļ����Ͳ���ȷ "&Dvbbs.GroupSetting(44)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		End Select
		Exit Sub
	Else
		For Each FormName In upload.File		''�г������ϴ��˵��ļ�
			If upNum >= Int(Dvbbs.GroupSetting(40)) or dateupnum+upNum >= clng(Dvbbs.GroupSetting(50)) then
				Response.write "�Ѵﵽ�ϴ��������ޡ�"
				EXIT SUB
			End If
			Set File=upload.File(FormName)		''����һ���ļ�����
			FileExt=FixName(File.FileExt)
			'�ж��ļ�����
			If CheckFileExt(FileExt)=false then
				Response.write "�ļ���ʽ����ȷ,����Ϊ�ա�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				EXIT SUB
			End If
			'��ֵ����
			F_Type		=	CheckFiletype(FileExt)
			File_name	=	CreateName()
			Filename	=	File_name&"."&FileExt
			rename		=	CreatePath&Filename&"|"
			Filename	=	FormPath&Filename
			Filesize	=	File.FileSize
			'��¼�ļ�
			If Filesize>0 then								'��� FileSize > 0 ˵�����ļ�����
			File.SaveToFile Server.mappath(FileName)		''ִ���ϴ��ļ�
			'��������Ԥ��ͼƬ
				If upload_ViewType<>999 and F_Type=1 then
					F_Viewname=previewpath&"pre"&File_name&".jpg"
					Call CreateView(FileName,F_Viewname)
				End If
			'��¼�ļ�
			Call checksave()
			UpCount=UpCount+1
			End If
			Set File=Nothing
		Next
	End If
	Set upload=Nothing
	Call Suc_upload(UpCount,upNum)
End Sub

''===========================lyfupload����ϴ�1.2��============================
Sub upload_1()
	Dim obj,Filepath,FileExt_a,UpCount
	Dim ss,i
	Dim TempExt,TempFileValue
	UpCount=0
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'���ƴ�С
	obj.maxsize=Int(Dvbbs.GroupSetting(44))*1024
	'��������
	obj.extname=Replace(Dvbbs.Board_Setting(19),"|",",")
	Filepath=Server.MapPath(FormPath)
	'��Ŀ¼���(/)
	If Right(Filepath,1)<>"\" Then Filepath=Filepath&"\"
	For i=1 to obj.Request("upcount")
	 TempFileValue	=	"file"&i
		FileExt_a	=	Split(obj.Request(TempFileValue),"""")
		TempExt		=	FileExt_a(1)
		If TempExt="" or isnull(TempExt) then
			Response.write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			Exit Sub
		End If
		FileExt		=	Mid(TempExt, InStrRev(TempExt, ".")+1)
		FileExt		=	FixName(FileExt)
		File_name	=	CreateName()
		Filename	=	File_name&"."&FileExt
		rename		=	CreatePath&Filename & "|"
		Filesize	=	obj.Filesize
		'�ж��ļ����ͼ���ֵ����
		F_Type		=CheckFiletype(FileExt)

		'�ж��ļ�����
		If CheckFileExt(FileExt)=False then
			Response.write "�ļ���ʽ����ȷ,����Ϊ�ա�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
			EXIT SUB
		End If
		'�����ļ�
		ss=obj.SaveFile(TempFileValue,Server.MapPath(FormPath), false,Filename)
		If ss= "3" then
			Response.write ("�ļ����ظ�!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
			EXIT SUB
		ElseIf ss= "0" then
			Response.write ("�ļ���С���������� "&Dvbbs.GroupSetting(44)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
			EXIT SUB
		ElseIf ss = "1" then
			Response.write ("�ļ�����ָ�������ļ�!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
			EXIT SUB
		ElseIf ss = "" then
			Response.write ("�ļ��ϴ�ʧ��!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
			EXIT SUB
		Else
			Filename=FormPath&Filename
			'��������Ԥ��ͼƬ
			If upload_ViewType<>999 and F_Type=1 then
				F_Viewname=previewpath&"pre"&File_name&".jpg"
				call CreateView(FileName,F_Viewname)
			End If
			'��¼�ļ�
			call checksave()
			UpCount=UpCount+1
		End If
	Next
	set obj=nothing
	Call Suc_upload(UpCount,upNum)
End Sub

''===========================Aspupload3.0����ϴ�============================
sub upload_2()
	On Error Resume Next
	Dim Upload,File
	Dim FilePath
	Dim Count,UpCount
	UpCount=0
	Set Upload = Server.CreateObject("Persits.Upload") 
	Upload.OverwriteFiles = false								'���ܸ���
	Upload.IgnoreNoPost = True
	Upload.SetMaxSize Int(Dvbbs.GroupSetting(44))*1024, True	'���ƴ�С
	Count = Upload.Save
	If Err.Number = 8 Then 
	   Response.write "�ļ���С���������� "&Dvbbs.GroupSetting(44)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]" 
	Else 
		If Err <> 0 Then 
			Response.write "������Ϣ: " & Err.Description
			EXIT SUB
		Else
			If Count < 1 Then 
				Response.write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				EXIT SUB
			End If
		For Each File in Upload.Files	'�г������ϴ��ļ�
			If upNum >= Int(Dvbbs.GroupSetting(40)) or dateupnum+upNum >= Clng(Dvbbs.GroupSetting(50)) Then
				Response.write "�Ѵﵽ�ϴ��������ޡ�"
				Exit Sub
			End If
			FileExt = Replace(File.ext,".","")
			FileExt = FixName(FileExt)
			'�ж��ļ�����
			If CheckFileExt(FileExt)=False then
				Response.write "�ļ���ʽ����ȷ,����Ϊ�ա�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				EXIT SUB
			End If
			'�ļ�������ֵ
			File_name	=	CreateName()
			Filename	=	File_name&"."&FileExt
			rename		=	CreatePath&Filename & "|"
			Filename	=	FormPath&Filename
			Filesize	=	File.Size
			F_Type		=	CheckFiletype(FileExt)
			File.saveas Server.MapPath(Filename)	'�ϴ������ļ�
			'��������Ԥ��ͼƬ
			If upload_ViewType<>999 and F_Type=1 then
				F_Viewname=previewpath&"pre"&File_name&".jpg"
				Call CreateView(FileName,F_Viewname)
			End If
			'��¼�ļ�
			Call checksave()			'��¼�ļ�
			UpCount=UpCount+1
		Next
		Call Suc_upload(UpCount,upNum)
		End If 
	End If
	Set Upload = Nothing
End Sub

''===========================SA-FileUp 4.0����ϴ�FileUpSE V4.09============================
sub upload_3()
	Dim oFileUp,UpCount,FileExt_a
	UpCount=0
	Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")
	'oFileUp.Path = Server.MapPath(FormPath)
	For Each FormName In oFileUp.Form
		If IsObject(oFileUp.Form(FormName)) Then
			If Not oFileUp.Form(FormName).IsEmpty Then
				oFileUp.Form(FormName).Maxbytes=int(Dvbbs.GroupSetting(44))*1024	'���ƴ�С
				Filesize=oFileUp.Form(FormName).TotalBytes
				If Filesize>int(Dvbbs.GroupSetting(44))*1024 then
					Response.write "�ļ���С���������� "&Dvbbs.GroupSetting(44)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]" 
					Exit sub
				End If
				Filename	= oFileUp.Form(FormName).ShortFileName	 'ԭ�ļ���
				FileExt		= Mid(Filename, InStrRev(Filename, ".")+1)
				FileExt		= FixName(FileExt)
				'�ж��ļ�����
				If CheckFileExt(FileExt)=false then
					Response.write "�ļ���ʽ����ȷ,����Ϊ�ա�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
					Exit Sub
				End If
				'�ļ�������ֵ
				File_name	=	CreateName()
				Filename	=	File_name&"."&FileExt
				rename		=	CreatePath&Filename & "|"
				Filename	=	FormPath&Filename
				F_Type		=	CheckFiletype(FileExt)
				
				'�����ļ�
				oFileUp.Form(FormName).Saveas Server.MapPath(Filename)
				'��������Ԥ��ͼƬ
				If upload_ViewType<>999 and F_Type=1 then
					F_Viewname=previewpath&"pre"&File_name&".jpg"
					Call CreateView(FileName,F_Viewname)
				End If
				'��¼�ļ�
				Call checksave()			'��¼�ļ�
				UpCount	= UpCount+1
			Else
				Response.write "����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				EXIT SUB
			End If
		End If
	Next
	Set oFileUp = Nothing
	Call Suc_upload(UpCount,upNum)
End Sub

'===========================DvFile.Upload V1.0����ϴ�============================
Sub upload_4()
	Dim upload,File,UpCount
	UpCount=0
	Set upload = Server.CreateObject("DvFile.Upload")			''�����ϴ�����
	Upload.InceptFileType = Replace(Dvbbs.Board_Setting(19),"|",",")
	Upload.MaxSize = Int(Dvbbs.GroupSetting(44))*1024
	Upload.Install		'ȡ���ϴ�����
	If upload.err > 0 Then
		Select Case Upload.Err
			Case 1 : Response.Write Upload.Description	''����ѡ����Ҫ�ϴ����ļ�
			Case 2 : Response.Write Upload.Description & Dvbbs.GroupSetting(44) &"KB"	''ͼƬ��С���������� "&Upload.MaxSize&"K��
			Case 3 : Response.Write Upload.Description	''�Ƿ����ϴ�����
			Case 4 : Response.Write Upload.Description	''���ϴ���������ϵͳ����
			Case 5 : Response.Write Upload.Description	''���������ϴ�������ֹ
		End Select
		Response.Write "��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
		Exit Sub
	Else
		For Each FormName In upload.File		''�г������ϴ��˵��ļ�
			If upNum >= Int(Dvbbs.GroupSetting(40)) or dateupnum+upNum >= Clng(Dvbbs.GroupSetting(50)) then
				Response.write "�Ѵﵽ�ϴ��������ޡ�"
				EXIT SUB
			End If
			Set File = upload.File(FormName)		''����һ���ļ�����
			FileExt = FixName(File.FileExt)
			'�ж��ļ�����
			If CheckFileExt(FileExt)=False then
				Response.write "�ļ���ʽ����ȷ,����Ϊ�ա�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
				EXIT SUB
			End If
			'��ֵ����
			F_Type		=	CheckFiletype(FileExt)
			File_name	=	CreateName()
			Filename	=	File_name&"."&FileExt
			Rename		=	CreatePath&Filename&"|"
			Filename	=	FormPath&Filename
			Filesize	=	File.FileSize
			'��¼�ļ�
			If Filesize>0 then								''��� FileSize > 0 ˵�����ļ�����
			Upload.SaveToFile Server.Mappath(FileName),FormName		''�����ļ�
			'��������Ԥ��ͼƬ
				If upload_ViewType<>999 And F_Type=1 Then
					F_Viewname = previewpath & "pre" & File_name & ".jpg"
					Call CreateView(FileName,F_Viewname)
				End If
			'��¼�ļ�
			Call Checksave()
			UpCount=UpCount+1
			End If
			Set File=Nothing
		Next
	End If
	Set upload=Nothing
	Call Suc_upload(UpCount,upNum)
End Sub

'�����ϴ����ݲ����ظ���ID
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
	REM ���������ϴ��ۼӸ������� 2004-5-14 Dv.Yz
	If upNum < Clng(Dvbbs.GroupSetting(40)) And Dateupnum+UpCount < Clng(Dvbbs.GroupSetting(50)) Then
		Response.Write UpCount & "���ļ��ϴ��ɹ�,Ŀǰ�����ܹ��ϴ���" & Dateupnum+UpCount & "������ [ <a href=post_upload.asp?boardid=" & Dvbbs.BoardID & ">�����ϴ�</a> ]"
	Else
		Response.write UpCount & "���ļ��ϴ��ɹ�!�����Ѵﵽ�ϴ������ޡ�"
	End If
	Dvbbs.Execute("UPDATE [Dv_user] SET UserToday = '" & Dvbbs.UserToday(0) & "|" & Dvbbs.UserToday(1) & "|" & Dvbbs.UserToday(2)+UpCount & "' WHERE UserID = " & Dvbbs.UserID & "")
	Dim iUserInfo
	iUserInfo = Session(Dvbbs.CacheName & "UserID")
	iUserInfo(36) = Dvbbs.UserToday(0) & "|" & Dvbbs.UserToday(1) & "|" & Dvbbs.UserToday(2)+UpCount
	Session(Dvbbs.CacheName & "UserID") = iUserInfo
End Sub


'����Ԥ��ͼƬ:call CreateView(ԭʼ�ļ���·��,Ԥ���ļ�����·��)
Sub CreateView(imagename,tempFilename)
	'�������
	Dim PreviewImageFolderName
	Dim ogvbox,objFont
	Dim Logobox,LogoPath
	LogoPath = Server.MapPath("images") & "\logo.gif"  '//����ͼƬ����·�����ļ���

	Select Case upload_ViewType
	Case 0
	'---------------------CreatePreviewImage---------------
		set ogvbox = Server.CreateObject("CreatePreviewImage.cGvbox")
		ogvbox.SetSavePreviewImagePath=Server.MapPath(tempFilename)			'Ԥ��ͼ���·��
		ogvbox.SetPreviewImageSize =SetPreviewImageSize						'Ԥ��ͼ���
		ogvbox.SetImageFile = trim(Server.MapPath(imagename))				'imagenameԭʼ�ļ�������·��
		'����Ԥ��ͼ���ļ�
		If ogvbox.DoImageProcess=false Then
		Response.write "����Ԥ��ͼ����:"& ogvbox.GetErrString
		End If
	Case 1
	'---------------------AspJpegV1.2---------------
		
		'Set Logobox = Server.CreateObject("Persits.Jpeg")
		'*���ˮӡͼƬ	���ʱ��ر�ˮӡ����*
		'//��ȡ��ӵ�ͼƬ
		'Logobox.Open LogoPath
		'//��������ͼƬ�Ĵ�С
		'Logobox.Width = 180		'// ����ͼƬ��ԭ���
		'Logobox.Height = 60		'// ����ͼƬ��ԭ�߶�
		'*���ˮӡͼƬ*

		Set ogvbox = Server.CreateObject("Persits.Jpeg")
		' ��ȡҪ�����ԭ�ļ�
		ogvbox.Open Trim(Server.MapPath(imagename))
		If ogvbox.OriginalWidth<Cint(ImageWidth) or ogvbox.Originalheight<Cint(ImageHeight) Then
			F_Viewname=""
			Set ogvbox = Nothing
			Exit Sub
		Else
			IF ImageMode<>"" and FileExt<>"gif" Then
				'//�����޸����弰������ɫ��
				ogvbox.Canvas.Font.Color	= &HFF0000		'// ���ֵ���ɫ
				ogvbox.Canvas.Font.Family	= "monospace"	'// ���ֵ�����
				'ogvbox.Canvas.Font.Bold = True
				' Draw frame: black, 2-pixel width
				ogvbox.Canvas.Print 10, 10, ImageMode		'// �������ֵ�λ������
				ogvbox.Canvas.Pen.Color		= &H000000		'// �߿����ɫ
				ogvbox.Canvas.Pen.Width		= 1				'// �߿�Ĵ�ϸ
				ogvbox.Canvas.Brush.Solid	= False			'// ͼƬ�߿����Ƿ������ɫ
				'ogvbox.DrawImage 0, 0, Logobox				'// ����ͼƬ��λ�ü����꣨���ˮӡͼƬ��
				ogvbox.Canvas.Bar 0, 0, ogvbox.Width, ogvbox.Height	'// ͼƬ�߿��ߵ�λ������
				ogvbox.Save Server.MapPath(imagename)		'// �����ļ�
			End If
			ogvbox.Width	= ImageWidth
			ogvbox.height	= ImageHeight
			'ogvbox.height	= ogvbox.Originalheight*ImageWidth\ogvbox.OriginalWidth
			ogvbox.Sharpen 1, 120
			ogvbox.Save Server.MapPath(tempFilename)		'// ����Ԥ���ļ�
		End If
		Set Logobox=Nothing
	Case 2
	'---------------------SoftArtisans ImgWriter V1.21---------------
		Set ogvbox = Server.CreateObject("SoftArtisans.ImageGen")
		' ��ȡҪ�����ԭ�ļ�
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
			ogvbox.ColorResolution = 24	'24ɫ����
			ogvbox.ResizeImage ImageWidth,ImageHeight,0,0
			'0=saiFile,1=saiMemory,2=saiBrowser,4=saiDatabaseBlob
			'saiBMP=1,saiGIF=2,saiJPG=3,saiPNG=4,saiPCX=5,saiTIFF=6,saiWMF=7,saiEMF=8,saiPSD=9 
			ogvbox.SaveImage 0, 3, Server.MapPath(tempFilename)
			Response.Write Server.MapPath(tempFilename)
		End If
	Case 3
	'---------------------����è��������ͼ��� SJCatSoft V2.6---------------
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

'���·��Զ������ϴ��ļ���,��Ҫ�ƣӣ����֧�֡�
Private Function CreatePath()
	Dim objFSO,Fsofolder,uploadpath
	uploadpath=year(now)&"-"&month(now)	'�����´����ϴ��ļ��У���ʽ��2003��8
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

'��ȡ�ϴ�Ŀ¼
Function CheckFolder()
	If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
	CheckFolder = Replace(Replace(Dvbbs.Forum_Setting(76),Chr(0),""),".","")
	'��Ŀ¼���(/)
	If Right(CheckFolder,1)<>"/" Then CheckFolder=CheckFolder&"/"
End Function

'�ļ�����
Private Function CreateName()
	Dim ranNum
	Randomize
	ranNum=int(999*rnd)
	CreateName=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
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

'�ж��ļ������Ƿ�ϸ�
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

'�ж��ļ�����:0=����,1=ͼƬ,2=FLASH,3=����,4=��Ӱ
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

'�����ļ���MIME����
'GIF�ļ�  "image/gif"
'BMP�ļ� "image/bmp"
'JPG�ļ� "image/jpeg"
'PNG�ļ� "IMAGE/X-PNG"
'zip�ļ� "application/x-zip-compressed"
'DOC�ļ� "application/msword"
'�ı��ļ� "text/plain"
'HTML�ļ� "text/html"
'һ���ļ� "application/octet-stream"

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
'ShrinkToFit:�Զ����У�Ĭ��Ϊ��TRUE��
%>
</td></tr>
</table>
</body>
</html>