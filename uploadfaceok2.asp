<!--#include file=inc/conn.asp -->
<!--#include file=inc/function.asp -->
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>文档上传</title>
<LINK href=site.css rel=stylesheet>
</head>
<body>
<%
'上传方法，0＝无组件，1＝chinaaspupload
dim upload_type
upload_type=0

dim uploadsuc
dim Forumupload
dim ranNum
dim uploadfilestype
dim upload,file,formName,formPath,iCount,filename,fileExt
response.write "<body leftmargin=0 topmargin=0>"
select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case else
	response.write "本系统未开放插件功能"
	response.end
end select

sub upload_0()
set upload=new upload_5xSoft ''建立上传对象

 formPath=upload.form("filepath")
 ''在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

response.write "<body leftmargin=5 topmargin=3>"

for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.filesize<1024 then
 	response.write "<span style=""font-family: 宋体; font-size: 9pt"">请先选择你要上传的文档　[ <a href=# onclick=history.go(-1)>继续上传</a> ]</span>"
	response.end
 end if
 	

 if file.filesize>2024000 then
 	response.write "<span style=""font-family: 宋体; font-size: 9pt"">文档大小超过了限制　[ <a href=# onclick=history.go(-1)>继续上传</a> ]</span>"
	response.end
 end if
 
 fileExt=lcase(getFileExtName(file.fileName))
 
 if fileext<>"jpg" and fileext<>"bmp" and fileext<>"gif" then
 	response.write "<span style=""font-family: 宋体; font-size: 9pt"">该文件格式不允许上传　[ <a href=# onclick=history.go(-1)>继续上传</a> ]</span>"
	response.end
 end if
 
randomize
rannum=int(90000*rnd)+10000
filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&rannum&"."&fileExt
%>
<%
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveAs Server.mappath(formPath &filename)   ''保存文件
%> <script>parent.document.myform.Document2.value="uploadfiles/<%=FileName%>"</script>
 <%end if
 set file=nothing
 set file=nothing
next
set upload=nothing

Htmend iCount&" 个文件上传结束!"
end sub

sub HtmEnd(Msg)
  response.write "<span style=""font-family: 宋体; font-size: 9pt"">文档上传成功 [ <a href=# onclick=history.go(-1)>继续上传</a> ]</span>"
response.end
end sub

sub upload_1()
set FileUp=server.createobject("ChinaASP.UpLoad") ''建立上传对象

filepath=server.MapPath("uploadfiles/")

response.write "<body leftmargin=5 topmargin=3>"
for each f in fileup.Files ''列出所有上传了的文件

 if f.filesize<100 then
 	response.write "<span style=""font-family: 宋体; font-size: 9pt"">请先选择你要上传的文档　[ <a href=# onclick=history.go(-1)>继续上传</a> ]</span>"
	response.end
 end if
 	

 if f.filesize>3024000 then
 	response.write "<span style=""font-family: 宋体; font-size: 9pt"">文档大小超过了限制　[ <a href=# onclick=history.go(-1)>继续上传</a> ]</span>"
	response.end
 end if
 
 fileExt=lcase(getFileExtName(f.fileName)) 
 
 if fileext<>"jpg" and fileext<>"bmp" and fileext<>"gif" and fileext<>"txt" then
 	response.write "<span style=""font-family: 宋体; font-size: 9pt"">该文件格式不允许上传　[ <a href=# onclick=history.go(-1)>继续上传</a> ]</span>"
	response.end
 end if

randomize
rannum=int(90000*rnd)+10000
filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&rannum&"."&fileExt
%>
<%
 if f.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  f.saveas filePath & "\"&filename   ''保存文件
 response.write "<script>parent.document.myform.Document1.value='uploadfiles/"&FileName&"'</script>"
 end if
 set f=nothing
next
set FileUp=nothing

Htmend iCount&" 个文件上传结束!"
end sub

sub HtmEnd(Msg)
  response.write "<span style=""font-family: 宋体; font-size: 9pt"">文档上传成功 [ <a href=# onclick=history.go(-1)>继续上传</a> ]</span>"

response.end
end sub


%>
</body>
</html>