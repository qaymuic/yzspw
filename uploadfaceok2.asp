<!--#include file=inc/conn.asp -->
<!--#include file=inc/function.asp -->
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>�ĵ��ϴ�</title>
<LINK href=site.css rel=stylesheet>
</head>
<body>
<%
'�ϴ�������0���������1��chinaaspupload
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
	response.write "��ϵͳδ���Ų������"
	response.end
end select

sub upload_0()
set upload=new upload_5xSoft ''�����ϴ�����

 formPath=upload.form("filepath")
 ''��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

response.write "<body leftmargin=5 topmargin=3>"

for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<1024 then
 	response.write "<span style=""font-family: ����; font-size: 9pt"">����ѡ����Ҫ�ϴ����ĵ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
	response.end
 end if
 	

 if file.filesize>2024000 then
 	response.write "<span style=""font-family: ����; font-size: 9pt"">�ĵ���С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
	response.end
 end if
 
 fileExt=lcase(getFileExtName(file.fileName))
 
 if fileext<>"jpg" and fileext<>"bmp" and fileext<>"gif" then
 	response.write "<span style=""font-family: ����; font-size: 9pt"">���ļ���ʽ�������ϴ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
	response.end
 end if
 
randomize
rannum=int(90000*rnd)+10000
filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&rannum&"."&fileExt
%>
<%
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveAs Server.mappath(formPath &filename)   ''�����ļ�
%> <script>parent.document.myform.Document2.value="uploadfiles/<%=FileName%>"</script>
 <%end if
 set file=nothing
 set file=nothing
next
set upload=nothing

Htmend iCount&" ���ļ��ϴ�����!"
end sub

sub HtmEnd(Msg)
  response.write "<span style=""font-family: ����; font-size: 9pt"">�ĵ��ϴ��ɹ� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
response.end
end sub

sub upload_1()
set FileUp=server.createobject("ChinaASP.UpLoad") ''�����ϴ�����

filepath=server.MapPath("uploadfiles/")

response.write "<body leftmargin=5 topmargin=3>"
for each f in fileup.Files ''�г������ϴ��˵��ļ�

 if f.filesize<100 then
 	response.write "<span style=""font-family: ����; font-size: 9pt"">����ѡ����Ҫ�ϴ����ĵ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
	response.end
 end if
 	

 if f.filesize>3024000 then
 	response.write "<span style=""font-family: ����; font-size: 9pt"">�ĵ���С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
	response.end
 end if
 
 fileExt=lcase(getFileExtName(f.fileName)) 
 
 if fileext<>"jpg" and fileext<>"bmp" and fileext<>"gif" and fileext<>"txt" then
 	response.write "<span style=""font-family: ����; font-size: 9pt"">���ļ���ʽ�������ϴ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
	response.end
 end if

randomize
rannum=int(90000*rnd)+10000
filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&rannum&"."&fileExt
%>
<%
 if f.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  f.saveas filePath & "\"&filename   ''�����ļ�
 response.write "<script>parent.document.myform.Document1.value='uploadfiles/"&FileName&"'</script>"
 end if
 set f=nothing
next
set FileUp=nothing

Htmend iCount&" ���ļ��ϴ�����!"
end sub

sub HtmEnd(Msg)
  response.write "<span style=""font-family: ����; font-size: 9pt"">�ĵ��ϴ��ɹ� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"

response.end
end sub


%>
</body>
</html>