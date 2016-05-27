<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=5    '操作权限
%>
<!--#include file="ChkPurview.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim rs,sql,ErrMsg,FoundErr
dim id,Title,Content,Author,Hits,SmallClassName,bigclassname
dim IncludePic,DefaultPicUrl,UploadFiles,arrUploadFiles,istop
dim ObjInstalled
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
FoundErr=false
id=Trim(Request.Form("id"))
SmallClassName=Trim(Request.Form("SmallClassName"))
bigclassname=Trim(Request.Form("bigclassname"))
Title=trim(request.form("Title"))
istop=trim(request.form("istop"))
Content=trim(request.form("Content"))
IncludePic=trim(request.form("IncludePic"))
DefaultPicUrl=trim(request.form("DefaultPicUrl"))
UploadFiles=trim(request.form("UploadFiles"))
Author=session("admin")

if Title="" then
	founderr=true
	errmsg="<li>内容标题不能为空</li>"
end if
if Content="" then
	founderr=true
	errmsg=errmsg+"<li>内容内容不能为空</li>"
end if

if founderr=false then
	Title=dvhtmlencode(Title)
	Content=ubbcode(Content)
	set rs=server.createobject("adodb.recordset")
	if request("action")="add" then
		sql="select top 1 * from ytiinews" 
		rs.open sql,conn,1,3
		rs.addnew
		call SaveData()
		rs.update
		id=rs("id")
		rs.close
		set rs=nothing
	elseif request("action")="Modify" then
  		if id<>"" then
			sql="select * from ytiinews where id=" & id
			rs.open sql,conn,1,3
			if not (rs.bof and rs.eof) then
				call SaveData()
				rs.update
				rs.close
				set rs=nothing
 			else
				founderr=true
				errmsg=errmsg+"<li>找不到此内容，可能已经被其他人删除。</li>"
				call WriteErrMsg()
			end if
		else
			founderr=true
			errmsg=errmsg+"<li>不能确定id的值</li>"
			call WriteErrMsg()
		end if
	else
		founderr=true
		errmsg=errmsg+"<li>没有选定参数</li>"
		call WriteErrMsg()
	end if

	call CloseConn()
%>
<html>
<head>
<title></title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>
<div align="center">
<br><br>
<table class="border" align=center width="50%" border="0" cellpadding="4" cellspacing="0" bordercolor="#999999">
  <tr align=center>
    <td width="100%" class="title"  height="20"><b>
<%if request("action")="add" then%>添加<%else%>修改<%end if%>内容成功</b></td>
  </tr>
  <tr>
    <td class="tdbg"><p align="left">
        <p align="center">【<a href="ArticleModify.asp?id=<%=id%>">修改本文</a>】&nbsp;【<a href="ArticleAdd.asp">继续添加内容</a>】&nbsp;【<a href="ArticleManage.asp">内容管理</a>】&nbsp;</p></td>
  </tr>
</table>
</div>
</body>
</html>
<%
else
	WriteErrMsg
end if

sub SaveData()
	rs("Title")=Title
	rs("Content")=Content
	rs("Author")=Author
	rs("bigclassname")=bigclassname
	rs("smallclassname")=smallclassname
	if IncludePic="yes" then
		rs("IncludePic")=True
	else
		rs("IncludePic")=False
	end if
	if istop="yes" then
		rs("istop")=True
	else
		rs("istop")=False
	end if
	'***************************************
	'删除无用的上传文件
	if ObjInstalled=True and UploadFiles<>"" then
		dim fso,strRubbishFile
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if instr(UploadFiles,"|")>1 then
			dim arrUploadFiles,intTemp
			arrUploadFiles=split(UploadFiles,"|")
			UploadFiles=""
			for intTemp=0 to ubound(arrUploadFiles)
				if instr(Content,arrUploadFiles(intTemp))<=0 and arrUploadFiles(intTemp)<>DefaultPicUrl then
					strRubbishFile=server.MapPath("../" & arrUploadFiles(intTemp))
					if fso.FileExists(strRubbishFile) then
						fso.DeleteFile(strRubbishFile)
						response.write "<br><li>" & arrUploadFiles(intTemp) & "在内容中没有用到，也没有被设为首页图片，所以已经被删除！</li>"
					end if
				else
					if intTemp=0 then
						UploadFiles=arrUploadFiles(intTemp)
					else
						UploadFiles=UploadFiles & "|" & arrUploadFiles(intTemp)
					end if
				end if
			next
		else
			if instr(Content,UploadFiles)<=0 and UploadFiles<>DefaultPicUrl then
				strRubbishFile=server.MapPath("../" & UploadFiles)
				if fso.FileExists(strRubbishFile) then
					fso.DeleteFile(strRubbishFile)
					response.write "<br><li>" & UploadFiles & "在内容中没有用到，也没有被设为首页图片，所以已经被删除！</li>"
				end if
				UploadFiles=""
			end if
		end if
		set fso=nothing
	end If
	'结束
	'***************************************
	rs("DefaultPicUrl")=DefaultPicUrl
	rs("UploadFiles")=UploadFiles
end sub
	
%>