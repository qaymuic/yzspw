<!--#include file = "conn.asp"-->
<!-- #include file = "inc/const.asp" -->
<!-- #include file = "inc/DvADChar.asp" -->
<%
Head()
dim admin_flag
dim objFSO
dim uploadfolder
dim uploadfiles
dim upname
dim uid,faceid
dim usernames
dim userface,dnum
dim upfilename
dim pagesize, page,filenum, pagenum
admin_flag = ",34,"
If not Dvbbs.master or instr(","&session("flag")&",",admin_flag) = 0 Then
	Errmsg = ErrMsg + "<BR><li>本页面为管理员专用，请<a href = admin_index.asp target = _top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	dvbbs_error()
else
	call main()
	Footer()
End If

sub main()
%>
<table width = "95%" border = "0" cellspacing = "1" cellpadding = "3"  align = center>
<tr>
<td valign = top>
注意：本功能需要主机开放FSO权限，FSO相关帮助请看微软帮助文档<BR>
在这里您可以管理论坛所有用户自定义头像上传文件，搜索用户头像请用用户ID进行搜索<BR>
用户ID的获得可以通过用户信息管理中搜索相关用户，然后将鼠标移到用户名连接上，查看连接属性，参数UserID = 后面既是用户的ID
</td>
</tr>
</table>
<table width = "95%" border = "0" cellspacing = "1" cellpadding = "3"  align = center class = "tableBorder" style = "table-layout:fixed;word-break:break-all">
<tr align = center><th width = "*" height = 25>文件名</th><th width = "100">所属用户</th><th width = "50">大小</th><th width = "120">最后访问</th><th width = "120">上传日期</th><th width = "35">管理</th></tr>
<form method="POST" action="?action=delall">
<%
pagesize = 20
page = request.querystring("page")
If page = "" or not isnumeric(page) Then
	page = 1
Else
	page = int(page)
End If

If trim(request("action"))<>"" Then
	If trim(request("action")) = "delall" Then
		call delface()
	Else
		call maininfo()
	End If
Else
	call maininfo()
End If
call foot()
End Sub

sub maininfo()
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
If request("filename")<>"" Then
objFSO.DeleteFile(Server.MapPath("uploadFace\"&request("filename")))
End If
Set uploadFolder = objFSO.GetFolder(Server.MapPath("uploadFace\"))
Set uploadFiles = uploadFolder.Files
filenum = uploadfiles.count
pagenum = int(filenum/pagesize)
If filenum mod pagesize>0 Then
	pagenum = pagenum+1
End If
If page> pagenum Then
	page = 1
End If
i = 0
For Each Upname In uploadFiles
	i = i+1
	If i>(page-1)*pagesize and i <= page*pagesize Then
	upfilename = "uploadFace/"&upname.name
		If instr(upname.name,"_") Then    '取出头像的用户名
			uid = split(upname.name,"_")
			faceid = uid(0)
			If  IsNumeric(faceid)	then
				set rs = Dvbbs.Execute("select username from [dv_user] where   userid = "&faceid&"  ")
				If not rs.eof  Then
					usernames = rs(0)
				End If
				rs.close
				Set rs = Nothing
			End If		
		End If
		response.write "<tr><td class = forumRow height = 23><a href=""uploadface/"&upname.name&""" target=_blank>"&upname.name&"</a></td>"
		response.write "<td align = right class = forumRowHighlight>"&usernames&"</td>"
		response.write "<td align = right class = forumRow>"& upname.size &"</td>"
		response.write "<td align = center class = forumRowHighlight>"& upname.datelastaccessed &"</td>"
		response.write "<td align = center class = forumRow>"& upname.datecreated &"</td>"
		response.write "<td align = center class = forumRowHighlight><a href = '?filename="&upname.name&"'>删除</a></td></tr>"
	ElseIf i>page*pagesize Then
		Exit For
	End If
	usernames = ""
Next

End Sub 

'清理头像
Sub delface()
Dim DllUserFace
dnum = 0
DllUserFace = Request("filename")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
'删除返还头像参数的文件;
If DllUserFace<>"" Then
	DllUserFace = Replace(DllUserFace,"..","")
	objFSO.DeleteFile(Server.MapPath("uploadFace/"&DllUserFace))
End If
Set uploadFolder = objFSO.GetFolder(Server.MapPath("uploadFace/"))
Set uploadFiles = uploadFolder.Files
filenum = uploadfiles.count
pagenum = int(filenum/pagesize)
If filenum mod pagesize>0 Then
	pagenum = pagenum+1
End If
If page> pagenum Then
	page = 1
End If
i = 0
For Each Upname In uploadFiles
i = i+1
If i>(page-1)*pagesize and i <= page*pagesize Then
upfilename = "uploadFace/"&upname.name
	'取出头像的用户名
	If instr(upname.name,"_") Then    
		uid = split(upname.name,"_")
		faceid = uid(0)
		If IsNumeric(faceid) Then
			Set rs = Dvbbs.Execute("select username,userface from [dv_user] where userid = "& faceid)
			If not rs.eof Then
			usernames = rs(0)
			userface = trim(rs(1))
				If instr(upfilename,userface) = 0 Then
					objFSO.DeleteFile(Server.MapPath(upfilename))
					Response.Write "头像已更改,用户"& usernames &"旧头像文件："& upfilename &"已删除<br>"
					dnum = dnum+1
				End If
			Else
				objFSO.DeleteFile(Server.MapPath(upfilename))
				response.write "用户"& uid(1) &"已注销,文件："& upfilename &"已删除<br>"
				dnum = dnum+1
			End If
			rs.close
			set rs = nothing
		End If
	Else
	'清理没有用户ID的头像文件
		sql = "select top 1 userid from [dv_user] where userface = '"& upfilename &"' "
		Set rs = Dvbbs.Execute(sql)
		If rs.eof Then
			objFSO.DeleteFile(Server.MapPath(upfilename))
			response.write "已清查删除文件："& upfilename &"<br>"
			dnum = dnum+1
		End If
		rs.close
		Set rs = nothing
	End If
ElseIf i>page*pagesize Then
	Exit For
End If
Next
response.write " 共清理 "& dnum &" 个文件  "
End Sub

Sub foot()
Set uploadFolder = Nothing
Set uploadFiles = Nothing
%>
<tr><td colspan=6 class=forumRow height=30>
<%
If page>1 Then
	response.write "<a href=?page=1>首页</a>&nbsp;&nbsp;<a href=""?page="& page-1 &""">上一页</a>&nbsp;&nbsp;"
Else
	response.write "首页&nbsp;&nbsp;上一页&nbsp;&nbsp;"
End If
If page<i/pagesize Then
	response.write "<a href=""?page="& page+1 &""">下一页</a>&nbsp;&nbsp;<a href=""?page="& pagenum &""">尾页</a>"
Else
	response.write "下一页&nbsp;&nbsp;尾页"
End If
%>
<input type="submit" value="清理"></td><tr></form></table><br>
<% End Sub %>
