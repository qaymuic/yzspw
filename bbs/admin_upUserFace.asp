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
	Errmsg = ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href = admin_index.asp target = _top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
ע�⣺��������Ҫ��������FSOȨ�ޣ�FSO��ذ����뿴΢������ĵ�<BR>
�����������Թ�����̳�����û��Զ���ͷ���ϴ��ļ��������û�ͷ�������û�ID��������<BR>
�û�ID�Ļ�ÿ���ͨ���û���Ϣ��������������û���Ȼ������Ƶ��û��������ϣ��鿴�������ԣ�����UserID = ��������û���ID
</td>
</tr>
</table>
<table width = "95%" border = "0" cellspacing = "1" cellpadding = "3"  align = center class = "tableBorder" style = "table-layout:fixed;word-break:break-all">
<tr align = center><th width = "*" height = 25>�ļ���</th><th width = "100">�����û�</th><th width = "50">��С</th><th width = "120">������</th><th width = "120">�ϴ�����</th><th width = "35">����</th></tr>
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
		If instr(upname.name,"_") Then    'ȡ��ͷ����û���
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
		response.write "<td align = center class = forumRowHighlight><a href = '?filename="&upname.name&"'>ɾ��</a></td></tr>"
	ElseIf i>page*pagesize Then
		Exit For
	End If
	usernames = ""
Next

End Sub 

'����ͷ��
Sub delface()
Dim DllUserFace
dnum = 0
DllUserFace = Request("filename")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
'ɾ������ͷ��������ļ�;
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
	'ȡ��ͷ����û���
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
					Response.Write "ͷ���Ѹ���,�û�"& usernames &"��ͷ���ļ���"& upfilename &"��ɾ��<br>"
					dnum = dnum+1
				End If
			Else
				objFSO.DeleteFile(Server.MapPath(upfilename))
				response.write "�û�"& uid(1) &"��ע��,�ļ���"& upfilename &"��ɾ��<br>"
				dnum = dnum+1
			End If
			rs.close
			set rs = nothing
		End If
	Else
	'����û���û�ID��ͷ���ļ�
		sql = "select top 1 userid from [dv_user] where userface = '"& upfilename &"' "
		Set rs = Dvbbs.Execute(sql)
		If rs.eof Then
			objFSO.DeleteFile(Server.MapPath(upfilename))
			response.write "�����ɾ���ļ���"& upfilename &"<br>"
			dnum = dnum+1
		End If
		rs.close
		Set rs = nothing
	End If
ElseIf i>page*pagesize Then
	Exit For
End If
Next
response.write " ������ "& dnum &" ���ļ�  "
End Sub

Sub foot()
Set uploadFolder = Nothing
Set uploadFiles = Nothing
%>
<tr><td colspan=6 class=forumRow height=30>
<%
If page>1 Then
	response.write "<a href=?page=1>��ҳ</a>&nbsp;&nbsp;<a href=""?page="& page-1 &""">��һҳ</a>&nbsp;&nbsp;"
Else
	response.write "��ҳ&nbsp;&nbsp;��һҳ&nbsp;&nbsp;"
End If
If page<i/pagesize Then
	response.write "<a href=""?page="& page+1 &""">��һҳ</a>&nbsp;&nbsp;<a href=""?page="& pagenum &""">βҳ</a>"
Else
	response.write "��һҳ&nbsp;&nbsp;βҳ"
End If
%>
<input type="submit" value="����"></td><tr></form></table><br>
<% End Sub %>
