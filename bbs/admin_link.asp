<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
	dim admin_flag
	admin_flag=",13,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		dim body
		dim readme,Tlink
		call main()
		set rs=nothing
		Footer()
	end if

Sub main()
Select Case request("action")
	Case "saveall"
		Call saveall()
	Case "add" 
		Call addlink()
	Case "edit"
		Call editlink()
	Case "savenew"
		Call savenew()
	Case "savedit"
		Call savedit()
	Case "del"
		Call del()
	Case "orders"
		Call orders()
	Case "updatorders"
		Call updateorders()
	Case Else
		call linkinfo()
End Select 
Response.Write body
End Sub

Sub addlink()
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3""  align=center class=""tableBorder"">"
	Response.Write "<form action=""admin_link.asp?action=savenew"" method = post> <tr>"
	Response.Write "<th width=""100%"" colspan=2 height=25>���������̳ </th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=forumrow>��̳���� </td>"
	Response.Write "<td width=""60%"" height=25 class=forumrow>"
	Response.Write "<input type=""text"" name=""name"" size=40>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=forumrow>����URL </td>"
	Response.Write "<td width=""60%"" class=forumrow>"
	Response.Write "<input type=""text"" name=""url"" size=40>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=forumrow>����LOGO��ַ </td>"
	Response.Write "<td width=""60%"" class=forumrow>"
	Response.Write "<input type=""text"" name=""logo"" size=40>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=forumrow>��̳��� </td>"
	Response.Write "<td width=""60%"" class=forumrow>"
	Response.Write "<input type=""text"" name=""readme"" size=40>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=forumrow>����ҳ���������ӻ���LOGO���� </td>"
	Response.Write "<td width=""60%"" class=forumrow>"
	Response.Write "��������<input type=""radio"" name=""islogo"" value=0 checked>&nbsp;&nbsp;LOGO����<input type=""radio"" name=""islogo"" value=1>"
	Response.Write "&nbsp;&nbsp;<Input type=""submit"" name=""Submit"" value=""�� ��"">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</form>"
	Response.Write "</table>"
End Sub

sub editlink()
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from dv_bbslink where id="&Request("id")
	rs.open sql,conn,1,1
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write "<form action=""admin_link.asp?action=savedit"" method=post>"
	Response.Write "<input type=hidden name=id value="
	Response.Write Request("id")
	Response.Write "><tr> <th width=""100%"" colspan=2 height=25>�༭������̳</th>"
	Response.Write "</tr><tr> "
	Response.Write "<td width=""40%"" class=forumrow>"
	Response.Write "��̳���ƣ� </td>"
	Response.Write "<td width=""60%"" class=forumrow>"
	Response.Write "<input type=""text"" name=""name"" size=40 value="
	Response.Write server.htmlencode(rs("boardname"))
	Response.Write ">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" class=forumrow>"
	Response.Write "����URL�� </td>"
	Response.Write "<td width=""60%"" class=forumrow>"
	Response.Write "<input type=""text"" name=""url"" size=40 value="
	Response.Write server.htmlencode(rs("url"))
	Response.Write ">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" class=forumrow>"
	Response.Write "����LOGO��ַ�� </td>"
	Response.Write "<td width=""60%"" class=forumrow>"
	Response.Write "<input type=""text"" name=""logo"" size=40 value="""
	If Rs("logo")<>"" or Not IsNull(Rs("logo")) Then Response.Write server.htmlencode(rs("logo"))
	Response.Write """>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" class=forumrow>"
	Response.Write "��̳��飺 </td>"
	Response.Write "<td width=""60%"" class=forumrow>"
	Response.Write "<input type=""text"" name=""readme"" size=40 value="
	If Rs("logo")<>"" or Not IsNull(Rs("logo")) Then Response.Write server.htmlencode(rs("readme"))
	Response.Write ">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td width=""40%"" height=25 class=forumrow>����ҳ���������ӻ���LOGO���� </td>"
	Response.Write "<td width=""60%"" class=forumrow>"
	Response.Write "��������<input type=""radio"" name=""islogo"" value=0 "
	If rs("islogo")=0 Then
	 	Response.Write " checked"
	End If
	Response.Write ">&nbsp;&nbsp;LOGO����<input type=""radio"" name=""islogo"" value=1 "
	If rs("islogo")=1 Then
		Response.Write " checked"
	End If 
	Response.Write ">"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td height=""15"" colspan=""2"" class=forumrow>"
	Response.Write "<div align=""center"">"
	Response.Write "<input type=""submit"" name=""Submit"" value=""�� ��"">"
	Response.Write "</div>"
	Response.Write "</td>"
	Response.Write "</tr></form>"
	Response.Write "</table>"
	Rs.Close
	Set Rs=Nothing
End Sub

Sub linkinfo()
	Dim i 
	i=0
	addlink()
	Set rs= server.createobject ("adodb.recordset")
	sql = " select * from dv_bbslink order by id"
	rs.open sql,conn,1,1       
	Response.Write "<br><table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3""  align=center class=""tableBorder"">"
	Response.Write "<form action=""admin_link.asp?action=saveall"" method = post>"
	Response.Write "<tr>"
	Response.Write "<th height=""22"" colspan=4>������̳�б������޸� </th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td height=""22"" colspan=4 class=forumrowhighlight><b>ע�����</b><li>����Ա༭��������������ϢȻ��һ�����ύ��<li>������Ų��ֿ��԰��������Ҫ�޸ģ��������ظ���<li>�������ݾ�������ʹ�õ����ţ������ƻ�ҳ�����ִ�С�</td>"
	Response.Write "</tr>"
	If rs.eof and rs.bof Then
		Response.Write "<tr><td height=""25"" colspan=4 align=""center"" class=forumrowhighlight>��δ���������̳</td></tr>"
	Else
		
	Do While Not Rs.EOF
	Response.Write "<tr align=left>"
    Response.Write "<td height=25 class=forumrow>"
    Response.Write "<B>��ţ�</B> <input type=""text"" name=""id"" size=4 value="
	Response.Write Rs("id")
	Response.Write "></td>"
	Response.Write "<td class=forumrow>"
	Response.Write "<B>���ƣ�</B><input type=""text"" name=""boardname"&i&""" size=30 value="
	Response.Write server.htmlencode(Rs("boardname")&"")
	Response.Write "></td>"
	Response.Write "<td class=forumrow>"
	Response.Write "<B> URL��</B><input type=""text"" name=""url"&i&""" size=35 value="
	Response.Write server.htmlencode(Rs("url")&"")
	Response.Write "></td>"
	Response.Write "<td class=forumrow ><a href=""admin_link.asp?action=orders&id="
	Response.Write Rs("id")
	Response.Write """>����</a>  <a href=""admin_link.asp?action=edit&id="
	Response.Write Rs("id")
	Response.Write """>�༭</a>  <a href=""admin_link.asp?action=del&id="
	Response.Write Rs("id")
	Response.Write """>ɾ��</a></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td class=forumrow ><b>�Ƿ�ͼƬ</b><br>"
	Response.Write "��<input type=""radio"" name=""islogo"&i&""" value=""1"" "
	If Rs("islogo")=1 Then	
		Response.Write " checked"
	End If
	Response.Write ">"
	
	Response.Write "��<input type=""radio"" name=""islogo"&i&""" value=""0"" "
	If Rs("islogo")=0 Then	
		Response.Write " checked"
	End If
	Response.Write ">"
	Response.Write "</td><td class=forumrow >"
	Response.Write "<b>logo��</b>"
	Response.Write "<input type=""text"" name=""logo"&i&""" size=30 value="
	If Rs("logo")<>"" or Not IsNull(Rs("logo")) Then Response.Write server.htmlencode(Rs("logo"))
	Response.Write "></td>"
	Response.Write "<td class=forumrow colspan=4><B>��飺</B>"
	Response.Write "<input type=""text"" name=""readme"&i&""" size=50 value="
	If Rs("logo")<>"" or Not IsNull(Rs("logo")) Then Response.Write server.htmlencode(Rs("readme"))
	Response.Write "></td></tr>"
	Response.Write "<tr><th height=""1"" colspan=4></td></tr>"
	i=i+1
	rs.movenext
	loop
	
	Response.Write "<tr><td height=""25"" colspan=4 align=""center""><input type=""submit"" name=""Submit"" value=""��������""></td></tr>"
	Response.Write "</from>"
	End If
	Response.Write "</table>"
        rs.Close
	set rs=Nothing
End Sub 

sub savenew()
if Request("url")<>"" and Request("readme")<>"" and request("name")<>"" then
	dim linknum
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from dv_bbslink order by id desc"
	rs.Open sql,conn,1,3
	if rs.eof and rs.bof then
	linknum=1
	else
	linknum=rs("id")+1
	end if
	sql="insert into dv_bbslink(id,boardname,readme,logo,url,islogo) values("&linknum&",'"&fixjs(Trim(Request.Form ("name")))&"','"&fixjs(Trim(Request.Form ("readme")))&"','"&fixjs(trim(request.Form("logo")))&"','"&fixjs(Request.Form ("url"))&"',"&CInt(request.Form("islogo"))&")"
	Dvbbs.Execute(sql) 
	rs.Close
	set rs=Nothing 
	Call cache_link()
	body=body+"<br>"+"���³ɹ������������������"
else
	body=body+"<br>"+"����������������̳��Ϣ��"
end if
end sub

sub savedit()
	set rs= server.createobject ("adodb.recordset")
	sql = "select * from dv_bbslink where id="&request("id")
	rs.Open sql,conn,1,3
	if rs.eof and rs.bof then
	body=body+"<br>"+"����û���ҵ�������̳��"
	else
	rs("boardname") = fixjs(Trim(Request.Form ("name")))
	rs("readme") =  fixjs(Trim(Request.Form ("readme")))
	rs("logo")=fixjs(Trim(request.Form("logo")))
	rs("url") = fixjs(Request.Form ("url"))
	rs("islogo")=request.Form("islogo")
	rs.Update
	end if 
	rs.Close
	set rs=nothing
	Call cache_link()
	body=body+"<br>"+"���³ɹ������������������"
end sub

sub del
	dim id
	id = request("id")
	sql="delete from dv_bbslink where id="+id
	Dvbbs.Execute(sql)
	body=body+"<br>"+"ɾ���ɹ������������������"
	Call cache_link()
end sub

sub orders()
	Response.Write "<br>"
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=center class=tableborder>"
	Response.Write "			<tr><th height=24>������̳��������</th></tr>"
	Response.Write "<tr>"
	Response.Write "<td height=""23"" class=forumrowhighlight>"
	Response.Write "ע�⣺������Ӧ��̳���������������Ӧ��������ţ�<font color=red>ע�ⲻ�ܺͱ��������̳����ͬ���������</font>��</font>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<td class=forumrow>"

	set rs= server.createobject ("adodb.recordset")
	sql="select * from dv_bbslink where id="&cstr(request("id"))
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "û���ҵ���Ӧ��������̳��"
	else
		response.write "<form action=admin_link.asp?action=updatorders method=post>"
		response.write ""&rs("boardname")&"  <input type=text name=newid size=2 value="&rs("id")&">"
		response.write "<input type=hidden name=id value="&request("id")&">"
		response.write "<input type=submit name=Submit value=�޸�></form>"
	end if
	rs.close
	set rs=Nothing
	Response.Write"</td>"
	Response.write"</tr>"
	Response.write"<tr>"
	Response.write"<td>"
	Response.write"<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=Left class=tableborder>"
	Response.write"<tr><th height=24 colspan=4>��ǰ�������</th></tr>"
	Response.Write "<tr>"
	Dim a
	a=0		
	Set rs=Dvbbs.Execute("select id,boardname from dv_bbslink order by id")
	Do While Not Rs.EOF
	Response.Write "<td height=22 width=""25%"">"
	Response.Write Rs(0)
	Response.Write "��"
	Response.Write Rs(1)
	Response.Write "</td>"
	
	Rs.MoveNext
	a=a+1
	If a=4 Then
		a=0
		Response.Write "</tr><tr><th height=""1"" colspan=4></td></tr><tr>"		
	End If
	Loop
	Set rs=Nothing
	Response.Write "</tr>"
	Response.Write "<tr><th height=""1"" colspan=4></td></tr>"
	Response.Write " </table>"
	Response.Write"</td></tr>"
	Response.write"</table>"			
end sub

sub updateorders()
if isnumeric(request("id")) and isnumeric(request("newid")) and request("newid")<>request("id") then
	set rs=Dvbbs.Execute("select id from dv_bbslink where id="&request("newid"))
	if rs.eof and rs.bof then
	sql="update dv_bbslink set id="&request("newid")&" where id="&cstr(request("id"))
	Dvbbs.Execute(sql)
	response.write "���³ɹ���"
	else
	response.write "����ʧ�ܣ���ָ���˺�����������̳��ͬ����ţ�"
	end if
else
	response.write "����ʧ�ܣ���������ַ����Ϸ������������˺�ԭ����ͬ����ţ�"
end if
end sub

sub cache_link()
	Dvbbs.Name="link"
	Dim Rs,SQl
	SQL="select boardname,readme,url,logo,islogo from [Dv_bbslink] Order by islogo,id"
	Set Rs=Dvbbs.Execute(SQL)
	If Not rs.eof Then
		Dvbbs.Value=RS.GetString (,,"!@#%|","$?&!@","")
	Else
		Dvbbs.Value=""
	End If
	Set Rs=Nothing
end sub
Sub saveall()
	Dim IDlist,id,i,tmpstr
	ID=Request.form("id")
	id=Replace(id," ","")
	IDlist=","&ID&","
	ID=split(id,",")
	For i=0 to UBound(id)
		tmpstr=","&ID(i)&","
		If InStr(IDlist,tmpstr)>0 Then
			If InStr(Len(tmpstr)-1+InStr(IDlist,tmpstr),IDlist,tmpstr)>0 Then
			Errmsg=ErrMsg + "������ͬ����ţ�"&ID(i)&",�뷵����ϸ��顣"
			Exit For
			End If 
		End If	
	Next 	
	If Errmsg<>"" Then
		dvbbs_error()
	End If
	'���ԭ�����ݱ��е�����,�������������ˡ�����
	Dvbbs.Execute("Delete from dv_bbslink")
	'��ʼ����ѭ����������
	Dim sql,boardname,readme,url,logo,islogo
	For i= 0 to UBound(id)
		boardname=fixjs(Request.form("boardname"&i))
		readme=fixjs(Request.form("readme"&i))
		url=fixjs(Request.form("url"&i))
		logo=fixjs(Request.form("logo"&i))
		islogo=Request.form("islogo"&i)
		sql="insert into dv_bbslink (id,boardname,readme,url,logo,islogo) values ("&CInt(id(i))&",'"&boardname&"','"&readme&"','"&url&"','"&logo&"',"&CInt(islogo)&")"
		Dvbbs.Execute(sql)	
	Next
	cache_link()
	Dv_suc("��̳�������³ɹ���")
End Sub 
%>
