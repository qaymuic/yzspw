<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
	dim Str
	dim admin_flag
	admin_flag=",11,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		if Request("action") = "unite" then
		call unite()
		else
		call boardinfo()
		end if
	Footer()
	end if

sub boardinfo()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
	<th height=25>�ϲ���̳����
	</th>
	</tr>
	<form action=admin_boardunite.asp?action=unite method=post>
	<tr>
	<td class=forumrow>
	<B>�ϲ���̳ѡ��</B>��<BR>
<B>������̳����������������Ӷ�ת����Ŀ����̳����ɾ������̳������������</B><BR><BR>
<%
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select boardid,boardtype,depth from dv_board order by rootid,orders"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "û����̳"
	else
		response.write " ����̳ "
		response.write "<select name=oldboard size=1>"
		do while not rs.eof
%>
<option value="<%=rs("boardid")%>"><%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
��
<%next%>
<%end if%><%=rs("boardtype")%></option>
<%
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	sql="select boardid,boardtype,depth from dv_board order by rootid,orders"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "û����̳"
	else
		response.write " �ϲ��� "
		response.write "<select name=newboard size=1>"
		do while not rs.eof
%>
<option value="<%=rs("boardid")%>"><%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
��
<%next%>
<%end if%><%=rs("boardtype")%></option>
<%
		rs.movenext
		loop
		response.write "</select>"
	end if
	rs.close
	set rs=nothing
	response.write " <BR><BR><input type=submit name=Submit value=�ϲ���̳><BR><BR>"
%>
	</td>
	</tr>
	<tr>
	<td class=forumrow><B>ע������</B>��<BR><FONT COLOR="red">���в��������棬�����ز���</FONT><BR> ������ͬһ�������ڽ��в��������ܽ�һ������ϲ�����������̳�С�<BR>�ϲ�������ָ������̳�����߰�����������̳������ɾ�����������ӽ�ת�Ƶ�����ָ����Ŀ����̳��
	</td>
	</tr></form>
	</table>
<%
end sub

sub unite()
dim newboard
dim oldboard
dim ParentStr,iParentStr
dim depth,iParentID,child
Dim ParentID,RootID
if clng(request("newboard"))=clng(request("oldboard")) then
	response.write "�벻Ҫ����ͬ�����ڽ��в�����"
	exit sub
end if
newboard=clng(request("newboard"))
oldboard=clng(request("oldboard"))
'������̳����������������Ӷ�ת����Ŀ����̳����ɾ������̳������������
'�õ���ǰ����������̳
set rs=Dvbbs.Execute("select ParentStr,Boardid,depth,ParentID,child,RootID from dv_board where boardid="&oldboard)
if rs(0)="0" then
	ParentStr=rs(1)
	iParentID=rs(1)
	ParentID=0
else
	ParentStr=rs(0) & "," & Rs(1)
	iParentID=rs(3)
	ParentID=rs(3)
end if
iParentStr=rs(1)
depth=rs(2)
child=rs(4)+1
RootID=rs(5)
i=0
If ParentID=0 Then
set rs=Dvbbs.Execute("select Boardid from dv_board where boardid="&newboard&" and RootID="&RootID)
Else
set rs=Dvbbs.Execute("select Boardid from dv_board where boardid="&newboard&" and ParentStr like '%"&ParentStr&"%'")
End If
if not (rs.eof and rs.bof) then
	response.write "���ܽ�һ������ϲ�����������̳��"
	exit sub
end if
'�õ���ǰ����������̳ID
i=0
set rs=Dvbbs.Execute("select Boardid from dv_board where RootID="&RootID&" And ParentStr like '%"&ParentStr&"%'")
if not (rs.eof and rs.bof) then
do while not rs.eof
	if i=0 then
		iParentStr=rs(0)
	else
		iParentStr=iParentStr & "," & rs(0)
	end if
	i=i+1
rs.movenext
loop
end if
if i>0 then
	ParentStr=iParentStr & "," & oldboard
else
	ParentStr=oldboard
end if
'������ԭ��������̳������
if depth>0 then
Dvbbs.Execute("update dv_board set child=child-"&child&" where boardid="&iparentid)
'������ԭ��������̳���ݣ������൱�ڼ�֦�����迼��
for i=1 to depth
	'�õ��丸��ĸ���İ���ID
	set rs=Dvbbs.Execute("select parentid from dv_board where boardid="&iparentid)
	if not (rs.eof and rs.bof) then
		iparentid=rs(0)
		Dvbbs.Execute("update dv_board set child=child-"&child&" where boardid="&iparentid)
	end if
next
end if
'������̳��������
For i=0 to ubound(AllPostTable)
	Dvbbs.Execute("update "&AllPostTable(i)&" set boardid="&newboard&" where boardid in ("&ParentStr&")")
	'���»���վ��������
	Dvbbs.Execute("Update "&AllPostTable(i)&" Set LockTopic="&newboard&" Where BoardID=444 And LockTopic In ("&ParentStr&")")
Next
Dvbbs.Execute("update dv_topic set boardid="&newboard&" where boardid in ("&ParentStr&")")
Dvbbs.Execute("update dv_besttopic set boardid="&newboard&" where boardid in ("&ParentStr&")")
'���»���վ��������
Dvbbs.Execute("Update Dv_Topic Set LockTopic="&newboard&" Where BoardID=444 And LockTopic In ("&ParentStr&")")
'shinzeal��������ϴ��ļ�����
Dvbbs.Execute("update DV_Upfile set F_boardid="&newboard&" where F_boardid in ("&ParentStr&")")
'ɾ�����ϲ���̳
set rs=Dvbbs.Execute("select sum(postnum),sum(topicnum),sum(todayNum) from dv_board where RootID="&RootID&" And boardid in ("&ParentStr&")")
Dvbbs.Execute("delete from dv_board where RootID="&RootID&" And boardid in ("&ParentStr&")")
'��������̳���Ӽ���
dim trs
set trs=Dvbbs.Execute("select ParentStr,boardid from dv_board where boardid="&newboard)
if trs(0)="0" then
ParentStr=trs(1)
else
ParentStr=trs(0)
end if
Dvbbs.Execute("update dv_board set postnum=postnum+"&rs(0)&",topicnum=topicnum+"&rs(1)&",todaynum=todaynum+"&rs(2)&" where boardid in ("&ParentStr&")")
response.write "�ϲ��ɹ����Ѿ������ϲ���̳��������ת�������ϲ���̳��"
set rs=nothing
set trs=nothing
Dvbbs.ReloadAllBoardInfo()
Dvbbs.Name="setup"
Dvbbs.ReloadSetup
Dvbbs.CacheData=Dvbbs.value
Dim Forum_Boards
Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
For i=0 To Ubound(Forum_Boards)
	Dvbbs.ReloadBoardInfo(Forum_Boards(i))
Next
end sub

%>