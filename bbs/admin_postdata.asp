<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
	Server.ScriptTimeOut=9999999
	dim admin_flag
	admin_flag="24,25"
	if not Dvbbs.master or instr(","&session("flag")&",",",24,")=0 or instr(","&session("flag")&",",",25,")=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		if request("action")="Nowused" then
		call nowused()
		elseif request("action")="update" then
		call update()
		elseif request("action")="del" then
		call del()
		elseif request("action")="CreatTable" then
		call creattable()
		elseif request("action")="search" then
		call search()
		elseif request("action")="update2" then
		call update2()
		elseif request("action")="update3" then
		call update3()
		else
		call main()
		end if
		Footer()
	end if

sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="2" class=Forumrow><B>˵��</B>��<BR>������ѡ����������֮һ��ģʽ�������������ڲ�ͬ��֮���ת����</td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>ģʽһ������Ҫת�Ƶ�����</th>
</tr>
<FORM METHOD=POST ACTION="?action=search">
<tr> 
<td height="23" width="20%" class=Forumrow><B>��������</B></td>
<td height="23" width="80%" class=Forumrow><input type="text" name="keyword">&nbsp;
<select name="tablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
&nbsp;<input type="checkbox" name="username" value="yes" checked>�û�&nbsp;<input type="checkbox" name="topic" value="yes">����&nbsp;<input type="submit" name="submit" value="����"></td>
</tr>
</FORM>
<tr> 
<td height="23" colspan="2" class=Forumrow><B>ע��</B>��������������ڱ������ͷ����û����ݣ������������������в���</td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>ģʽ�����ڲ�ͬ��ת������</th>
</tr>
<FORM METHOD=POST ACTION="?action=update2">
<tr> 
<td height="23" width="100%" class=Forumrow colspan="2">&nbsp;
<select name="OutTablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
 <input type="checkbox" name="top1000" value="yes" checked>ǰ <input type="checkbox" name="end1000" value="yes">�� <input type=text name="selnum" value="100" size=3>�� ��¼ת�Ƶ�
<select name="InTablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
&nbsp;<input type="submit" name="submit" value="�ύ">
</td>
</tr>
</FORM>
<tr> 
<td height="23" colspan="2" class=Forumrow><B>ע��</B>����ǰN����¼ָ���ݿ������緢������ӣ����ƽ��ÿ��������5���ظ�����ô100������������ĸ���������500����¼������ͨ��Ҫ���ܳ���ʱ�䣬���µ��ٶ�ȡ�������ķ����������Լ��������ݵĶ��١�ִ�б����轫���Ĵ����ķ�������Դ���������ڷ����������ٵ�ʱ����߱��ؽ��и��²�����</td>
</tr>
</table>
<%
end sub

sub nowused()
%>
<form method="POST" action="?action=update">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="5" class=Forumrow><B>˵��</B>��<BR>�������ݱ���ѡ�е�Ϊ��ǰ��̳��ʹ���������������ݵı�һ�������ÿ�����е�����Խ����̳������ʾ�ٶ�Խ�죬�������е����������ݱ��е������г������������ʱ��������һ�����ݱ��������������ݣ�SQL�汾�û�����ÿ�������ݴﵽ20���Ժ������ӱ�����������ᷢ����̳�ٶȿ�ܶ�ܶࡣ<BR>��Ҳ���Խ���ǰ��ʹ�õ����ݱ����������ݱ����л�����ǰ��ʹ�õ��������ݱ���ǰ��̳�û�����ʱĬ�ϵı����������ݱ�</td>
</tr>
<tr> 
<th height="23" colspan="5">��ǰ���ݱ��趨</th>
</tr>
<tr> 
<td width="20%" class=forumHeaderBackgroundAlternate><b>����<B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>˵��</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>��ǰ����</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>��ǰĬ��</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>ɾ��</B></td>
</tr>
<%
for i=0 to ubound(AllPostTable)
%>
<tr> 
<td width="20%" class=Forumrow><%=AllPostTable(i)%></td>
<td width="20%" class=Forumrow><%=AllPostTableName(i)%></td>
<td width="20%" class=Forumrow>
<%
set rs=Dvbbs.Execute("select count(*) from "&AllPostTable(i)&"")
response.write rs(0)
%>
</td>
<td width="20%" class=Forumrow><input value="<%=AllPostTable(i)%>" name="TableName" type="radio" <%if Trim(Lcase(Dvbbs.NowUseBBS))=Lcase(AllPostTable(i)) then%>checked<%end if%>></td>
<td width="20%" class=Forumrow><a href="?action=del&tablename=<%=AllPostTable(i)%>"  onclick="{if(confirm('ɾ�������������ݱ��������ӣ���������ɾ�������ݽ����ɻָ���ȷ��ɾ����?')){return true;}return false;}">ɾ��</a></td>
</tr>
<%
next
%>
<tr> 
<td width="100%" colspan=5 class=Forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</form>
<FORM METHOD=POST ACTION="?action=CreatTable">
<tr> 
<th height="23" colspan="5">������ݱ�</th>
</tr>
<tr> 
<td width="20%" class=Forumrow>��ӵı���</td>
<td width="80%" class=Forumrow colspan=4><input type=text name="tablename" value="Dv_bbs<%=ubound(AllPostTable)+2%>">&nbsp;ֻ����Dv_bbs+���ֱ�ʾ����Dv_bbs5������������಻�ܳ���9</td>
</tr>
<tr> 
<td width="20%" class=Forumrow>��ӱ��˵��</td>
<td width="80%" class=Forumrow colspan=4><input type=text name="tablereadme">&nbsp;�������ñ����;�����������Ӻ�������ز���������ʾ</td>
</tr>
<tr> 
<td width="100%" colspan=5 class=Forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</FORM>
</table>
<%
end sub

sub update()
	Dvbbs.Execute("update Dv_setup set Forum_NowUseBBS='"&request.form("TableName")&"'")
	Dvbbs.Name="setup"
	Dvbbs.ReloadSetup
	response.write "���³ɹ���"
end sub

sub del()
	dim nAllPostTable,nAllPostTableName,ii
	if Trim(request("tablename"))=Trim(Dvbbs.NowUseBBS) then
		Errmsg=ErrMsg + "<BR><li>��ǰ����ʹ�õı���ɾ����"
		dvbbs_error()
		exit sub
	end if
	
	Dvbbs.Execute("delete from dv_Tablelist where TableName='"&Trim(request("TableName"))&"'")
	Dvbbs.Execute("drop table "&request("tablename")&"")
	Dvbbs.Execute("delete from dv_BestTopic where RootID in (select TopicID from dv_topic where PostTable='"&request("tablename")&"')")
	Dvbbs.Execute("delete from dv_Topic where PostTable='"&request("tablename")&"'")
	response.write "ɾ���ɹ���"
end sub

sub CreatTable()
if request.form("tablename")="" then
	Errmsg=ErrMsg + "<BR><li>�����������"
	dvbbs_error()
	exit sub
elseif len(request.form("tablename"))<>7 then
	Errmsg=ErrMsg + "<BR><li>����ı������Ϸ���"
	dvbbs_error()
	exit sub
elseif not isnumeric(right(request.form("tablename"),1)) then
	Errmsg=ErrMsg + "<BR><li>����ı������Ϸ���"
	dvbbs_error()
	exit sub
elseif cint(right(request.form("tablename"),1))>9 or cint(right(request.form("tablename"),1))<0 then
	Errmsg=ErrMsg + "<BR><li>����ı������Ϸ���"
	dvbbs_error()
	exit sub
end if
if request.form("tablereadme")="" then
	Errmsg=ErrMsg + "<BR><li>��������˵����"
	dvbbs_error()
	exit sub
end if
for i=0 to ubound(AllPostTable)
	if AllPostTable(i)=request.form("tablename") then
		Errmsg=ErrMsg + "<BR><li>������ı����Ѿ����ڣ����������롣"
		dvbbs_error()
		exit sub
	end if
next

Dim NewAllPostTable,NewAllPostTableName
'�������ݱ��б�

Dvbbs.Execute("insert into dv_TableList(TableName,TableType)Values('"&request.form("tablename")&"','"&request.form("tablereadme")&"') ")
'NewAllPostTable=rs(0) & "|" & request.form("tablename")
'NewAllPostTableName=rs(1) & "|" & request.form("tablereadme")

'Set conn = Server.CreateObject("ADODB.connection")
'connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("dvbbs5.mdb")
'conn.open connstr
'�����±�
If IsSqlDataBase=1 Then
sql="CREATE TABLE [dbo].["&request.form("tablename")&"] (AnnounceID int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_"&request.form("tablename")&" PRIMARY KEY,"&_
		"ParentID int default 0,"&_
		"BoardID int default 0,"&_
		"UserName varchar(50),"&_
		"PostUserID int default 0,"&_
		"Topic varchar(250),"&_
		"Body text,"&_
		"DateAndTime datetime default "&SqlNowString&","&_
		"length int Default 0,"&_
		"RootID int Default 0,"&_
		"layer int Default 0,"&_
		"orders int Default 0,"&_
		"isbest tinyint Default 0,"&_
		"ip varchar(40) NULL,"&_
		"Expression varchar(20) NULL,"&_
		"locktopic int Default 0,"&_
		"signflag tinyint Default 0,"&_
		"emailflag tinyint Default 0,"&_
		"isagree varchar(50) NULL,"&_
		"isupload tinyint default 0,"&_
		"isaudit tinyint default 0,"&_
		"PostBuyUser text,"&_
		"UbbList varchar(255))"
Else
sql="CREATE TABLE "&request.form("tablename")&" (AnnounceID int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,"&_
		"ParentID int default 0,"&_
		"BoardID int default 0,"&_
		"UserName varchar(50),"&_
		"PostUserID int default 0,"&_
		"Topic varchar(250),"&_
		"Body text,"&_
		"DateAndTime datetime default Now(),"&_
		"length int Default 0,"&_
		"RootID int Default 0,"&_
		"layer int Default 0,"&_
		"orders int Default 0,"&_
		"isbest tinyint Default 0,"&_
		"ip varchar(40) NULL,"&_
		"Expression varchar(20) NULL,"&_
		"locktopic int Default 0,"&_
		"signflag tinyint Default 0,"&_
		"emailflag tinyint Default 0,"&_
		"isagree varchar(50) NULL,"&_
		"isupload tinyint default 0,"&_
		"isaudit tinyint default 0,"&_
		"PostBuyUser text,"&_
		"UbbList varchar(255))"
End If 
Dvbbs.Execute(sql)

'�������
Dvbbs.Execute("create index dispbbs on "&request.form("tablename")&" (boardid,rootid)")
Dvbbs.Execute("create index save_1 on "&request.form("tablename")&" (rootid,orders)")
Dvbbs.Execute("create index disp on "&request.form("tablename")&" (boardid)")

'Dvbbs.Execute("update config set AllPostTable='"&NewAllPostTable&"',AllPostTableName='"&NewAllPostTableName&"'")
response.write "��ӱ�ɹ����뷵�ء�"
end sub

'ģʽ2����
sub update2()
dim trs
dim ForNum,TopNum
Dim orderby,PostUserID
if request.form("outtablename")=request.form("intablename") then
	Errmsg=ErrMsg + "<BR><li>��������ͬ���ݱ���ת�����ݡ�"
	dvbbs_error()
	exit sub
end if
if (not isnumeric(request.form("selnum"))) or request.form("selnum")="" then
	Errmsg=ErrMsg + "<BR><li>����д��ȷ�ĸ���������"
	dvbbs_error()
	exit sub
end if
if request.form("top1000")="yes" then
orderby=""
else
orderby=" desc"
end if
TopNum=Clng(request.form("selnum"))
if TopNum>100 then
	ForNum=int(TopNum/100)+1
	TopNum=100
else
	ForNum=1
end if

Dim C1
C1=TopNum
%>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<tr><td colspan=2 class=forumrow>
���濪ʼת����̳�������ϣ�Ԥ�Ʊ��ι���<%=C1%>��������Ҫ����
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=ffffff height=9><img src="skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>
<%
Response.Flush

dim myrs,maxannid
for i=1 to ForNum
set rs=Dvbbs.Execute("select top "&TopNum&" topicid,title from dv_topic where PostTable='"&request.form("outtablename")&"' order by topicid "&orderby&"")
if rs.eof and rs.bof then
	Errmsg=ErrMsg + "<BR><li>����ѡ�񵼳������ݱ��Ѿ�û���κ�����"
	dvbbs_error()
	exit sub
else
	do while not rs.eof
		'��ȡ�����������ݱ�
		set trs=Dvbbs.Execute("select * from "&request.form("outtablename")&" where rootid="&rs("topicid")&" order by Announceid")
		if not (trs.eof and trs.bof) then
		do while not trs.eof
		'���뵼���������ݱ�
		If IsNull(trs("postuserid")) Or trs("postuserid")="" Then
			PostUserID=0
		Else
			PostUserID=trs("postuserid")
		End If
		Dvbbs.Execute("insert into "&request("intablename")&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isagree,isupload,isaudit,PostBuyUser,UbbList) values ("&trs("boardid")&","&trs("parentid")&",'"&Dvbbs.checkstr(trs("username"))&"','"&Dvbbs.checkstr(trs("topic"))&"','"&Dvbbs.checkstr(trs("body"))&"','"&trs("dateandtime")&"',"&trs("length")&","&trs("rootid")&","&trs("layer")&","&trs("orders")&",'"&trs("ip")&"','"&trs("Expression")&"',"&trs("locktopic")&","&trs("signflag")&","&trs("emailflag")&","&trs("isbest")&","&PostUserID&",'"&trs("isagree")&"',"&trs("isupload")&","&trs("isaudit")&",'"&trs("PostBuyUser")&"','"&Dvbbs.checkstr(trs("UbbList"))&"')")
		'���¾���
		if trs("isbest")=1 then
			set myrs=Dvbbs.Execute("select max(announceid) from "&request.form("intablename")&" where boardid="&trs("boardid"))
			maxannid=myrs(0)
			myrs.close
			set myrs=nothing
			Dvbbs.Execute("update dv_besttopic set AnnounceID="&maxannid&" where rootid="&rs("topicid"))
		end if
		trs.movenext
		loop
		end if
		'ɾ�������������ݱ��Ӧ����
		Dvbbs.Execute("delete from "&request.form("outTableName")&" where RootID="&rs("TopicID"))
		'��������ָ�����ӱ�
		Dvbbs.Execute("update dv_topic set PostTable='"&request.form("inTableName")&"' where TopicID="&rs("topicid"))
		i=i+1
		'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""������"&Server.HtmlEncode(rs(1))&"�����ݣ����ڸ�����һ���������ݣ�" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Server.HtmlEncode(Rs(1)) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
		'End If
	rs.movenext
	loop
end if
next
set trs=nothing
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
dv_suc("ת�����ݸ��³ɹ���")
end sub

sub search()
dim keyword
dim totalrec
dim n
dim currentpage,page_count,Pcount,PostUserID
currentPage=request("page")
if currentpage="" or not IsNumeric(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
if request("keyword")="" then
	Errmsg=ErrMsg + "<BR><li>��������Ҫ��ѯ�Ĺؼ��֡�"
	dvbbs_error()
	exit sub
else
	keyword=replace(request("keyword"),"'","")
end if
if request("username")="yes" then
Set Rs=Dvbbs.Execute("Select UserID From Dv_User Where UserName='"&keyword&"'")
If Rs.Eof And Rs.Bof Then
	Errmsg=ErrMsg + "<BR><li>Ŀ���û��������ڣ����������롣"
	dvbbs_error()
	exit sub
Else
	PostUserID=Rs(0)
End If
sql="select * from dv_topic where PostTable='"&request("tablename")&"' and PostUserID="&PostUserID&" order by LastPostTime desc"
elseif request("topic")="yes" then
sql="select * from dv_topic where PostTable='"&request("tablename")&"' and title like '%"&keyword&"%' order by LastPostTime desc"
else
	Errmsg=ErrMsg + "<BR><li>��ѡ������ѯ�ķ�ʽ��"
	dvbbs_error()
	exit sub
end if
%>
<form method="POST" action="?action=update3">
<input type=hidden name="topic" value="<%=request("topic")%>">
<input type=hidden name="username" value="<%=request("username")%>">
<input type=hidden name="keyword" value="<%=keyword%>">
<input type=hidden name="tablename" value="<%=request("tablename")%>">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="6" class=Forumrow><B>˵��</B>��<BR>�����Զ����е������������ת�����ݱ�Ĳ�������������ͬ���ڽ���ת��������</td>
</tr>
<tr> 
<th height="23" colspan="6">����<%=request("tablename")%>���</th>
</tr>
<tr> 
<td width="6%" class=forumHeaderBackgroundAlternate align=center><b>״̬<B></td>
<td width="45%" class=forumHeaderBackgroundAlternate align=center><B>����</B></td>
<td width="15%" class=forumHeaderBackgroundAlternate align=center><B>����</B></td>
<td width="6%" class=forumHeaderBackgroundAlternate align=center><B>�ظ�</B></td>
<td width="22%" class=forumHeaderBackgroundAlternate align=center><B>ʱ��</B></td>
<td width="6%" class=forumHeaderBackgroundAlternate align=center><B>����</B></td>
</tr>
<%
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	response.write "<tr> <td class=Forumrow colspan=6 height=25>û��������������ݡ�</td></tr>"
else
	rs.PageSize = Dvbbs.Forum_Setting(11)
	rs.AbsolutePage=currentpage
	page_count=0
	totalrec=rs.recordcount
	while (not rs.eof) and (not page_count = rs.PageSize)
%>
<tr> 
<td width="6%" class=Forumrow align=center>
<%
if rs("locktopic")=1 then
	response.write "����"
elseif rs("isvote")=1 then
	response.write "ͶƱ"
elseif rs("isbest")=1 then
	response.write "����"
else
	response.write "����"
end if
%>
</td>
<td width="45%" class=Forumrow><%=dvbbs.htmlencode(rs("title"))%></td>
<td width="15%" class=Forumrow align=center><a href="admin_user.asp?action=modify&userid=<%=rs("postuserid")%>"><%=dvbbs.htmlencode(rs("postusername"))%></a></td>
<td width="6%" class=Forumrow align=center><%=rs("child")%></td>
<td width="22%" class=Forumrow><%=rs("dateandtime")%></td>
<td width="6%" class=Forumrow align=center><input type="checkbox" name="topicid" value="<%=rs("topicid")%>"></td>
</tr>
<%
  	page_count = page_count + 1
	rs.movenext
	wend
	dim endpage
	Pcount=rs.PageCount
	response.write "<tr><td valign=middle nowrap colspan=2 class=forumrow height=25>&nbsp;&nbsp;��ҳ�� "

	if currentpage > 4 then
	response.write "<a href=""?page=1&action=search&keyword="&keyword&"&topic="&request("topic")&"&username="&request("username")&"&tablename="&request("tablename")&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color="&Dvbbs.mainsetting(1)&">["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&"&action=search&keyword="&keyword&"&topic="&request("topic")&"&username="&request("username")&"&tablename="&request("tablename")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&action=search&keyword="&keyword&"&topic="&request("topic")&"&username="&request("username")&"&tablename="&request("tablename")&""">["&Pcount&"]</a>"
	end if
	response.write "</td>"
	response.write "<td colspan=3 class=forumrow>���в�ѯ���<input type=checkbox name=allsearch value=yes>"
	response.write "&nbsp;<select name=toTablename>"

	for i=0 to ubound(AllPostTable)
		response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
	next

	response.write "</select>&nbsp;<input type=submit name=submit value=ת��>"
	response.write "</td>"
	response.write "<td class=forumrow align=center><input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">"
	response.write "</td></tr>"
end if
rs.close
set rs=nothing
response.write "</table></form><BR><BR>"
end sub

'���������������
sub update3()
dim keyword,trs,PostUserID
if request.form("tablename")=request.form("totablename") then
	Errmsg=ErrMsg + "<BR><li>��������ͬ���ݱ��ڽ�������ת����"
	dvbbs_error()
	exit sub
end if
if request.form("allsearch")="yes" then
	if request("keyword")="" then
		Errmsg=ErrMsg + "<BR><li>��������Ҫ��ѯ�Ĺؼ��֡�"
		dvbbs_error()
		exit sub
	else
		keyword=replace(request("keyword"),"'","")
	end if
	if request("username")="yes" then
		Set Rs=Dvbbs.Execute("Select UserID From Dv_User Where UserName='"&keyword&"'")
		If Rs.Eof And Rs.Bof Then
			Errmsg=ErrMsg + "<BR><li>Ŀ���û��������ڣ����������롣"
			dvbbs_error()
			exit sub
		Else
			PostUserID=Rs(0)
		End If
		sql="select topicid,title from dv_topic where PostTable='"&request("tablename")&"' and PostUserID="&PostUserID&" order by LastPostTime desc"
	elseif request("topic")="yes" then
		sql="select topicid,title from dv_topic where PostTable='"&request("tablename")&"' and title like '%"&keyword&"%' order by LastPostTime desc"
	else
		Errmsg=ErrMsg + "<BR><li>��ѡ������ѯ�ķ�ʽ��"
		dvbbs_error()
		exit sub
	end if
else
	if request.form("topicid")="" then
		Errmsg=ErrMsg + "<BR><li>��ѡ��Ҫת�Ƶ����ӡ�"
		dvbbs_error()
		exit sub
	end if
	sql="select topicid,title from dv_topic where PostTable='"&request("tablename")&"' and TopicID in ("&request.form("TopicID")&")"
end if

'set rs=Dvbbs.Execute(sql)
Set Rs=server.createobject("adodb.recordset")
Rs.Open SQL,Conn,1,1
Dim C1,myrs,maxannid
C1=Rs.ReCordCount
%>
&nbsp;<BR>
<table cellpadding="0" cellspacing="0" border="0" width="95%" class="tableBorder" align=center>
<tr><td colspan=2 class=forumrow>
���濪ʼת����̳�������ϣ�Ԥ�Ʊ��ι���<%=C1%>��������Ҫ����
<table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr> 
<td bgcolor=ffffff height=9><img src="skins/default/bar/bar3.gif" width=0 height=16 id=img2 name=img2 align=absmiddle></td></tr></table>
</td></tr></table> <span id=txt2 name=txt2 style="font-size:9pt">0</span><span style="font-size:9pt">%</span></td></tr>
</table>
<%
Response.Flush
if rs.eof and rs.bof then
	Errmsg=ErrMsg + "<BR><li>û���κμ�¼��ת����"
	dvbbs_error()
	exit sub
else
	do while not rs.eof
	'ȡ��ԭ������
	set trs=Dvbbs.Execute("select * from "&request("tablename")&" where rootid="&rs("topicid")&" order by Announceid")
	if not (trs.eof and trs.bof) then
	'�����±�
	do while not trs.eof
		If IsNull(trs("postuserid")) Or trs("postuserid")="" Then
			PostUserID=0
		Else
			PostUserID=trs("postuserid")
		End If
	Dvbbs.Execute("insert into "&request("totablename")&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isagree,isupload,isaudit,PostBuyUser,UbbList) values ("&trs("boardid")&","&trs("parentid")&",'"&Dvbbs.checkstr(trs("username"))&"','"&Dvbbs.checkstr(trs("topic"))&"','"&Dvbbs.checkstr(trs("body"))&"','"&trs("dateandtime")&"',"&trs("length")&","&trs("rootid")&","&trs("layer")&","&trs("orders")&",'"&trs("ip")&"','"&trs("Expression")&"',"&trs("locktopic")&","&trs("signflag")&","&trs("emailflag")&","&trs("isbest")&","&PostUserID&",'"&trs("isagree")&"',"&trs("isupload")&","&trs("isaudit")&",'"&trs("PostBuyUser")&"','"&Dvbbs.checkstr(trs("UbbList"))&"')")
	'���¾���
	if Not IsNull(trs("isbest")) And trs("isbest")<>"" then
		If trs("isbest")=1 Then
		set myrs=Dvbbs.Execute("select max(announceid) from "&request.form("totablename")&" where boardid="&trs("boardid"))
		maxannid=myrs(0)
		myrs.close
		set myrs=nothing
		Dvbbs.Execute("update dv_besttopic set AnnounceID="&maxannid&" where rootid="&rs("topicid"))
		End If
	end if
	trs.movenext
	loop
	end if
	'ɾ��ԭ�����������
	Dvbbs.Execute("delete from "&request("tablename")&" where rootid="&rs("topicid"))
	'���¸��������
	Dvbbs.Execute("update dv_topic set PostTable='"&request("totablename")&"' where topicid="&rs("topicid"))
		i=i+1
		'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""������"&Server.HtmlEncode(rs(1))&"�����ݣ����ڸ�����һ���������ݣ�" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Server.HtmlEncode(Rs(1)) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
		'End If
	rs.movenext
	loop
end if
set trs=nothing
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
dv_suc("ת�����ݸ��³ɹ���")
end sub
%>
<script language="javascript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
