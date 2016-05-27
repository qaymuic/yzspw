<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
	Server.ScriptTimeOut=9999999
	dim admin_flag
	admin_flag="24,25"
	if not Dvbbs.master or instr(","&session("flag")&",",",24,")=0 or instr(","&session("flag")&",",",25,")=0 then
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
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
<td height="23" colspan="2" class=Forumrow><B>说明</B>：<BR>您可以选择下列其中之一种模式进行帖子数据在不同表之间的转换。</td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>模式一：搜索要转移的帖子</th>
</tr>
<FORM METHOD=POST ACTION="?action=search">
<tr> 
<td height="23" width="20%" class=Forumrow><B>搜索条件</B></td>
<td height="23" width="80%" class=Forumrow><input type="text" name="keyword">&nbsp;
<select name="tablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
&nbsp;<input type="checkbox" name="username" value="yes" checked>用户&nbsp;<input type="checkbox" name="topic" value="yes">主题&nbsp;<input type="submit" name="submit" value="搜索"></td>
</tr>
</FORM>
<tr> 
<td height="23" colspan="2" class=Forumrow><B>注意</B>：这里仅搜索所在表的主题和发表用户数据，搜索后对搜索结果进行操作</td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>模式二：在不同表转移数据</th>
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
 <input type="checkbox" name="top1000" value="yes" checked>前 <input type="checkbox" name="end1000" value="yes">后 <input type=text name="selnum" value="100" size=3>条 记录转移到
<select name="InTablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
&nbsp;<input type="submit" name="submit" value="提交">
</td>
</tr>
</FORM>
<tr> 
<td height="23" colspan="2" class=Forumrow><B>注意</B>：最前N条记录指数据库中最早发表的帖子（如果平均每个帖子有5个回复，那么100个主题在这里的更新量将是500条记录），这通常要花很长的时间，更新的速度取决于您的服务器性能以及更新数据的多少。执行本步骤将消耗大量的服务器资源，建议您在访问人数较少的时候或者本地进行更新操作。</td>
</tr>
</table>
<%
end sub

sub nowused()
%>
<form method="POST" action="?action=update">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="5" class=Forumrow><B>说明</B>：<BR>下列数据表中选中的为当前论坛所使用来保存帖子数据的表，一般情况下每个表中的数据越少论坛帖子显示速度越快，当您下列单个帖子数据表中的数据有超过几万的帖子时不妨新添一个数据表来保存帖子数据（SQL版本用户建议每个表数据达到20万以后进行添加表操作），您会发现论坛速度快很多很多。<BR>您也可以将当前所使用的数据表在下列数据表中切换，当前所使用的帖子数据表即当前论坛用户发贴时默认的保存帖子数据表</td>
</tr>
<tr> 
<th height="23" colspan="5">当前数据表设定</th>
</tr>
<tr> 
<td width="20%" class=forumHeaderBackgroundAlternate><b>表名<B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>说明</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>当前帖数</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>当前默认</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>删除</B></td>
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
<td width="20%" class=Forumrow><a href="?action=del&tablename=<%=AllPostTable(i)%>"  onclick="{if(confirm('删除将包括该数据表所有帖子，本操作所删除的内容将不可恢复，确定删除吗?')){return true;}return false;}">删除</a></td>
</tr>
<%
next
%>
<tr> 
<td width="100%" colspan=5 class=Forumrow> 
<input type="submit" name="Submit" value="提 交">
</td>
</tr>
</form>
<FORM METHOD=POST ACTION="?action=CreatTable">
<tr> 
<th height="23" colspan="5">添加数据表</th>
</tr>
<tr> 
<td width="20%" class=Forumrow>添加的表名</td>
<td width="80%" class=Forumrow colspan=4><input type=text name="tablename" value="Dv_bbs<%=ubound(AllPostTable)+2%>">&nbsp;只能用Dv_bbs+数字表示，如Dv_bbs5，最后的数字最多不能超过9</td>
</tr>
<tr> 
<td width="20%" class=Forumrow>添加表的说明</td>
<td width="80%" class=Forumrow colspan=4><input type=text name="tablereadme">&nbsp;简单描述该表的用途，在搜索帖子和其他相关操作部分显示</td>
</tr>
<tr> 
<td width="100%" colspan=5 class=Forumrow> 
<input type="submit" name="Submit" value="提 交">
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
	response.write "更新成功！"
end sub

sub del()
	dim nAllPostTable,nAllPostTableName,ii
	if Trim(request("tablename"))=Trim(Dvbbs.NowUseBBS) then
		Errmsg=ErrMsg + "<BR><li>当前正在使用的表不能删除。"
		dvbbs_error()
		exit sub
	end if
	
	Dvbbs.Execute("delete from dv_Tablelist where TableName='"&Trim(request("TableName"))&"'")
	Dvbbs.Execute("drop table "&request("tablename")&"")
	Dvbbs.Execute("delete from dv_BestTopic where RootID in (select TopicID from dv_topic where PostTable='"&request("tablename")&"')")
	Dvbbs.Execute("delete from dv_Topic where PostTable='"&request("tablename")&"'")
	response.write "删除成功！"
end sub

sub CreatTable()
if request.form("tablename")="" then
	Errmsg=ErrMsg + "<BR><li>请输入表名。"
	dvbbs_error()
	exit sub
elseif len(request.form("tablename"))<>7 then
	Errmsg=ErrMsg + "<BR><li>输入的表名不合法。"
	dvbbs_error()
	exit sub
elseif not isnumeric(right(request.form("tablename"),1)) then
	Errmsg=ErrMsg + "<BR><li>输入的表名不合法。"
	dvbbs_error()
	exit sub
elseif cint(right(request.form("tablename"),1))>9 or cint(right(request.form("tablename"),1))<0 then
	Errmsg=ErrMsg + "<BR><li>输入的表名不合法。"
	dvbbs_error()
	exit sub
end if
if request.form("tablereadme")="" then
	Errmsg=ErrMsg + "<BR><li>请输入表的说明。"
	dvbbs_error()
	exit sub
end if
for i=0 to ubound(AllPostTable)
	if AllPostTable(i)=request.form("tablename") then
		Errmsg=ErrMsg + "<BR><li>您输入的表名已经存在，请重新输入。"
		dvbbs_error()
		exit sub
	end if
next

Dim NewAllPostTable,NewAllPostTableName
'更新数据表列表

Dvbbs.Execute("insert into dv_TableList(TableName,TableType)Values('"&request.form("tablename")&"','"&request.form("tablereadme")&"') ")
'NewAllPostTable=rs(0) & "|" & request.form("tablename")
'NewAllPostTableName=rs(1) & "|" & request.form("tablereadme")

'Set conn = Server.CreateObject("ADODB.connection")
'connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("dvbbs5.mdb")
'conn.open connstr
'建立新表
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

'添加索引
Dvbbs.Execute("create index dispbbs on "&request.form("tablename")&" (boardid,rootid)")
Dvbbs.Execute("create index save_1 on "&request.form("tablename")&" (rootid,orders)")
Dvbbs.Execute("create index disp on "&request.form("tablename")&" (boardid)")

'Dvbbs.Execute("update config set AllPostTable='"&NewAllPostTable&"',AllPostTableName='"&NewAllPostTableName&"'")
response.write "添加表成功，请返回。"
end sub

'模式2更新
sub update2()
dim trs
dim ForNum,TopNum
Dim orderby,PostUserID
if request.form("outtablename")=request.form("intablename") then
	Errmsg=ErrMsg + "<BR><li>不能在相同数据表内转移数据。"
	dvbbs_error()
	exit sub
end if
if (not isnumeric(request.form("selnum"))) or request.form("selnum")="" then
	Errmsg=ErrMsg + "<BR><li>请填写正确的更新数量。"
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
下面开始转移论坛帖子资料，预计本次共有<%=C1%>个帖子需要更新
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
	Errmsg=ErrMsg + "<BR><li>您所选择导出的数据表已经没有任何内容"
	dvbbs_error()
	exit sub
else
	do while not rs.eof
		'读取导出帖子数据表
		set trs=Dvbbs.Execute("select * from "&request.form("outtablename")&" where rootid="&rs("topicid")&" order by Announceid")
		if not (trs.eof and trs.bof) then
		do while not trs.eof
		'插入导入帖子数据表
		If IsNull(trs("postuserid")) Or trs("postuserid")="" Then
			PostUserID=0
		Else
			PostUserID=trs("postuserid")
		End If
		Dvbbs.Execute("insert into "&request("intablename")&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isagree,isupload,isaudit,PostBuyUser,UbbList) values ("&trs("boardid")&","&trs("parentid")&",'"&Dvbbs.checkstr(trs("username"))&"','"&Dvbbs.checkstr(trs("topic"))&"','"&Dvbbs.checkstr(trs("body"))&"','"&trs("dateandtime")&"',"&trs("length")&","&trs("rootid")&","&trs("layer")&","&trs("orders")&",'"&trs("ip")&"','"&trs("Expression")&"',"&trs("locktopic")&","&trs("signflag")&","&trs("emailflag")&","&trs("isbest")&","&PostUserID&",'"&trs("isagree")&"',"&trs("isupload")&","&trs("isaudit")&",'"&trs("PostBuyUser")&"','"&Dvbbs.checkstr(trs("UbbList"))&"')")
		'更新精华
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
		'删除导出帖子数据表对应数据
		Dvbbs.Execute("delete from "&request.form("outTableName")&" where RootID="&rs("TopicID"))
		'更新主题指定帖子表
		Dvbbs.Execute("update dv_topic set PostTable='"&request.form("inTableName")&"' where TopicID="&rs("topicid"))
		i=i+1
		'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""更新完"&Server.HtmlEncode(rs(1))&"的数据，正在更新下一个帖子数据，" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
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
dv_suc("转移数据更新成功！")
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
	Errmsg=ErrMsg + "<BR><li>请输入您要查询的关键字。"
	dvbbs_error()
	exit sub
else
	keyword=replace(request("keyword"),"'","")
end if
if request("username")="yes" then
Set Rs=Dvbbs.Execute("Select UserID From Dv_User Where UserName='"&keyword&"'")
If Rs.Eof And Rs.Bof Then
	Errmsg=ErrMsg + "<BR><li>目标用户并不存在，请重新输入。"
	dvbbs_error()
	exit sub
Else
	PostUserID=Rs(0)
End If
sql="select * from dv_topic where PostTable='"&request("tablename")&"' and PostUserID="&PostUserID&" order by LastPostTime desc"
elseif request("topic")="yes" then
sql="select * from dv_topic where PostTable='"&request("tablename")&"' and title like '%"&keyword&"%' order by LastPostTime desc"
else
	Errmsg=ErrMsg + "<BR><li>请选择您查询的方式。"
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
<td height="23" colspan="6" class=Forumrow><B>说明</B>：<BR>您可以对下列的搜索结果进行转移数据表的操作，不能在相同表内进行转换操作。</td>
</tr>
<tr> 
<th height="23" colspan="6">搜索<%=request("tablename")%>结果</th>
</tr>
<tr> 
<td width="6%" class=forumHeaderBackgroundAlternate align=center><b>状态<B></td>
<td width="45%" class=forumHeaderBackgroundAlternate align=center><B>标题</B></td>
<td width="15%" class=forumHeaderBackgroundAlternate align=center><B>作者</B></td>
<td width="6%" class=forumHeaderBackgroundAlternate align=center><B>回复</B></td>
<td width="22%" class=forumHeaderBackgroundAlternate align=center><B>时间</B></td>
<td width="6%" class=forumHeaderBackgroundAlternate align=center><B>操作</B></td>
</tr>
<%
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	response.write "<tr> <td class=Forumrow colspan=6 height=25>没有搜索到相关内容。</td></tr>"
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
	response.write "锁定"
elseif rs("isvote")=1 then
	response.write "投票"
elseif rs("isbest")=1 then
	response.write "精华"
else
	response.write "正常"
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
	response.write "<tr><td valign=middle nowrap colspan=2 class=forumrow height=25>&nbsp;&nbsp;分页： "

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
	response.write "<td colspan=3 class=forumrow>所有查询结果<input type=checkbox name=allsearch value=yes>"
	response.write "&nbsp;<select name=toTablename>"

	for i=0 to ubound(AllPostTable)
		response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
	next

	response.write "</select>&nbsp;<input type=submit name=submit value=转换>"
	response.write "</td>"
	response.write "<td class=forumrow align=center><input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">"
	response.write "</td></tr>"
end if
rs.close
set rs=nothing
response.write "</table></form><BR><BR>"
end sub

'根据搜索结果更新
sub update3()
dim keyword,trs,PostUserID
if request.form("tablename")=request.form("totablename") then
	Errmsg=ErrMsg + "<BR><li>不能在相同数据表内进行数据转换。"
	dvbbs_error()
	exit sub
end if
if request.form("allsearch")="yes" then
	if request("keyword")="" then
		Errmsg=ErrMsg + "<BR><li>请输入您要查询的关键字。"
		dvbbs_error()
		exit sub
	else
		keyword=replace(request("keyword"),"'","")
	end if
	if request("username")="yes" then
		Set Rs=Dvbbs.Execute("Select UserID From Dv_User Where UserName='"&keyword&"'")
		If Rs.Eof And Rs.Bof Then
			Errmsg=ErrMsg + "<BR><li>目标用户并不存在，请重新输入。"
			dvbbs_error()
			exit sub
		Else
			PostUserID=Rs(0)
		End If
		sql="select topicid,title from dv_topic where PostTable='"&request("tablename")&"' and PostUserID="&PostUserID&" order by LastPostTime desc"
	elseif request("topic")="yes" then
		sql="select topicid,title from dv_topic where PostTable='"&request("tablename")&"' and title like '%"&keyword&"%' order by LastPostTime desc"
	else
		Errmsg=ErrMsg + "<BR><li>请选择您查询的方式。"
		dvbbs_error()
		exit sub
	end if
else
	if request.form("topicid")="" then
		Errmsg=ErrMsg + "<BR><li>请选择要转移的帖子。"
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
下面开始转移论坛帖子资料，预计本次共有<%=C1%>个帖子需要更新
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
	Errmsg=ErrMsg + "<BR><li>没有任何记录可转换。"
	dvbbs_error()
	exit sub
else
	do while not rs.eof
	'取出原表数据
	set trs=Dvbbs.Execute("select * from "&request("tablename")&" where rootid="&rs("topicid")&" order by Announceid")
	if not (trs.eof and trs.bof) then
	'插入新表
	do while not trs.eof
		If IsNull(trs("postuserid")) Or trs("postuserid")="" Then
			PostUserID=0
		Else
			PostUserID=trs("postuserid")
		End If
	Dvbbs.Execute("insert into "&request("totablename")&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isagree,isupload,isaudit,PostBuyUser,UbbList) values ("&trs("boardid")&","&trs("parentid")&",'"&Dvbbs.checkstr(trs("username"))&"','"&Dvbbs.checkstr(trs("topic"))&"','"&Dvbbs.checkstr(trs("body"))&"','"&trs("dateandtime")&"',"&trs("length")&","&trs("rootid")&","&trs("layer")&","&trs("orders")&",'"&trs("ip")&"','"&trs("Expression")&"',"&trs("locktopic")&","&trs("signflag")&","&trs("emailflag")&","&trs("isbest")&","&PostUserID&",'"&trs("isagree")&"',"&trs("isupload")&","&trs("isaudit")&",'"&trs("PostBuyUser")&"','"&Dvbbs.checkstr(trs("UbbList"))&"')")
	'更新精华
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
	'删除原表该帖子数据
	Dvbbs.Execute("delete from "&request("tablename")&" where rootid="&rs("topicid"))
	'更新该主题表名
	Dvbbs.Execute("update dv_topic set PostTable='"&request("totablename")&"' where topicid="&rs("topicid"))
		i=i+1
		'If (i mod 100) = 0 Then
		Response.Write "<script>img2.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
		Response.Write "txt2.innerHTML=""更新完"&Server.HtmlEncode(rs(1))&"的数据，正在更新下一个帖子数据，" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
		Response.Write "img2.title=""" & Server.HtmlEncode(Rs(1)) & "(" & i & ")"";</script>" & VbCrLf
		Response.Flush
		'End If
	rs.movenext
	loop
end if
set trs=nothing
set rs=nothing
Response.Write "<script>img2.width=400;txt2.innerHTML=""100"";</script>"
dv_suc("转移数据更新成功！")
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
