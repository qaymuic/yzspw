<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
Head()
Dim path
Dim objFSO
Dim uploadfolder
Dim uploadfiles
Dim upname
Dim UpFolder
Dim upfilename
Dim admin_flag
admin_flag=",35,"
Dim sfor(30,2)
Dim seachstr,sqlstr,delsql
Dim currentpage,page_count,Pcount
Dim totalrec,endpage

if Request("path")<>"" then
path=Request("path")
else 
path="UploadFile"
end if
If Dvbbs.Forum_Setting(76)="0" Or  Dvbbs.Forum_Setting(76)="" Then Dvbbs.Forum_Setting(76)="UploadFile/"
path=Dvbbs.Forum_Setting(76)
currentPage=Request("currentpage")
if currentpage="" or not IsNumeric(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
	if err then
		currentpage=1
		err.clear
	end if
end if

if Request("filesearch")<>"" and IsNumeric(Request("filesearch")) then
seachstr="&filesearch="&Request("filesearch")
end if

'----------------------------------
'��������ѯ����������ʼ
'----------------------------------
if Request("filesearch")=7 and IsNumeric(Request("filesearch")) then

	'�����������
	if Request("class")<>"" and IsNumeric(Request("class")) and Request("class")<>0 then
	seachstr=seachstr+"&class="&cint(Request("class"))
	sqlstr=" and F_BoardID="&cint(Request("class"))
	end if

	'������������
	if Request("f_type")<>"" and IsNumeric(Request("f_type")) then
	seachstr=seachstr+"&f_type="&cint(Request("f_type"))
	sqlstr=sqlstr+" and f_type="&cint(Request("f_type"))
	end if

	'������������
	if Request("f_filetype")<>"" then
	seachstr=seachstr+"&f_filetype="&Request("f_filetype")
	sqlstr=sqlstr+" and f_filetype='"&dvbbs.checkstr(Request("f_filetype"))&"'"
	end if

	'���ش�������f_downnum
	if Request("f_downnum")<>"" and IsNumeric(Request("f_downnum")) then
		if Request("downtype")="more" then
		sqlstr=sqlstr+" and f_downnum>="&clng(Request("f_downnum"))
		else
		sqlstr=sqlstr+" and f_downnum<="&clng(Request("f_downnum"))
		end if
		seachstr=seachstr+"&f_downnum="&cint(Request("f_downnum"))&"&downtype="&Request("downtype")
	end if

	'�����������f_viewnum
	if Request("f_viewnum")<>"" and IsNumeric(Request("f_viewnum")) then
		if Request("viewtype")="more" then
		sqlstr=sqlstr+" and f_viewnum>="&clng(Request("f_viewnum"))
		else
		sqlstr=sqlstr+" and f_viewnum<="&clng(Request("f_viewnum"))
		end if
		seachstr=seachstr+"&f_viewnum="&cint(Request("f_viewnum"))&"&viewtype="&Request("viewtype")
	end if

	'������С����f_size
	if Request("f_size")<>"" and IsNumeric(Request("f_size")) then
		if Request("sizetype")="more" then
		sqlstr=sqlstr+" and F_FileSize>="&clng(Request("f_size"))*1024
		else
		sqlstr=sqlstr+" and F_FileSize<="&clng(Request("f_size"))*1024
		end if
		seachstr=seachstr+"&f_size="&cint(Request("f_size"))&"&sizetype="&Request("sizetype")
	end if

	'�������ڷ�������f_adddatenum
	if Request("f_adddatenum")<>"" and IsNumeric(Request("f_adddatenum")) then
		If IsSqlDataBase=1 Then
			if Request("timetype")="more" then
			sqlstr=sqlstr+" and datediff(day,F_AddTime,"&SqlNowString&") >= "&clng(Request("f_adddatenum"))
			else
			sqlstr=sqlstr+" and datediff(day,F_AddTime,"&SqlNowString&") <= "&clng(Request("f_adddatenum"))
			end if
		Else
			if Request("timetype")="more" then
			sqlstr=sqlstr+" and datediff('d',F_AddTime,"&SqlNowString&") >= "&clng(Request("f_adddatenum"))
			else
			sqlstr=sqlstr+" and datediff('d',F_AddTime,"&SqlNowString&") <= "&clng(Request("f_adddatenum"))
			end if
		End If
		seachstr=seachstr+"&f_adddatenum="&cint(Request("f_adddatenum"))&"&timetype="&Request("timetype")
	end if

	'�������ߣ�
	if Request("f_username")<>"" then
		if Request("usernamechk")="yes" then
		sqlstr=sqlstr+" and f_username='"&dvbbs.checkstr(Request("f_username"))&"'"
		else
		sqlstr=sqlstr+" and f_username like '%"&dvbbs.checkstr(Request("f_username"))&"%'"
		end if
		seachstr=seachstr+"&f_username="&Request("f_username")&"&usernamechk="&Request("usernamechk")
	end if
	'����˵����
	if Request("f_readme")<>"" then
		if Request("f_readmechk")="yes" then
		sqlstr=sqlstr+" and f_readme='"&dvbbs.checkstr(Request("f_readme"))&"'"
		else
		sqlstr=sqlstr+" and f_readme like '%"&dvbbs.checkstr(Request("f_readme"))&"%'"
		end if
		seachstr=seachstr+"&f_readme="&Request("f_readme")&"&f_readmechk="&Request("f_readmechk")
	end if
end if
'----------------------------------
'��������ѯ������������
'----------------------------------

if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й�����ҳ���Ȩ�ޡ�"
	dvbbs_error()
else
%>
  <table border="0" cellpadding="3" cellspacing="1" width="95%" class=tableborder align=center>
    <tr>
      <th height="23" colspan="2">��̳�ϴ���������</th>
    </tr>
    <tr>
      <td width="20%" height="23" class="forumRowHighlight">ע�����</td>
      <td width="80%" class=forumRow>
	 �١������ܱ��������֧��FSOȨ�޷���ʹ�ã�FSOʹ�ð��������΢����վ���������������֧��FSO���ֶ�������	<BR>�ڡ��°棨�ģ֣���֮��İ汾�ϴ�Ŀ¼ǿ�ƶ���ΪUploadFile��ֻ�и�Ŀ¼���ļ��ɽ����ļ��Զ������������°�֮ǰ�İ汾�ϴ��ļ�ֻ���ֶ���������ϴ��ļ������ģ֣���������������ϴ��������Զ���ŵ����Զ�����ļ����У��ļ�Ŀ¼�Ե���������������Ҫ�ռ�֧�֣ƣӣ϶�дȨ�ޣ�
	 <br>�ۡ��Զ������ļ������������ϴ��ļ����к�ʵ���緢���ļ�û�б����������ʹ�ã���ִ���Զ��������
	  </td>
    </tr>
	<tr>
	<form action="?action=FileSearch" method=post>
      <td width="20%" height="23" class="forumRowHighlight">���ٲ�ѯ��</td>
      <td width="80%" class=forumRow>
	  <select size=1 name="FileSearch" onchange="javascript:submit()">
	<option value="0">��ѡ���ѯ����</option>
	<option value="1" <%if Request("FileSearch")=1 then%>selected<%end if%>>�г������ϴ�����</option>
	<option value="2" <%if Request("FileSearch")=2 then%>selected<%end if%>>���	����Сʱ���ϴ��ĸ���</option>
	<option value="3" <%if Request("FileSearch")=3 then%>selected<%end if%>>������������ϴ��ĸ���</option>
	<option value="4" <%if Request("FileSearch")=4 then%>selected<%end if%>>������������ϴ��ĸ���</option>
	<option value="5" <%if Request("FileSearch")=5 then%>selected<%end if%>>����ǰ���������ĸ���</option>
	<option value="6" <%if Request("FileSearch")=6 then%>selected<%end if%>>���ǰ���������ĸ���</option>
	</select>
	  </td>
	 </FORM>
    </tr>
  </table>
<%
	if Request("Submit")="���������ϴ���¼" then
		call delall()
	elseif Request("Submit")="���δ��¼�ļ�" then
		call delall1()
	elseif Request("Submit")="������ǰ�б���¼" then
		call delall()
	elseif Request("action")="FileSearch" then
		call FileSearch()
	elseif Request("action")="delfiles" then
		call delfiles()
	else
		call main()
	end if
	Footer()
end if

sub main()
%>
<br><table border="0" cellpadding="3" cellspacing="1" width="95%" class=tableborder align=center>
<form action="?action=FileSearch" method=post>
<tr>
	<th height="23" colspan="2" align=left>�߼���ѯ</th>
</tr>
<tr>
<td width=20% class="forumRowHighlight">ע������</td>
<td width=80% class=forumrow colspan=5>�ڼ�¼�ܶ���������������Խ���ѯԽ�����뾡�����ٲ�ѯ������</td>
</tr>
<tr>
	<td width="20%" height="23" class="forumRowHighlight">������飺</td>
	<td width="80%" class=forumRow>
	<select name=class>
	<option value="0">������̳���</option>
<%
Dim rs_c
set rs_c= server.CreateObject ("adodb.recordset")
sql = "select * from dv_board order by rootid,orders"
rs_c.open sql,conn,1,1
do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <%if Request("editid")<>"" and clng(Request("editid"))=rs_c("boardid") then%>selected<%end if%>>
<%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>
��
<%next%>
<%end if%><%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close
set rs_c=nothing
%>
	</select>
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="forumRowHighlight">�ļ����ش�����</td>
	<td width="80%" class=forumRow><input size=45 name="f_downnum" type=text>
	<input type=radio value=more name="downtype" checked >&nbsp;����&nbsp;
	<input type=radio value=less name="downtype" >&nbsp;����
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="forumRowHighlight">�������������</td>
	<td width="80%" class=forumRow><input size=45 name="f_viewnum" type=text>
	<input type=radio value=more name="viewtype" checked >&nbsp;����&nbsp;
	<input type=radio value=less name="viewtype" >&nbsp;����
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="forumRowHighlight">�ϴ�������</td>
	<td width="80%" class=forumRow><input size=45 name="f_adddatenum" type=text>
	<input type=radio value=more name="timetype" checked >&nbsp;����&nbsp;
	<input type=radio value=less name="timetype" >&nbsp;����
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="forumRowHighlight">�������ߣ�</td>
	<td width="80%" class=forumRow><input size=45 name="f_username" type=text>
	&nbsp;<input type=checkbox name="usernamechk" value="yes" checked>�û�������ƥ��
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="forumRowHighlight">����˵����</td>
	<td width="80%" class=forumRow><input size=45 name="f_readme" type=text>
	&nbsp;<input type=checkbox name="f_readmechk" value="yes" checked>˵����������ƥ��
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="forumRowHighlight">������С��</td>
	<td width="80%" class=forumRow><input size=45 name="f_size" type=text>&nbsp;(��λ��K)
	<input type=radio value=more name="sizetype" checked >&nbsp;����&nbsp;
	<input type=radio value=less name="sizetype" >&nbsp;С��
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="forumRowHighlight">�������ࣺ</td>
	<td width="80%" class=forumRow>
	<select name="f_type">
	<option value="all">���з���</option>
	<option value="1">ͼƬ������</option>
	<option value="2">FLASH������</option>
	<option value="3">���ּ�����</option>
	<option value="4">��Ӱ������</option>
	<option value="0">�ļ�������</option>
	</select>
	</td>
</tr>
<tr>
	<td width="20%" height="23" class="forumRowHighlight">�������ͣ�</td>
	<td width="80%" class=forumRow>
	<select name="f_filetype">
	<option value="">�����ļ�����</option>
	<option value="gif">gif</option><option value="jpg">jpg</option>
	<option value="bmp">bmp</option><option value="zip">zip</option>
	<option value="rar">rar</option><option value="exe">exe</option>
	<option value="swf">swf</option><option value="swi">swi</option>
	<option value="mid">mid</option><option value="mp3">mp3</option>
	<option value="rm">rm</option><option value="txt">txt</option>
	<option value="doc">doc</option><option value="exl">exl</option>
	</select>
	</td>
</tr>
<tr>
<th height="23" colspan="2"><input name="submit" type=submit value="��ʼ����"></th>
</tr>
<input type=hidden value="7" name="FileSearch">
</form>
</table>
<%
end sub

sub FileSearch()
%>
<form method=post action="?action=delfiles" name="formpost">
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr>
<th colspan=8 align=left height=23 ID=TableTitleLink><a href=admin_uploadlist.asp>�ϴ��ļ�����</a> -->�������</th>
</tr>
<tr>
<td class=forumRowHighlight align=center><B>����</B></td>
<td class=forumRowHighlight height=23 align=center><B>�û���</B></td>
<td class=forumRowHighlight align=center><B>�� �� ��</B></td>
<td class=forumRowHighlight align=center><B>�������</B></td>
<td class=forumRowHighlight align=center><B>��С</B></td>
<td class=forumRowHighlight align=center><B>ʱ��/���/����</B></td>
<td class=forumRowHighlight align=center><B>����</B></td>
<td class=forumRowHighlight align=center><B>ɾ��</B></td>
</tr>
<%
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select F_ID,F_AnnounceID,F_BoardID,F_Filename,F_Username,F_FileType,F_Type,F_FileSize,F_DownNum,F_ViewNum,F_AddTime ,B.Boardtype from [DV_Upfile] U inner join dv_Board B on B.boardid=U.F_BoardID where F_Flag=0 "
	'������ѯ
	select case Request("FileSearch")
	case 1
		sql=sql+" order by F_ID desc"
	case 2
		If IsSqlDataBase=1 Then
		sql=sql+" and datediff(hour,F_AddTime,"&SqlNowString&")<25"
		else
		sql=sql+" and datediff('h',F_AddTime,"&SqlNowString&")<25"
		end if
		sql=sql+" order by F_ID desc"
	case 3
		If IsSqlDataBase=1 Then
		sql=sql+" and datediff(month,F_AddTime,"&SqlNowString&")<1"
		else
		sql=sql+" and datediff('m',F_AddTime,"&SqlNowString&")<1"
		end if
		sql=sql+" order by F_ID desc"
	case 4
		If IsSqlDataBase=1 Then
		sql=sql+" and datediff(month,F_AddTime,"&SqlNowString&")<3"
		else
		sql=sql+" and datediff('m',F_AddTime,"&SqlNowString&")<3"
		end if
		sql=sql+" order by F_ID desc"
	case 5
		sql="select top 100 F_ID,F_AnnounceID,F_BoardID,F_Filename,F_Username,F_FileType,F_Type,F_FileSize,F_DownNum,F_ViewNum,F_AddTime ,B.Boardtype from [DV_Upfile] U inner join dv_Board B on B.boardid=U.F_BoardID where F_Flag=0 and F_BoardID<>0"
		sql=sql+" order by F_DownNum Desc,F_ID desc"
	case 6
		sql="select top 100 F_ID,F_AnnounceID,F_BoardID,F_Filename,F_Username,F_FileType,F_Type,F_FileSize,F_DownNum,F_ViewNum,F_AddTime ,B.Boardtype from [DV_Upfile] U inner join dv_Board B on B.boardid=U.F_BoardID where F_Flag=0 and F_BoardID<>0"
		sql=sql+" order by F_ViewNum Desc,F_ID desc"
	case 7
		sql=sql+sqlstr
		sql=sql+" order by F_ID desc"
	case else
		sql=sql+" order by F_ID desc"
	end select
	'response.write SQL
	rs.open sql,conn,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 class=forumrow>û���ҵ���ؼ�¼��</td></tr>"
	else
		rs.PageSize = Cint(Dvbbs.Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Dvbbs.Forum_Setting(11)))
		'�б�����'''''''''''''''''''''
%>
<tr>
<td class=forumRowHighlight align=center width=20>
	<img src="skins/default/filetype/<%=rs("F_FileType")%>.gif" border=0>
</td>
<td class=forumRow height=23 align=center><%=rs("F_Username")%></td>
<td class="forumRowHighlight">
	<%If Dvbbs.Forum_Setting(75)="1" Then%>
		<a href="<%=path%><%=rs("F_Filename")%>" target=_blank><%=rs("F_Filename")%></a>
	<%Else%>
		<a href="<%=path%>/<%=rs("F_Filename")%>" target=_blank><%=rs("F_Filename")%></a>
	<%End If%>
</td>
<td class=forumRow><%=rs("Boardtype")%></td>
<td class=forumRowHighlight><%=getsize(rs("F_FileSize"))%></td>
<td class=forumRow>
	<%=formatdatetime(rs("F_AddTime"),1)%>/
	<FONT COLOR=RED><%=rs("F_ViewNum")%></FONT>/
	<%=rs("F_DownNum")%>
</td>
<td class="forumRowHighlight" align=center><%=filetypename(rs("F_Type"))%></td>
<td class="forumRow" width=20><input type="checkbox" name="delid" value="<%=rs("F_ID")%>" ></td>
</tr>
<%		page_count = page_count + 1
		rs.movenext
		wend
		Pcount=rs.PageCount
	end if 
	rs.close
	if Request("FileSearch")=1 then sql=""
	if Request("FileSearch")=7 and sqlstr="" then sql=""
%>
<input type=hidden value="<%=sql%>" name="delsql">
<tr><th height=25 align=left colspan=8>�ļ���¼����������</th></tr>
<tr>
<td colspan=5 height=25 class="forumRowHighlight"><LI>��ѡȡҪɾ�����ļ���Ȼ��ִ��ɾ��������<font color=red>������ֱ�Ӵӷ�������ɾ�������ָܻ���</font></td>
<td colspan=3 height=25 class="forumRowHighlight"><input type="submit" name="Submit" value="ִ��ɾ����ѡ�ļ�"></td></tr>
<tr>
<td colspan=5 height=25 class="forumRow"><LI>����ͬʱ�Ƿ�ֱ�Ӵӷ�������ɾ���ļ���<font color=red>ɾ�����ļ������ָܻ� ��</font></td>
<td colspan=3 height=25 class="forumRow">
<input type=radio name=delfile value=1 >��&nbsp;
<input type=radio name=delfile value=2 checked>��
</td></tr>
<tr>
<td colspan=5 height=25 class="forumRowHighlight"><li>���ݵ�ǰ�б����ݽ����������������������������ɾ�ĵĸ�����</td>
<td colspan=3 height=25 class="forumRowHighlight">
<input type="submit" name="Submit" value="������ǰ�б���¼">
</td></tr>
<tr>
<td colspan=5 height=25 class="forumRow"><li>���ϴ���¼�У�������ط������������ݽ������������ɾ�ĵĸ�����</td>
<td colspan=3 height=25 class="forumRow">
<input type="submit" name="Submit" value="���������ϴ���¼">
</td></tr>
<tr><th height=25 align=left colspan=8>�ռ丽����������</th></tr>
<tr><td height=25 colspan=8 class="forumRowHighlight">
<li>������ڷ������ռ��û�м�¼���ϴ����е������ϴ�������
<li>����д�������ϴ�Ŀ¼��Ĭ�ϸ�Ŀ¼Ϊ����UploadFile����
<li>Ŀ¼��ʽ�涨���꣭�£��磺2003-8)��
</td></tr>
<tr><td colspan=5 height=25 class="forumRow">��Ҫ�������ϴ�Ŀ¼��
<INPUT TYPE="text" NAME="path" Id="path" value="<%=path%>">
<select onchange="Changepath(this.options[this.selectedIndex].value)">
<option value="UploadFile">ѡȡ��Ҫ������Ŀ¼</option>
<%
Dim uploadpath,ii
for ii=0 to datediff("m","2003-8",now())
uploadpath=DateAdd("m",-ii,now())
uploadpath=year(uploadpath)&"-"&month(uploadpath)
response.write "<option value="""&uploadpath&""">"&uploadpath&"</option>"
next
%>
</select>
</td>
<td colspan=3 height=25 class="forumRow">
<input type="submit" name="Submit" value="���δ��¼�ļ�" onclick="{if(confirm('��ȷ��ִ�еĲ�����?��ɾ������δ�м�¼���ϴ��ļ�,�����ָܻ���')){this.document.formpost.submit();return true;}return false;}">
</td></tr>
</form>
<SCRIPT LANGUAGE="JavaScript">
<!--
function Changepath(addTitle) {
document.getElementById("path").value=addTitle; 
document.getElementById("path").focus(); 
return; }
//-->
</SCRIPT>
<%
Response.Write "<tr><td class=""forumRowHighlight"" align=center colspan=8>"
call list()
Response.Write "</td></tr></table>"

end sub

SUB LIST()
'��ҳ����
If totalrec="" Then totalrec=0:Pcount=0
response.write "<table cellspacing=0 cellpadding=0 align=center width=""100%""><form method=post action=""?action=FileSearch"&seachstr&""" ><tr><td width=35% class=""forumRowHighlight"">��<b>"&totalrec&"</b>���ļ�������<b><font color=red>"&Pcount&"</font></b>ҳ��</td><td width=* valign=middle align=right nowrap class=""forumRowHighlight"">"

if currentpage > 4 then
	response.write "<a href=""?action=FileSearch&currentpage=1"&seachstr&""">[1]</a> ..."
end if
if Pcount>currentpage+3 then
	endpage=currentpage+3
else
	endpage=Pcount
end if
for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=red><b>["&i&"]</b></font>"
		else
        response.write " <a href=""?action=FileSearch&currentpage="&i&seachstr&""">["&i&"]</a>"
		end if
	end if
next
if currentpage+3 < Pcount then 
	response.write "... <a href=""?action=FileSearch&currentpage="&Pcount&seachstr&""">["&Pcount&"]</a>"
end if
response.write " ת��:<input type=text name=currentpage size=3 maxlength=10  value='"& currentpage &"'><input type=submit value=Go  id=button1 name=button1 >"     
response.write "</td></tr></form></table>"
END SUB

SUB delfiles()
Dim delid,F_filename
if instrRev(path,"/")=0 then path=path&"/"
response.write "<table cellspacing=1 cellpadding=3 align=center class=tableBorder width=""95%""><tr><td>"
delid=replace(Request.form("delid"),"'","")
if delid="" then 
response.write "��ѡ��Ҫɾ�����ļ���"
else
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select F_id,F_Filename from DV_Upfile where F_ID in ("&delid&")"
	rs.open sql,conn,1
	if not rs.eof then
	response.write "�ܹ�ɾ����¼���ļ�"&rs.recordcount&"����<br>"
	do while not rs.eof
		if InStr(rs(1),":")=0 or InStr(rs(1),"//")=0 then '�ж��ļ��Ƿ���̳������������ñ��еļ�¼��
			F_filename=path&rs(1)
		else
			F_filename=rs(1)
		end if
		if objFSO.fileExists(Server.MapPath(F_filename)) then
		objFSO.DeleteFile(Server.MapPath(F_filename))
		end if
		Dvbbs.Execute("delete from DV_Upfile where F_ID="&rs(0))
		response.write "�Ѿ�ɾ���ļ�"&F_filename&" ��<br>"
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
set objFSO=nothing
end if
response.write "</td></tr></table>"
END SUB

'�������м�¼
sub delall()
Server.ScriptTimeout=9999999
response.write "<table cellspacing=1 cellpadding=3 align=center class=tableBorder width=""95%""><tr><td>"
Dim TempFileName
Dim F_ID,F_AnnounceID,F_boardid,F_filename
Dim S_AnnounceID,s_Rootid
Dim drs,delfile
Dim delinfo
delfile=trim(Request.form("delfile"))
if cint(delfile)=1 then
delinfo="�ѱ�ɾ����"
else
delinfo="δ��ɾ����"
end if

if Request.form("delsql")<>"" then
	If Dvbbs.chkpost=False Then
		Dvbbs.AddErrmsg "���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ���ԡ�"
		exit sub
		else
		delsql=Request.form("delsql")
	End If
end if
i=0
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if delsql="" then
set rs=Dvbbs.Execute("select F_ID,F_AnnounceID,F_BoardID,F_Filename,F_Type from [DV_Upfile] where F_Flag=0 order by F_ID desc ")
else
set rs=Dvbbs.Execute(delsql)
end if
'response.write delsql
if rs.eof then
	response.write "��δ��"
else
	do while not rs.eof
	F_ID=rs(0)
	F_boardid=rs(2)
	if InStr(rs(3),":")=0 or InStr(rs(3),"//")=0 then '�ж��ļ��Ƿ���̳������������ñ��еļ�¼��
		F_filename="UploadFile/"&rs(3)
	else
		F_filename=rs(3)
	end if
	'Response.Write Rs("F_Type")&"<br>"
	If Rs("F_Type")<>1 Then		'��ͼƬ�ļ���
		TempFileName="viewfile.asp?ID="&F_ID
	Else
		TempFileName=F_filename
	End If
	TempFileName=Lcase(TempFileName)
	if rs(1)="" or isnull(rs(1)) then
		if InStr(rs(3),":")=0 or InStr(rs(3),"//")=0 then '�ж��ļ��Ƿ���̳������������ñ��еļ�¼��
			if objFSO.fileExists(Server.MapPath(F_filename)) then
				if delfile=1 then
					Dvbbs.Execute("delete from DV_Upfile where F_ID="&F_ID)
					objFSO.DeleteFile(Server.MapPath(F_filename))
				end if
				response.write "�ļ�δд����,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"<br>"
			else
				response.write "�ļ�δд����,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> �Ѳ����ڣ�<br>"
			end if
		else
			response.write "�ⲿ�ļ�<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"<br>"
		end if
		i=i+1
	else
		if isnumeric(rs(1)) then
			S_AnnounceID=rs(1)
		else
			F_AnnounceID=split(rs(1),"|")
			s_Rootid=F_AnnounceID(0)
			S_AnnounceID=F_AnnounceID(1)
		end if
		'Response.Write rs(1)&"<br>"
		If S_AnnounceID="" Then
			Response.Write F_filename &"�ļ�����������<br>"
		Else
		'ȡ���������ӱ���
		Dim PostTablename
		set drs=Dvbbs.Execute("select PostTable from dv_topic where TopicID="&s_Rootid)
			if not drs.eof then
			PostTablename=drs(0)
			else
			PostTablename=AllPostTable(0)
			end if
		drs.close

		'�ҳ���Ӧ�����ӽ����ж��ļ��Ƿ������������
		'Response.Write "select body from "&PostTablename&" where AnnounceID="&S_AnnounceID&"<br>"
		set drs=Dvbbs.Execute("select body from "&PostTablename&" where AnnounceID="&S_AnnounceID)
		if drs.eof then
			if delfile=1 then
			Dvbbs.Execute("delete from DV_Upfile where F_ID="&F_ID)
			end if
			if objFSO.fileExists(Server.MapPath(F_filename)) then
				if delfile=1 then
				objFSO.DeleteFile(Server.MapPath(F_filename))
				end if
				response.write "����δ�ҵ�,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"<br>"
			else
				response.write "����δ�ҵ�,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> �Ѳ����ڣ�<br>"
			end if
			i=i+1
		else
			'Response.Write TempFileName&"<br>"
			If Instr(Lcase(drs(0)),TempFileName)=0 Then
				if objFSO.fileExists(Server.MapPath(F_filename)) then
					if delfile=1 then
						objFSO.DeleteFile(Server.MapPath(F_filename))
						Dvbbs.Execute("delete from DV_Upfile where F_ID="&F_ID)
					end if
					response.write "�������ݲ���,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> "&delinfo&"[<a href=""dispbbs.asp?Boardid="&F_boardid&"&ID="&s_Rootid&"&replyID="&S_AnnounceID&"&skin=1"" target=""_blank"" title=""����������""><font color=red>�鿴�������</font></a> | <a href=myfile.asp?action=edit&editid="&F_ID&" target=""_blank"" title=""�༭�ļ�""><font color=red>�༭</font></a>]<br>"
				else
					response.write "�������ݲ���,<a href="&F_filename&" target=""_blank"">"&F_filename&"</a> �Ѳ����ڣ�[<a href=""dispbbs.asp?Boardid="&F_boardid&"&ID="&s_Rootid&"&replyID="&S_AnnounceID&"&skin=1"" target=""_blank"" title=""����������""><font color=red>�鿴�������</font></a> | <a href=myfile.asp?action=edit&editid="&F_ID&" target=""_blank"" title=""�༭�ļ�""><font color=red>�༭</font></a>]<br>"
				end if
				i=i+1
			end if
		end if
		drs.close
		End If
	End If
rs.movenext
loop
end if
rs.close
set drs=nothing
set rs=nothing
set objFSO=nothing

response.write"��������"&i&"���������ļ� ��<a href=?path="&path&" >����</a>��"
response.write "</td></tr></table>"
end sub


'ɾ������δ��¼���ϴ����е��ļ�
sub delall1()
response.write "<table cellspacing=1 cellpadding=3 align=center class=tableBorder width=""95%""><tr><td>"
Dim delfile,delinfo,datepath
delfile=dvbbs.checkStr(trim(Request.form("delfile")))
if cint(delfile)=1 then
	delinfo="Ŀǰ�ѱ�ɾ����"
else
	delinfo="Ŀǰδ��ɾ����"
end if

if instrRev(path,"/")=0 then path=path&"/"
If instr(path,"UploadFile")=0 Then
datepath=path
path="UploadFile/"&path
End If

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objFSO.FolderExists(Server.MapPath(path))=false then
	response.write "·����"&Path&"�����ڣ�"
else
	Set uploadFolder=objFSO.GetFolder(Server.MapPath(path))
	Set uploadFiles=uploadFolder.Files
	i=0
	For Each Upname In uploadFiles
		upfilename=path&upname.name
		'Response.Write "select top 1 F_ID from DV_Upfile where F_Filename = '"&datepath&upname.name&"'<br>"
		set rs=Dvbbs.Execute("select top 1 F_ID from DV_Upfile where F_Filename = '"&datepath&upname.name&"'")
		if rs.eof then
			i=i+1
			if delfile=1 then
			objFSO.DeleteFile(Server.MapPath(upfilename))
			end if
			response.write "<a href="&upfilename&" target=""_blank"">"
			response.write upfilename&"</a>�ڿ���û�м�¼��"&delinfo&"<br>"
		end if
		rs.close
		set rs=nothing
	next
	response.write"��ɾ����"&i&"���������ļ� ��<a href=?path="&path&" >����</a>��"
	set uploadFolder=nothing
	set uploadFiles=nothing
end if
set objFSO=nothing
response.write "</td></tr></table>"
end sub

function folder(path)
on error resume  next
       Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
          Set uploadFolder=objFSO.GetFolder(Server.MapPath(path))
		  if err.number<>"0" then
		  response.write Err.Description
		  response.end
		  end if
          For Each UpFolder In uploadFolder.SubFolders
            response.write "��<A HREF=?path="&path&"/"&upfolder.name&" >"&upfolder.name&"</a>�� | "
next
set uploadFolder=nothing
end function

function procGetFormat(sName)
 Dim str
 procGetFormat=0
 if instrRev(sName,".")=0 then exit function
 str=lcase(mid(sName,instrRev(sName,".")+1))
 for i=0 to uBound(sFor,1)
  if str=sFor(i,0) then 
    procGetFormat=sFor(i,1)
    exit for
  end if
 next
end function

function filetypename(stype)
if isempty(stype) or not isnumeric(stype) then exit function
select case cint(stype)
case 1
filetypename="ͼƬ��"
case 2
filetypename="FLASH��"
case 3
filetypename="���ּ�"
case 4
filetypename="��Ӱ��"
case else
filetypename="�ļ���"
end select 
end function

function getsize(size)
if isEmpty(size) then exit function
	if size>1024 then
 		   size=(size\1024)
 		   getsize=size & "&nbsp;KB"
	else
		   getsize=size & "&nbsp;B"
 	end if
 	if size>1024 then
 		   size=(size/1024)
 		   getsize=formatnumber(size,2) & "&nbsp;MB"		
 	end if
 	if size>1024 then
 		   size=(size/1024)
 		   getsize=formatnumber(size,2) & "&nbsp;GB"	   
 	end if   
end function
%>