<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
	dim admin_flag
	admin_flag=",5,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if request("action")="save" then
		call savegrade()
		else
		call grade()
		end if
		If  founderr Then dvbbs_error()
		Footer()
	end if

sub grade()
dim sel
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0"  align=center>
<tr> 
<th height="23" colspan="2" >�û���������</th>
</tr>
<tr> 
<td width="100%" class=ForumrowHighlight colspan=2>
<B>˵��</B>��<BR>1����ѡ����ѡ���Ϊ��ǰ��ʹ������ģ�壬����ɲ鿴��ģ�����ã�������ģ��ֱ�Ӳ鿴��ģ�岢�޸����á������Խ�����������ñ����ڶ����̳������<BR>2����Ҳ���Խ������趨����Ϣ���沢Ӧ�õ�����ķ���̳���������У��ɶ�ѡ<BR>3�����������һ���������ñ�İ�������ã�ֻҪ����ð������ƣ������ʱ��ѡ��Ҫ���浽�İ����������Ƽ��ɡ�<BR>
4��Ĭ��ģ���еĻ�������Ϊ��̳����ҳ�棨<font color=blue>�������������̳����</font>��ʹ�ã����¼��ע�����ط�ֵ���������̳��������в�ͬ�Ļ������ã��緢����ɾ���ȣ���Ȼ��Ҳ���Ը���������趨�����趨���а���Ļ������ö���һ���ġ�
</td>
</tr>
<FORM METHOD=POST ACTION="">
<tr> 
<td width="100%" class="forumRowHighlight" colspan=2>
�鿴�ְ���������ã���ѡ�������������Ӧ����&nbsp;&nbsp;
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option value="">�鿴�ְ�������ѡ��</option>
<%
Dim ii
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=dvbbs.boardid then
Response.Write " selected"
end if
Response.Write " value=""admin_wealth.asp?boardid="&rs(0)&""">"
Select Case rs(2)
	Case 0
		Response.Write "��"
	Case 1
		Response.Write "&nbsp;&nbsp;��"
End Select
If rs(2)>1 Then
	For ii=2 To rs(2)
		Response.Write "&nbsp;&nbsp;��"
	Next
	Response.Write "&nbsp;&nbsp;��"
End If
Response.Write rs(1)
Response.Write "</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
</FORM>
</table><BR>

<form method="POST" action=admin_wealth.asp?action=save>
<table width="95%" border="0" cellspacing="0" cellpadding="0"  align=center>

<tr> 
<td width="100%" class=ForumrowHighlight colspan=2>
<input type=checkbox name="getskinid" value="1" <%if request("getskinid")="1" or request("boardid")="" then Response.Write "checked"%>><a href="admin_wealth.asp?getskinid=1">��̳Ĭ�ϻ���</a><BR> ����˴�������̳Ĭ�ϻ������ã�Ĭ�ϻ������ð�������<FONT COLOR="blue">��</FONT>��������������ݣ��緢���������������ȣ�<FONT COLOR="blue">����</FONT>��ҳ�档<hr size=1 width="90%" color=blue>
</td>
</tr>
<tr>
<td width="200" class="forumrow">
�����汣��ѡ��<BR>
�밴 CTRL ����ѡ<BR>
<select name="getboard" size="40" style="width:100%" multiple>
<%
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=dvbbs.boardid then
Response.Write " selected"
end if
Response.Write " value="&rs(0)&">"
Select Case rs(2)
	Case 0
		Response.Write "��"
	Case 1
		Response.Write "&nbsp;&nbsp;��"
End Select
If rs(2)>1 Then
	For ii=2 To rs(2)
		Response.Write "&nbsp;&nbsp;��"
	Next
	Response.Write "&nbsp;&nbsp;��"
End If
Response.Write rs(1)
Response.Write "</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
<td class="forumrow" valign=top>
<table width=100% >
<tr> 
<th height="23" colspan="2">�û���Ǯ�趨</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>ע���Ǯ��</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthReg" size="35" value="<%=Dvbbs.Forum_user(0)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��¼���ӽ�Ǯ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthLogin" size="35" value="<%=Dvbbs.Forum_user(4)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������ӽ�Ǯ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthAnnounce" size="35" value="<%=Dvbbs.Forum_user(1)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������ӽ�Ǯ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthReannounce" size="35" value="<%=Dvbbs.Forum_user(2)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������ӽ�Ǯ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="BestWealth" size="35" value="<%=Dvbbs.Forum_user(15)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ɾ�����ٽ�Ǯ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="wealthDel" size="35" value="<%=Dvbbs.Forum_user(3)%>">
</td>
</tr>
<tr> 
<th height="23" colspan="2" >�û������趨</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>ע�ᾭ��ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epReg" size="35" value="<%=Dvbbs.Forum_user(5)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��¼���Ӿ���ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epLogin" size="35" value="<%=Dvbbs.Forum_user(9)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������Ӿ���ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epAnnounce" size="35" value="<%=Dvbbs.Forum_user(6)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������Ӿ���ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epReannounce" size="35" value="<%=Dvbbs.Forum_user(7)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>�������Ӿ���ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="bestuserep" size="35" value="<%=Dvbbs.Forum_user(17)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ɾ�����پ���ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="epDel" size="35" value="<%=Dvbbs.Forum_user(8)%>">
</td>
</tr>
<tr> 
<th height="23" colspan="2" >�û������趨</th>
</tr>
<tr> 
<td width="40%" class=Forumrow>ע������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpReg" size="35" value="<%=Dvbbs.Forum_user(10)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>��¼��������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpLogin" size="35" value="<%=Dvbbs.Forum_user(14)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>������������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpAnnounce" size="35" value="<%=Dvbbs.Forum_user(11)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>������������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpReannounce" size="35" value="<%=Dvbbs.Forum_user(12)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>������������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="bestusercp" size="35" value="<%=Dvbbs.Forum_user(16)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>ɾ����������ֵ</td>
<td width="60%" class=Forumrow> 
<input type="text" name="cpDel" size="35" value="<%=Dvbbs.Forum_user(13)%>">
</td>
</tr>
<tr> 
<td width="40%" class=Forumrow>&nbsp;</td>
<td width="60%" class=Forumrow> 
<div align="center"> 
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
<%
end sub

sub savegrade()
dim Forum_user,iforum_setting,forum_setting
Forum_user=request.form("wealthReg") & "," & request.form("wealthAnnounce") & "," & request.form("wealthReannounce") & "," & request.form("wealthDel") & "," & request.form("wealthLogin") & "," & request.form("epReg") & "," & request.form("epAnnounce") & "," & request.form("epReannounce") & "," & request.form("epDel") & "," & request.form("epLogin") & "," & request.form("cpReg") & "," & request.form("cpAnnounce") & "," & request.form("cpReannounce") & "," & request.form("cpDel") & "," & request.form("cpLogin") & "," & request.form("BestWealth") & "," & request.form("BestuserCP") & "," & request.form("BestuserEP")
'response.write Forum_user

'forum_info|||forum_setting|||forum_user|||copyright|||splitword|||stopreadme
Set rs=Dvbbs.execute("select forum_setting from dv_setup")
iforum_setting=split(rs(0),"|||")
forum_setting=iforum_setting(0) & "|||" & iforum_setting(1) & "|||" & forum_user & "|||" & iforum_setting(3) & "|||" & iforum_setting(4) & "|||" & iforum_setting(5)
forum_setting=dvbbs.checkstr(forum_setting)

if request("getskinid")="1" then
sql = "update dv_setup set Forum_setting='"&forum_setting&"'"
Dvbbs.Execute(sql)
Dvbbs.Name="setup"
Dvbbs.ReloadSetup
end if
if request("getboard")<>"" then
sql = "update dv_board set board_user='"&Forum_user&"' where boardid in ("&request("getboard")&")"
Dvbbs.Execute(sql)
Dim SplitBoardID
SplitBoardID=Split(Request("getboard"),",")
For i=0 To Ubound(SplitBoardID)
	If IsNumeric(SplitBoardID(i)) And SplitBoardID(i)<>"" Then
		Dvbbs.ReloadBoardCache Clng(SplitBoardID(i)),Forum_user,18,0
	End If
Next
end if
Dv_suc("��̳�������óɹ���")
End  Sub 

%>