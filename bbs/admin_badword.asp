<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->

<%
	Head()
	dim admin_flag
	admin_flag="26,27"
	if not Dvbbs.master or instr(","&session("flag")&",",",26,")=0 or instr(","&session("flag")&",",",27,")=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		call main()
	Footer()
	end if

	sub main()
	dim sel

if request("action") = "savebadword" then
call savebadword()
else

%>

<form action="admin_badword.asp?action=savebadword" method=post>

<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">

<%if request("reaction")="badword" then%>
<tr>
<th colspan=2 align=left height=23>���ӹ����ַ�</th>
</tr>
<tr>
<td class=forumrow width="100%" colspan=2><B>˵��</B>�������ַ��趨����Ϊ  <B>Ҫ���˵��ַ�=���˺���ַ�</B> ��ÿ�������ַ��ûس��ָ��</td>
</tr>
<tr>
<td class=forumrow width="100%" colspan=2>
<textarea name="badwords" cols="80" rows="8"><%
For i=0 To Ubound(Dvbbs.BadWords)
	If i > UBound(Dvbbs.rBadWord) Then
		Response.Write Dvbbs.BadWords(i) & "=*"
	Else
		Response.Write Dvbbs.BadWords(i) & "=" & Dvbbs.rBadWord(i)
	End If
	If i<Ubound(Dvbbs.BadWords) Then Response.Write chr(10)
Next
%></textarea>
<!--<input type="text" name="badwords" value="" size="80">--></td>
</tr>
<%elseif request("reaction")="splitreg" then%>
<tr>
<th colspan=2 align=left height=23>ע������ַ�</th>
</tr>
<tr>
<td class=forumrow width="20%">˵����</td>
<td class=forumrow width="80%">ע������ַ����������û�ע����������ַ������ݣ�������Ҫ���˵��ַ������룬����ж���ַ��������á�,���ָ��������磺ɳ̲,quest,ľ��</td>
</tr>
<tr>
<td class=forumrow width="20%">����������ַ�</td>
<td class=forumrow width="80%"><input type="text" name="splitwords" value="<%=split(Dvbbs.cachedata(1,0),"|||")(4)%>" size="80"></td>
</tr>
<%end if%>
<input type=hidden value="<%=request("reaction")%>" name="reaction">
<tr> 
<td class=forumrow width="20%"></td>
<td width="80%" class=forumrow><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</table>
</form>
<%end if%>
<%
end sub

sub savebadword()
dim iforum_setting,forum_setting
If request("reaction")="badword" then
dim badwords,badwords_1,badwords_2,badwords_3
badwords=request("badwords")
badwords=split(badwords,vbCrlf)
for i = 0 to ubound(badwords)
	if not (badwords(i)="" or badwords(i)=" ") then
		badwords_1 = split(badwords(i),"=")
		If ubound(badwords_1)=1 Then
			If i=0 Then
				badwords_2 = badwords_1(0)
				badwords_3 = badwords_1(1)
			Else
				badwords_2 = badwords_2 & "|" & badwords_1(0)
				badwords_3 = badwords_3 & "|" & badwords_1(1)
			End If
		End If
	End If
next

sql = "update dv_setup set Forum_Badwords='"&replace(badwords_2,"'","''")&"',Forum_rBadword='"&replace(badwords_3,"'","''")&"'"
dvbbs.execute(sql)
elseif request("reaction")="splitreg" then
'forum_info|||forum_setting|||forum_user|||copyright|||splitword|||stopreadme
Set rs=Dvbbs.execute("select forum_setting from dv_setup")
iforum_setting=split(rs(0),"|||")
forum_setting=iforum_setting(0) & "|||" & iforum_setting(1) & "|||" & iforum_setting(2) & "|||" & iforum_setting(3) & "|||" & request("splitwords") & "|||" & iforum_setting(5)
sql = "update dv_setup set forum_setting='"&replace(forum_setting,"'","''")&"'"
dvbbs.execute(sql)
end if
Dvbbs.Name="setup"
Dvbbs.ReloadSetup
response.write "���³ɹ���"
end sub

%>