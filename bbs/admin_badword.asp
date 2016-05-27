<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->

<%
	Head()
	dim admin_flag
	admin_flag="26,27"
	if not Dvbbs.master or instr(","&session("flag")&",",",26,")=0 or instr(","&session("flag")&",",",27,")=0 then
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
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
<th colspan=2 align=left height=23>帖子过滤字符</th>
</tr>
<tr>
<td class=forumrow width="100%" colspan=2><B>说明</B>：过滤字符设定规则为  <B>要过滤的字符=过滤后的字符</B> ，每个过滤字符用回车分割开。</td>
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
<th colspan=2 align=left height=23>注册过滤字符</th>
</tr>
<tr>
<td class=forumrow width="20%">说明：</td>
<td class=forumrow width="80%">注册过滤字符将不允许用户注册包含以下字符的内容，请您将要过滤的字符串添入，如果有多个字符串，请用“,”分隔开，例如：沙滩,quest,木鸟</td>
</tr>
<tr>
<td class=forumrow width="20%">请输入过滤字符</td>
<td class=forumrow width="80%"><input type="text" name="splitwords" value="<%=split(Dvbbs.cachedata(1,0),"|||")(4)%>" size="80"></td>
</tr>
<%end if%>
<input type=hidden value="<%=request("reaction")%>" name="reaction">
<tr> 
<td class=forumrow width="20%"></td>
<td width="80%" class=forumrow><input type="submit" name="Submit" value="提 交"></td>
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
response.write "更新成功！"
end sub

%>