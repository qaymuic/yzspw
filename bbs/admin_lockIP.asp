<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
	dim admin_flag
	admin_flag=",28,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		call main()
	end if

sub main()
dim userip,ips,GetIp1,GetIp2
if request("userip")<>"" then
userip=request("userip")
ips=Split(userIP,".")
If Ubound(ips)=3 Then GetIp1=ips(0)&"."&ips(1)&"."&ips(2)&".*"
end if
if request("action")="add" or request("userip")<>"" then
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2>IP���ƹ��������</th>
</tr>
<%
dim sip,str1,str2,str3,str4,num_1,num_2
if request.querystring("reaction")="save" then
	sip=cstr(request.form("ip1"))
	If sip<>"" Then
		If Trim(Dvbbs.cachedata(25,0))<>"" Then
			sip=Trim(Dvbbs.cachedata(25,0)) & "|" & Replace(sip,"|","")
		End If
	End If
	if sip<>"" then
		dvbbs.execute("update dv_setup set Forum_LockIP='"&replace(sip,"'","''")&"'")
		Dvbbs.Name="setup"
		dvbbs.reloadsetup
	end if
	Footer()
%>
<tr>
<td width="100%" colspan=2 class=forumrow>��ӳɹ���</td>
</tr>
<%
else
%>
<form action="admin_LockIP.asp?action=add&reaction=save" method="post">
<tr>
<td width="100%" class=forumrow colspan=2><B>˵��</B>����������Ӷ������IP��ÿ��IP��|�ŷָ�������IP����д��ʽ��202.152.12.1��������202.152.12.1���IP�ķ��ʣ���202.152.12.*����������202.152.12��ͷ��IP���ʣ�ͬ��*.*.*.*������������IP�ķ��ʡ�����Ӷ��IP��ʱ����ע�����һ��IP�ĺ��治Ҫ��|�������</td>
</tr>
<tr>
<td width="20%" class=forumrow>����I&nbsp;P</td>
<td width="80%" class=forumrow><input type="text" name="ip1" size="30" value="<%=GetIp1%>">&nbsp;��202.152.12.*</td>
</tr>
<tr>
<td width="20%" class=forumrow></td>
<td width="80%" class=forumrow>
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</form>
<%
end if
elseif request("action")="delip" then
	userip=request("ips")
	'userip=split(userip,chr(10))
	userip=split(userip,vbCrLf)
	for i = 0 to ubound(userip)
		if not (userip(i)="" or userip(i)=" ") then
			If i=0 Then
				getip1 = userip(i)
			Else
				getip1 = getip1 & "|" & userip(i)
			End If
		End If
	next
	dvbbs.execute("update dv_setup set forum_lockip='"&replace(getip1,"'","''")&"'")
	Dvbbs.Name="setup"
	dvbbs.reloadsetup
	response.write "��������IP�ɹ���"
else
%>
<FORM METHOD=POST ACTION="?action=delip">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2>IP���ƹ���������</th>
</tr>
<tr>
<td width="100%" class=forumrow colspan=2>
<B>˵��</B>����������Ӷ������IP��ÿ��IP�ûس��ָ�������IP����д��ʽ��202.152.12.1��������202.152.12.1���IP�ķ��ʣ���202.152.12.*����������202.152.12��ͷ��IP���ʣ�ͬ��*.*.*.*������������IP�ķ��ʡ�����Ӷ��IP��ʱ����ע�����һ��IP�ĺ��治Ҫ�ӻس�
</td>
</tr>
<tr>
<td width="100%" class=forumrow colspan=2>
<textarea name="ips" cols="80" rows="8"><%
userip=split(Trim(Dvbbs.cachedata(25,0)),"|")
For i = 0 To Ubound(Userip)
	Response.Write Userip(i)
	If i < Ubound(Userip) Then Response.Write Chr(10)
Next
%></textarea>
</td>
</tr>
<tr>
<td width="20%" class=forumrow></td>
<td width="80%" class=forumrow>
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
<%
	Footer()
%>
</FORM>
</table>
<%
end if
end Sub
%>