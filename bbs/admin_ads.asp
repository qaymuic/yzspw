<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
	dim admin_flag
	admin_flag=",2,"
	If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	Else 
		If request("action")="save" Then 
		call saveconst()
		else
		call consted()
		end if
		If founderr then call dvbbs_error()
		footer()
	End If 

sub consted()
dim sel
%>
<table width="95%" border="0" cellspacing="0" cellpadding="3"  align=center class="tableBorder">

<tr> 
<th height="23" colspan="2" class="tableHeaderText"><b>��̳�������</b>����Ϊ���÷���̳�����Ƿ���̳��ҳ��棬����ҳ��Ϊ������ʾҳ�棩</th>
</tr>
<tr> 
<td width="100%" class="forumRowHighlight" colspan=2><B>˵��</B>��<BR>1����ѡ����ѡ���Ϊ��ǰ��ʹ������ģ�壬����ɲ鿴��ģ�����ã�������ģ��ֱ�Ӳ鿴��ģ�岢�޸����á������Խ�����������ñ����ڶ����̳������<BR>2����Ҳ���Խ������趨����Ϣ���沢Ӧ�õ�����ķ���̳���������У��ɶ�ѡ<BR>3�����������һ���������ñ�İ�������ã�ֻҪ����ð������ƣ������ʱ��ѡ��Ҫ���浽�İ����������Ƽ��ɡ�
<hr size=1 width="90%" color=blue>
</td>
</tr>
<FORM METHOD=POST ACTION="">
<tr> 
<td width="100%" class="forumRowHighlight" colspan=2>
�鿴�ְ��������ã���ѡ�������������Ӧ����&nbsp;&nbsp;
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
Response.Write " value=""admin_ads.asp?boardid="&rs(0)&""">"
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
<form method="POST" action=admin_ads.asp?action=save>
<table width="95%" border="0" cellspacing="0" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td width="100%" class="forumRowHighlight" colspan=2>
<input type=checkbox name="getskinid" value="1" <%if request("getskinid")="1" or request("boardid")="" then Response.Write "checked"%>><a href="admin_ads.asp?getskinid=1">��̳Ĭ�Ϲ��</a><BR> ����˴�������̳Ĭ�Ϲ�����ã�Ĭ�Ϲ�����ð�������<FONT COLOR="blue">��</FONT>��������������ݣ��������б�������ʾ�����澫�������淢���ȣ�<FONT COLOR="blue">����</FONT>��ҳ�档<hr size=1 width="90%" color=blue>
</td>
</tr>
<tr> 
<td width="200" class="forumrow">
�����汣��ѡ��<BR>
�밴 CTRL ����ѡ<BR>
<select name="getboard" size="28" style="width:100%" multiple>
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
<table>
<tr>
<td width="200" class="forumrow"><B>��ҳ����������</B><BR>��������˻�����湦���еĶ�����棬�˴�����Ϊ��Ч</td>
<td width="*" class="forumrow"> 
<textarea name="index_ad_t" cols="50" rows="3"><%=server.htmlencode(Dvbbs.Forum_ads(0))%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��ҳβ��������</B></font></td>
<td width="*" class="forumrow"> 
<textarea name="index_ad_f" cols="50" rows="3"><%=server.htmlencode(Dvbbs.Forum_ads(1))%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>������ҳ�������</B></font></td>
<td width="*" class="forumrow"> 
<input type=radio name="index_moveFlag" value=0 <%if Dvbbs.Forum_ads(2)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="index_moveFlag" value=1 <%if Dvbbs.Forum_ads(2)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ�������ͼƬ��ַ</B></font></td>
<td width="*" class="forumrow"> 
<input type="text" name="MovePic" size="35" value="<%=Dvbbs.Forum_ads(3)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ����������ӵ�ַ</B></font></td>
<td width="*" class="forumrow"> 
<input type="text" name="MoveUrl" size="35" value="<%=Dvbbs.Forum_ads(4)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ�������ͼƬ���</B></font></td>
<td width="*" class="forumrow"> 
<input type="text" name="move_w" size="3" value="<%=Dvbbs.Forum_ads(5)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ�������ͼƬ�߶�</B></font></td>
<td width="*" class="forumrow"> 
<input type="text" name="move_h" size="3" value="<%=Dvbbs.Forum_ads(6)%>">&nbsp;����
</td>
</tr>
<input type=hidden name="Board_moveFlag" value=0>
<tr> 
<td width="200" class="forumrow"><B>������ҳ���¹̶����</B></font></td>
<td width="*" class="forumrow"> 
<input type=radio name="index_fixupFlag" value=0 <%if Dvbbs.Forum_ads(13)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="index_fixupFlag" value=1 <%if Dvbbs.Forum_ads(13)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ���¹̶����ͼƬ��ַ</B></font></td>
<td width="*" class="forumrow"> 
<input type="text" name="fixupPic" size="35" value="<%=Dvbbs.Forum_ads(8)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ���¹̶�������ӵ�ַ</B></font></td>
<td width="*" class="forumrow"> 
<input type="text" name="fixupUrl" size="35" value="<%=Dvbbs.Forum_ads(9)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ���¹̶����ͼƬ���</B></font></td>
<td width="*" class="forumrow"> 
<input type="text" name="fixup_w" size="3" value="<%=Dvbbs.Forum_ads(10)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ���¹̶����ͼƬ�߶�</B></font></td>
<td width="*" class="forumrow"> 
<input type="text" name="fixup_h" size="3" value="<%=Dvbbs.Forum_ads(11)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="*" class="forumrow" valign="top" colspan=2><B>��̳�������������</B></font> <br>֧��HTML�﷨��ÿ��������һ�У��ûس��ֿ���</td>
</tr>
<tr>
<td width="*" class="forumrow" colspan=2> 
<textarea name="Forum_ads(14)" style="width:100%" rows="10"><%If UBound(Dvbbs.Forum_ads)>13 Then
	Response.Write Dvbbs.Forum_ads(14)
End If	
%></textarea>
</td>
</tr>
<input type=hidden name="Board_fixupFlag" value=0>
<tr> 
<td width="200" class="forumrow">&nbsp;</td>
<td width="*" class="forumrow"> 
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

sub saveconst()
Dvbbs.Forum_ads=request("index_ad_t") & "$" & request("index_ad_f") & "$" & request("index_moveFlag") & "$" & request("MovePic") & "$" & request("MoveUrl") & "$" & request("move_w") & "$" & request("move_h") & "$" & request("Board_moveFlag") & "$" & request("fixupPic") & "$" & request("FixupUrl") & "$" & request("Fixup_w") & "$" & request("Fixup_h") & "$" & request("Board_fixupFlag") & "$" & request("index_fixupFlag") & "$"&Request("Forum_ads(14)")

if request("getskinid")="1" then
sql = "update dv_setup set Forum_ads='"&Replace(Dvbbs.Forum_ads,"'","''")&"'"
Dvbbs.Execute(sql)
Dvbbs.ReloadSetupCache Dvbbs.Forum_ads,2
end if
if request("getboard")<>"" then
sql = "update dv_board set board_ads='"&Replace(Dvbbs.Forum_ads,"'","''")&"' where boardid in ("&request("getboard")&")"
Dvbbs.Execute(sql)
Dim SplitBoardID
SplitBoardID=Split(Request("getboard"),",")
For i=0 To Ubound(SplitBoardID)
	If IsNumeric(SplitBoardID(i)) And SplitBoardID(i)<>"" Then
		Dvbbs.ReloadBoardCache Clng(SplitBoardID(i)),Dvbbs.Forum_ads,17,0
	End If
Next
end if
Dv_suc("������óɹ���")
End sub
%>