<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dvbbs.stats="��������ҳ��"
Dvbbs.LoadTemplates("")
Dvbbs.Nav()
Dim sql1,rs1,sql,Rs,i
If Dvbbs.UserID=0 Then Response.redirect "showerr.asp?ErrCodes=<li>���¼����в�����&action=OtherErr"

If DVbbs.BoardID=0 then
	Dvbbs.Head_var 2,0,"",""
Else
	Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
	GetBoardPermission
End If

If Not(Dvbbs.boardmaster or Dvbbs.master or Dvbbs.superboardmaster) Then Response.redirect "showerr.asp?ErrCodes=<li>ֻ�й���Ա���ܵ�¼��&action=OtherErr"
Main()

Dvbbs.Footer()

Sub Main()
%>
<TABLE cellpadding=0 cellspacing=1 class=tableborder1 align=center > 
        <tr >
          <th height=24 align=center colspan="2">��ӭ <%=Dvbbs.htmlencode(Dvbbs.membername)%>�����������ҳ��</th>
        </tr>
        <tr >
          <td height=24 align=center colspan="2" class=tablebody1>
        <b>����ѡ�<a href="announcements.asp?boardid=<%=Dvbbs.BoardID%>">���淢���͹���</a> | 
		<a href="admin_boardset.asp?action=editbminfo&boardid=<%=Dvbbs.BoardID%>">������Ϣ����</a> |  
				<a href="admin_boardset.asp?action=editbmads&boardid=<%=Dvbbs.BoardID%>">�ְ������</a> |  
				<a href="list.asp?action=batch&boardid=<%=Dvbbs.BoardID%>">������������</a>
		</b></td>
        </tr>
</table>
<BR>
<table cellpadding=0 cellspacing=0 width="<%=Dvbbs.mainsetting(0)%>" align=center style="word-break:break-all;">
		<tr>
              <td width="30%" valign=top align=left>
		<table cellpadding=3 cellspacing=1 class=tableborder1 style="width:100%;word-break:break-all;">
		<tr>
			<th width="100%" height=24  colspan="2">�� ������Ϣ�� ��
			</th>
		</tr>
		<tr>
			<td  height=24 class=tablebody2 colspan="2" align=center ><%=Dvbbs.BoardType%>
			</td>
		</tr>
		<tr>
			<td width="60%" height=24 class=tablebody1 >����������
			</td>
			<td width="40%" height=24 class=tablebody1 ><FONT COLOR=RED><%=Dvbbs.Board_Data(12,0)%></FONT>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody2 >�������ӣ�
			</td>
			<td  height=24 class=tablebody2 ><%=Dvbbs.Board_Data(10,0)%>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody1 >�������ӣ�
			</td>
			<td  height=24 class=tablebody1 ><%=Dvbbs.Board_Data(9,0)%>
			</td>
		</tr>
		<tr>
			<td width="30%" height=24 class=tablebody2 colspan="2">�����Ա��
		<%=Replace(Dvbbs.BoardMasterList&"","|",",")%>
			</td>
		</tr>
		<tr>
			<th width="100%" height=24  colspan="2">�� ����Ȩ�� ��
			</th>
		</tr>
		<tr>
			<td  height=24 class=tablebody1 >����������ɾ��������
			</td>
			<td  height=24 class=tablebody1 ><%if Dvbbs.Board_Setting(33)=1 then%>��<%else%><FONT COLOR=RED>�ر�</FONT><%end if%>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody2 >���������޸Ĺ�����ã�
			</td>
			<td  height=24 class=tablebody2 ><%if Dvbbs.Board_Setting(34)=1 then%>��<%else%><FONT COLOR=RED>�ر�</FONT><%end if%>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody1 >���а��������޸Ĺ�����ã�
			</td>
			<td  height=24 class=tablebody1 ><%if Dvbbs.Board_Setting(35)=1 then%>��<%else%><FONT COLOR=RED>�ر�</FONT><%end if%>
			</td>
		</tr>
		<tr>
			<td width="100%" height=24  colspan="2" class=tablebody2>
		<b>ע�⣺</b>������������������Լ��������ɷ�������Ͱ������ã�����Ա���������а��淢����������Ϣ���й��������
			</td>
		</tr>
		</table>
	      </td>
		  <td width="2%" valign=top align=center></td>
              <td width="70%" valign=top align=center>
      		<table cellpadding=3 cellspacing=1 class=tableborder1 style="width:100%;word-break:break-all;">
		  <tr>
			<td width="100%" height=24 class=tablebody1>
<B>ע��</B>��<BR>��ҳ��Ϊ����ר�ã�ʹ��ǰ�뿴������Ӧ�Ĺ����Ƿ�򿪣��ڽ��й������õ�ʱ�򣬲�Ҫ����������ã�������ģ�������д����������ȷ����д��
		  </td></tr>
		</table>
<%
Select Case request("action")
	Case "new"
		Call savenews()
	Case "manage"
		Call manage()
	Case "edit"
		Call Edit()
	Case "updat"
		Call Update()
	Case "del"
		Call del()
	Case "editbminfo"
		Call editbminfo()
	Case "saveditbm"
		Call savebminfo()
	Case "editbmads"
		Call editbmads()
	Case "savebmads"
		Call savebmads()
	Case  Else 
		'Call news()
End Select
%>
        </td>
    </tr>
</table>
<%
End Sub 

Sub editbmads()
dim master_1,chkedit
If Dvbbs.master Then
	chkedit=True
	Set Rs=Dvbbs.Execute("select boardmaster,Board_Ads from dv_board where boardid="&request("boardid"))
	If rs.eof and rs.bof Then Response.redirect "showerr.asp?ErrCodes=<li>��û��ָ����Ӧ��̳ID�����ܽ��й���&action=OtherErr"
	Dvbbs.Forum_Ads = Split(Rs(1),"$")
Else
	Set Rs=Dvbbs.Execute("select boardmaster,Board_Ads from dv_board where boardid="&request("boardid"))
	If rs.eof and rs.bof Then Response.redirect "showerr.asp?ErrCodes=<li>��û��ָ����Ӧ��̳ID�����ܽ��й���&action=OtherErr"
	If IsNull(Rs(0)) Then 
		Response.redirect "showerr.asp?ErrCodes=<li>����̳��δ�й���Ա��&action=OtherErr"
	Else
		master_1=split(rs(0),"|")
	End If
	Dvbbs.Forum_Ads = Split(Rs(1),"$")
	If Dvbbs.Board_Setting(35)=1 Then
		chkedit=True
	Else
		If Dvbbs.Board_Setting(34)=0 Then
			chkedit=False
		ElseIf Dvbbs.Board_Setting(34)=1 and Dvbbs.membername=master_1(0) Then
			chkedit=True
		Else
			chkedit=False
		End If
	End If
End If 

if chkedit=False Then
	Response.redirect "showerr.asp?ErrCodes=<li>�����Ϊ������ר�á�&action=OtherErr"
Else
%>
<form method="POST" action="?action=savebmads&boardid=<%=request("boardid")%>">
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:100%;word-break:break-all;">
<tr> 
<th height="23" colspan="2" class="tableHeaderText"><b>��̳�������</b>����Ϊ���÷���̳�����Ƿ���̳��ҳ��棬����ҳ��Ϊ������ʾҳ�棩</th>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��ҳ����������</B></font></td>
<td width="60%" class="Tablebody1"> 
<textarea name="index_ad_t" cols="50" rows="3"><%=enfixjs(Dvbbs.Forum_ads(0))%></textarea>
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��ҳβ��������</B></font></td>
<td width="60%" class="Tablebody1"> 
<textarea name="index_ad_f" cols="50" rows="3"><%=enfixjs(Dvbbs.Forum_ads(1))%></textarea>
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>������ҳ�������</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type=radio name="index_moveFlag" value=0 <%if Dvbbs.Forum_ads(2)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="index_moveFlag" value=1 <%if Dvbbs.Forum_ads(2)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ�������ͼƬ��ַ</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="MovePic" size="35" value="<%=Dvbbs.Forum_ads(3)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ����������ӵ�ַ</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="MoveUrl" size="35" value="<%=Dvbbs.Forum_ads(4)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ�������ͼƬ���</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="move_w" size="3" value="<%=Dvbbs.Forum_ads(5)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ�������ͼƬ�߶�</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="move_h" size="3" value="<%=Dvbbs.Forum_ads(6)%>">&nbsp;����
</td>
</tr>
<input type=hidden name="Board_moveFlag" value=0>
<tr> 
<td width="40%" class="Tablebody1"><B>������ҳ���¹̶����</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type=radio name="index_fixupFlag" value=0 <%if Dvbbs.Forum_ads(13)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="index_fixupFlag" value=1 <%if Dvbbs.Forum_ads(13)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ���¹̶����ͼƬ��ַ</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixupPic" size="35" value="<%=Dvbbs.Forum_ads(8)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ���¹̶�������ӵ�ַ</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixupUrl" size="35" value="<%=Dvbbs.Forum_ads(9)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ���¹̶����ͼƬ���</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixup_w" size="3" value="<%=Dvbbs.Forum_ads(10)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ���¹̶����ͼƬ�߶�</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixup_h" size="3" value="<%=Dvbbs.Forum_ads(11)%>">&nbsp;����
</td>
</tr>
<input type=hidden name="Board_fixupFlag" value=0>
<tr> 
<td width="40%" class="Tablebody1">&nbsp;</td>
<td width="60%" class="Tablebody1"> 
<div align="center"> 
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</table>
</form>
<%
end If
End Sub
Sub savebmads()
Dim master_1
Dim chkedit
Dim Forum_adsinfo
Set Rs=Dvbbs.Execute("select boardmaster from dv_board where boardid="&request("boardid"))
If rs.eof and rs.bof Then Response.redirect "showerr.asp?ErrCodes=<li>��û��ָ����Ӧ��̳ID�����ܽ��й���&action=OtherErr"
master_1=split(rs(0),"|")
If Dvbbs.Board_Setting(35)=1 Then
	chkedit=True
Else
	If Dvbbs.Board_Setting(34)=0 Then
		chkedit=False
	ElseIf Dvbbs.Board_Setting(34)=1 and Dvbbs.membername=master_1(0) Then
		chkedit=true
	Else
		chkedit=False
	End If
End If
If Dvbbs.master Then
	chkedit=true
end if
If chkedit=false Then
	Response.redirect "showerr.asp?ErrCodes=<li>�����Ϊ������ר�á�&action=OtherErr"
Else
	Forum_adsinfo=request("index_ad_t") & "$" & request("index_ad_f") & "$" & request("index_moveFlag") & "$" & request("MovePic") & "$" & request("MoveUrl") & "$" & request("move_w") & "$" & request("move_h") & "$" & request("Board_moveFlag") & "$" & request("fixupPic") & "$" & request("FixupUrl") & "$" & request("Fixup_w") & "$" & request("Fixup_h") & "$" & request("Board_fixupFlag") & "$" & request("index_fixupFlag")
	sql = "update dv_board set board_ads='"&Replace(Forum_adsinfo,"'","''")&"' where boardid="&Dvbbs.boardid&""
	Dvbbs.Execute(sql)
	Dvbbs.ReloadBoardInfo(Dvbbs.BoardID)
	response.write Dvbbs.BoardType&"������óɹ���"
End If
End Sub

Sub editbminfo()
If Not IsObject(Conn) Then ConnectionDatabase
Dim master_1
%>
<form action ="admin_boardset.asp?action=saveditbm&boardid=<%=Dvbbs.BoardID%>" method=post> 
<%
set rs= server.CreateObject("adodb.recordset")
sql = "select * from dv_board where boardid="&Dvbbs.boardid
rs.open sql,conn,1,1
If rs.eof and rs.bof Then Response.redirect "showerr.asp?ErrCodes=<li>��û��ָ����Ӧ��̳ID�����ܽ��й���&action=OtherErr"

If Not Dvbbs.master then
	If Dvbbs.Board_Setting(33)=1 Then
		master_1=rs("boardmaster")
		If Not IsNull(master_1) Then
			master_1=split(master_1,"|")
			If Dvbbs.membername<>master_1(0) Then Response.redirect "showerr.asp?ErrCodes=<li>�����Ϊ������ר�á�&action=OtherErr"
		Else
			Response.redirect "showerr.asp?ErrCodes=<li>�����Ϊ������ר�á�&action=OtherErr"
		End If
	Else
		Response.redirect "showerr.asp?ErrCodes=<li>��δ���޸����õ�Ȩ�ޡ�&action=OtherErr"
	End If
End If
%>
<Input type='hidden' name=editid value='<%=Dvbbs.BoardID%>'>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:100%;word-break:break-all;">
    <tr> 
    <th colspan="3" height=22 class=tablebody2><b>������Ϣ���� </b> 
 
  <tr> 
      <td height=22 class=tablebody1  align="center">��̳���ƣ�</td>
      <td  class=tablebody1>
	  <input type="text" name="BoardType" size="30" value='<%=enfixjs(rs("BoardType"))%>'>
	  </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody2  align="center">����˵����</td>
      <td  class=tablebody1>
      <textarea name="Readme" cols="80" rows="3"><%=enfixjs(rs("readme"))%></textarea>
      </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody1  align="center">�����޸ģ�</td>
      <td  class=tablebody1> 
        <input type="text" name="boardmaster" size="50" value='<%=rs("boardmaster")%>'><BR>(������������|�ָ����磺ɳ̲С��|wodeail)
      </td>
    </tr>
    <%If Cint(Dvbbs.Board_Setting(2))=1 Then%>
    <tr> 
      <td height=22 class=tablebody1  align="center">��֤�û���</td>
      <td  class=tablebody1> 
      <textarea name="boarduser" cols="80" rows="3"><%=replace(rs("boarduser")&"",",",chr(13)&chr(10))%></textarea><li>ÿ����<b>�س�</b>�ָ���
      </td>
    </tr>
    <%End If%>
    <tr> 
      <td height=22 class=tablebody1  align="center">ʹ������ģ��<br>
���ģ���а�����̳��ɫ��ͼƬ
������</td>
      <td  class=tablebody1> 
        <select name=sid>
<%	
	Dim rs_c
	set rs_c= server.CreateObject ("adodb.recordset")
	sql = "select * from dv_style"
	rs_c.open sql,conn,1,1
	if rs_c.eof and rs_c.bof then
	response.write "<option value=>�������ģ��"
	else
	do while not rs_c.EOF
%>
<option value=<%=rs_c("id")%> <% if cint(rs("sid")) = rs_c("id") then%> selected <%end if%>><%=rs_c("stylename")%> 
<%
	rs_c.MoveNext 
	loop
	end if
	rs_c.Close 
	Set rs_c=Nothing
%>
</select>
      </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody2>&nbsp;</td>
      <td  class=tablebody2> 
        <input type="submit" name="Submit" value="�ύ">
      </td>
    </tr>
  </table>
</form>
<%
rs.close
End Sub 
Sub savebminfo()
If Not IsObject(Conn) Then ConnectionDatabase
dim rname,i
dim readme,BoardType,boardmaster,sid,boarduser
readme=Dvbbs.checkStr(fixjs(Request.form("readme")))
BoardType=Dvbbs.checkStr(fixjs(Request.form("BoardType")))
boardmaster=Dvbbs.checkStr(fixjs(Request.form("boardmaster")))
If Cint(Dvbbs.Board_Setting(2))=1 Then
	boarduser=Dvbbs.checkStr(Request.form("boarduser"))
	boarduser=replace(boarduser,chr(13)&chr(10),",")
End If
sid=request("sid")
If IsNumeric(sid)=0 Or sid="" Then Response.redirect "showerr.asp?ErrCodes=<li>�Ƿ���ģ����&action=OtherErr"
If readme="" then Response.redirect "showerr.asp?ErrCodes=<li>��������̳��顣&action=OtherErr"
If BoardType="" then Response.redirect "showerr.asp?ErrCodes=<li>��������̳���ơ�&action=OtherErr"
If boardmaster="" then Response.redirect "showerr.asp?ErrCodes=<li>����������Ա��&action=OtherErr"
rname=split(boardmaster,"|")
For i=0 to ubound(rname)
	sql="select top 1 username from [dv_user] where username='"&replace(rname(i),"'","")&"'"
	set rs=Dvbbs.Execute(sql)
	If Rs.eof And rs.bof Then
	Response.redirect "showerr.asp?ErrCodes=<li>��̳û��"&replace(rname(i),"'","")&"����û����������Ϊ����&action=OtherErr"
	Exit For
	End If
	Set Rs=Nothing
Next

dim classname,titlepic
set rs=Dvbbs.Execute("select usertitle,GroupPic from dv_usergroups where usergroupid=3 order by Minarticle desc")
if not (rs.eof and rs.bof) then
classname=rs(0)
titlepic=rs(1)
end if
For i=0 to ubound(rname)
	sql="select top 1 UserGroupID from [dv_user] where username='"&replace(rname(i),"'","")&"'"
	Set Rs=Dvbbs.Execute(sql)
	If Rs(0)=4 Then Dvbbs.Execute("Update [dv_user] Set UserGroupID=3,userclass='"&classname&"',titlepic='"&titlepic&"' where username='"&replace(rname(i),"'","")&"'" )
	Set Rs=Nothing
Next

set rs=server.createobject("adodb.recordset")
sql = "select * from dv_board where boardid="+Cstr(request("boardid"))
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	Response.redirect "showerr.asp?ErrCodes=<li>��û��ָ����Ӧ��̳ID�����ܽ��й���&action=OtherErr"
End If
rs("boardmaster") = boardmaster
rs("readme") = readme
rs("BoardType")=BoardType
If Cint(Dvbbs.Board_Setting(2))=1 Then Rs("boarduser")=boarduser
Rs("sid")=Clng(sid)
rs.Update 
rs.Close 
response.write "<p>��̳�޸ĳɹ���"
Dvbbs.ReloadBoardInfo(Dvbbs.BoardID)
End Sub 
Function fixjs(Str)
	If Str <>"" Then
		str = replace(str,"\", "\\")
		Str = replace(str, chr(34), "\""")
		Str = replace(str, chr(39),"\'")
		Str = Replace(str, chr(13), "\n")
		Str = Replace(str, chr(10), "\r")
		str = replace(str,"'", "&#39;")
	End If
	fixjs=Str
End Function
Function enfixjs(Str)
	If Str <>"" Then
		Str = replace(str,"&#39;", "'")
		Str = replace(str,"\""" , chr(34))
		Str = replace(str, "\'",chr(39))
		Str = Replace(str, "\r", chr(10))
		Str = Replace(str, "\n", chr(13))
		Str = replace(str,"\\", "\")
	End If
	enfixjs=Str
End Function
%>

