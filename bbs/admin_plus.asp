<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
dim admin_flag
admin_flag=",36,"
If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	dvbbs_error()
Else
	Call main()
	Footer()
End if

Sub main()
	If  request("action")="save" Then
		call savenew()
	ElseIf request("action")="savedit" Then
		call savedit()
	ElseIf request("action")="del" Then
		call del()
	ElseIf request("action")="AddNew" OR request("action")="edit" Then
		AddNew()
	Else
		call gradeinfo()
	End If
End Sub
Sub AddNew()
dim trs
Dim PSetting
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><th width="100%" class="tableHeaderText" colspan=2>��̳����˵�����
</th>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow>ע�⣺������������ݽ��Զ���ʾ����̳ǰ̨�Ķ����˵�</td>
</tr>

<tr>
<td height="23" colspan="2" class="tableHeaderText">
<a href="admin_plus.asp">�˵�������ҳ</a>
<%if request("action")="edit" then
Set tRs=Dvbbs.Execute("Select * From Dv_Plus Where id="&request("id")&"")
PSetting=Split(Server.HTMLEncode(tRs("Plus_Setting")),"|||")
PSetting(0)=split(PSetting(0),"|")
%>
 | �༭�˵� | <a href="admin_plus.asp?action=AddNew">�½��˵�</a></td>
</tr>
<FORM METHOD=POST ACTION="?action=savedit">
<input type=hidden value="<%=trs("id")%>" name="id">
<tr>
<td height="23" colspan="2" class=forumrow>
���⣺ <input type=text size=50 name="title" value="<%=Server.HtmlEncode(tRs("Plus_name"))%>"> ����HTML�﷨
</td></tr>
<tr>
<td height="23" colspan="2" class=forumrow>
�Ƿ��ڵ�������ʾ�� ��  <Input type="radio" name="Isuse" value="1"
<%
If trs("Isuse")=1 Then 
%>
 checked
<%
End If 
%>
>  ��  <input type="radio" name="Isuse" value="0" 
<%
If trs("Isuse")=0 Then 
%>
 checked
<%
End If 
%>
>
</td></tr>
<tr>
<td height="23" colspan="2" class=forumrow>
���ࣺ
<Select Name="stype" size=1>
<%
Set Rs=Dvbbs.Execute("Select * From Dv_Plus Where Plus_type='0'")
If Rs.Eof And Rs.bof Then
	Response.Write "<option value=0>��Ϊһ���˵�</option>"
Else
	If Clng(tRs("Plus_type"))=0 Then
		Response.Write "<option value=0 selected>��Ϊһ���˵�</option>"
	Else
		Response.Write "<option value=0>��Ϊһ���˵�</option>"
	End If
	Do While Not Rs.Eof
		If CStr(request("id")) <> CStr(Rs("id")) Then 
			Response.Write "<option value="&rs("id")
			If Clng(tRs("Plus_type"))=Rs("ID") Then Response.Write " selected "
			Response.Write ">"&Server.htmlencode(rs("plus_name"))&"</option>"
		End If
	Rs.MoveNext
	Loop
End If
%>
</select>
��ѡ������Ϊһ���˵�<BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow>
ע�ͣ�
<input type=text size=50 name="readme" value="<%=server.htmlencode(trs("plus_copyright"))%>"> ��ʾ�������ϵ�titleע�ͣ�Ҳ�ǲ���İ�Ȩ��Ϣ<BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow>
ģʽ��
<select name="windowtype" size="1">
<option value="0" <%If PSetting(0)(0)=0 Then Response.Write "selected"%>>ԭ����</option>
<option value="1" <%If PSetting(0)(0)=1 Then Response.Write "selected"%>>�´���</option>
<option value="2" <%If PSetting(0)(0)=2 Then Response.Write "selected"%>>�̶���С����</option>
<option value="3" <%If PSetting(0)(0)=3 Then Response.Write "selected"%>>ȫ��</option>
</select>&nbsp;&nbsp;
���ڿ�<input type=text name="windowwidth" value="<%=PSetting(0)(1)%>" size=5>
&nbsp;&nbsp;
���ڸߣ�<input type=text name="windowheight" value="<%=PSetting(0)(2)%>" size=5>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow>
���ӣ�
<input type=text size=50 name="url" value="<%=server.htmlencode(trs("mainpage"))%>"> <BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow>
��̨�������ӣ�
<input type=text size=50 name="plus_adminpage" value="<%=server.htmlencode(trs("plus_adminpage")&"")%>"> <BR>
</td>
</tr>
<tr><th height="25" colspan="2" >�����������</th></tr>
<tr><td height="23" colspan="2" class=forumrow> ���ID&nbsp;&nbsp;<input type=text name="plusID" value="<%=Trs("plus_ID")%>" size="20">����������Ψһ�ı�ʶ��ע�ⲻ�����ظ��ġ�
<a href="http://bbs.dvbbs.net/Dv_plusInfo.asp" target="_blank" title="�������ٷ�վ��ѯ���ڲ������Ϣ" >��ò��ID����Ϣ</a> </td></tr>
<%
If UBound(PSetting)>2 Then
	PSetting(3)=Split(PSetting(3),",")
%>
<tr><td height="23" colspan="2" class=forumrow>�Ƿ�ʱ����&nbsp;&nbsp; ��  <Input type="radio" name="useTime" value="1"
<%
If PSetting(3)(0)="1" Then Response.Write "checked"
%>
>  ��  <input type="radio" name="useTime" value="0" 
<%
If PSetting(3)(0)="0" Then Response.Write "checked"
%>
></td></tr>
<tr><td height="23" colspan="2" class=forumrow>��ʱ������ֹʱ��&nbsp;&nbsp; <input type=text name="timesetting" value="<%=PSetting(3)(1)%>" size=10></td></tr>
<tr><td height="23" colspan="2" class=forumrow>��ʹ�ò�����û���&nbsp;&nbsp; <input type=text name="Groupsetting" value="<%=PSetting(3)(2)%>" size=30>�������ÿ���ʹ�ò�����û���</td></tr>
<tr><td height="23" colspan="2" class=forumrow>������Ա&nbsp;&nbsp; <textarea  name="plusmaster" cols="50" rows="5" ><%=Replace(PSetting(3)(3),"|",vbCrLf)%></textarea><br>�������ò���Ĺ���Ա��ÿ���û����ûس��ָ���.ϵͳĬ����̳����Ա���Թ����������������Ҫ�������ù���Ա��������Բ��</td></tr>
<tr><th height="25" colspan="2" >������</th></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò������������&nbsp;&nbsp;<input type=text name="Plus_UserPost" value=<%=PSetting(3)(4)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò������ͽ�Ǯ&nbsp;&nbsp;<input type=text name="Plus_userWealth" value=<%=PSetting(3)(5)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò������;���&nbsp;&nbsp;<input type=text name="Plus_UserEP" value=<%=PSetting(3)(6)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò�����������&nbsp;&nbsp;<input type=text name="Plus_UserCP" value=<%=PSetting(3)(7)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò�����������&nbsp;&nbsp;<input type=text name="Plus_UserPower" value=<%=PSetting(3)(8)%> size=5></td></tr>
<tr><th height="25" colspan="2" >������</th></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò����Ǯ�仯&nbsp;&nbsp;<input type=text name="Plus_ADDuserWealth" value=<%=PSetting(3)(9)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò������仯&nbsp;&nbsp;<input type=text name="Plus_ADDUserEP" value=<%=PSetting(3)(10)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò�������仯&nbsp;&nbsp;<input type=text name="Plus_ADDUserCP" value=<%=PSetting(3)(11)%> size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò�������仯&nbsp;&nbsp;<input type=text name="Plus_ADDUserPower" value=<%=PSetting(3)(12)%> size=5></td></tr>
<%
Else 
%>
<tr><td height="23" colspan="2" class=forumrow>�Ƿ�ʱ����&nbsp;&nbsp; ��  <Input type="radio" name="useTime" value="1">  ��  <input type="radio" name="useTime" value="0" checked></td></tr>
<tr><td height="23" colspan="2" class=forumrow>��ʱ������ֹʱ��&nbsp;&nbsp; <input type=text name="timesetting" value="0|24" size=10></td></tr>
<tr><td height="23" colspan="2" class=forumrow>��ʹ�ò�����û���&nbsp;&nbsp; <input type=text name="Groupsetting" value="1@2@3@4@5@6@7@8" size=30>�������ÿ���ʹ�ò�����û���</td></tr>
<tr><td height="23" colspan="2" class=forumrow>������Ա&nbsp;&nbsp; <textarea  name="plusmaster" cols="50" rows="5" ></textarea><br>�������ò���Ĺ���Ա��ÿ���û����ûس��ָ���.ϵͳĬ����̳����Ա���Թ����������������Ҫ�������ù���Ա��������Բ��</td></tr>
<tr><th height="25" colspan="2" >������</th></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò������������&nbsp;&nbsp;<input type=text name="Plus_UserPost" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò������ͽ�Ǯ&nbsp;&nbsp;<input type=text name="Plus_userWealth" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò������;���&nbsp;&nbsp;<input type=text name="Plus_UserEP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò�����������&nbsp;&nbsp;<input type=text name="Plus_UserCP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò�����������&nbsp;&nbsp;<input type=text name="Plus_UserPower" value=0 size=5></td></tr>
<tr><th height="25" colspan="2" >������</th></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò����Ǯ�仯&nbsp;&nbsp;<input type=text name="Plus_ADDuserWealth" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò������仯&nbsp;&nbsp;<input type=text name="Plus_ADDUserEP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò�������仯&nbsp;&nbsp;<input type=text name="Plus_ADDUserCP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò�������仯&nbsp;&nbsp;<input type=text name="Plus_ADDUserPower" value=0 size=5></td></tr>
<%
End If
%>
<tr><td height="23" colspan="2" class=forumrow>ע�⣬����û����еĿ�����(IDΪ7)����Ϊ�ɽ��룬��ô���е���������Ч��<Br>������������ǿ��˿�ʹ�ã��Կ�����Ч</td></tr>
<tr><th height="25" colspan="2" >����Զ�Ȩ������</th></tr>
<tr><td height="23" colspan="2" class=forumrow>
<textarea name="Plus_Setting" cols="80" rows="20">
<%	
	If UBound(PSetting)>1 Then 
		Dim i
		PSetting(1)=Split(PSetting(1),",")
		PSetting(2)=Split(PSetting(2),",")
		For i=0 to UBound (PSetting(1))
			Response.Write PSetting(2)(i)&"="&PSetting(1)(i)
			Response.Write vbCrLf
		Next
	Else
%>
�����ֶ�1=0
�����ֶ�2=0
�����ֶ�3=0
�����ֶ�4=0
�����ֶ�5=0
�����ֶ�6=0
�����ֶ�7=0
�����ֶ�8=0
�����ֶ�9=0
�����ֶ�10=0
�����ֶ�11=0
�����ֶ�12=0
�����ֶ�13=0
�����ֶ�14=0
�����ֶ�15=0
�����ֶ�16=0
�����ֶ�17=0
�����ֶ�18=0
�����ֶ�19=0

<%
End If
%>
</textarea>
</td></tr>
<tr><td height="25" colspan="2" >˵��������ÿ����������ò�������ȫһ���������ֶεĶ���Ҳ��һ������Щ������������������޸��ˡ�</td></tr>
<tr>
<td height="23" colspan="2" class=forumrowHighLight>
<input type=submit name=submit value="�ύ">
</td></tr>
</FORM>
<%else%>
<tr>
<th height="23" colspan="2">��Ӳ˵�</th>
</tr>
<FORM METHOD=POST ACTION="?action=save">
<tr>
<td height="23" colspan="2" class=forumrow>
���⣺ <input type=text size=50 name="title"> ����HTML�﷨
</td></tr>
<tr>
<td height="23" colspan="2" class=forumrow>
�Ƿ��ڵ�������ʾ�� ��  <Input type="radio" name="Isuse" value="1" checked>  ��  <input type="radio" name="Isuse" value="0" >
</td></tr>
<tr>
<td height="23" colspan="2" class=forumrow>
���ࣺ
<Select Name="stype" size=1>
<%
Set Rs=Dvbbs.Execute("Select * From Dv_Plus Where Plus_type='0'")
If Rs.Eof And Rs.bof Then
	Response.Write "<option value=0>��Ϊһ���˵�</option>"
Else
	Response.Write "<option value=0>��Ϊһ���˵�</option>"
	Do While Not Rs.Eof
		Response.Write "<option value="&rs("id")&">"&Server.htmlencode(rs("plus_name"))&"</option>"
	Rs.MoveNext
	Loop
End If
%>
</select>
��ѡ������Ϊһ���˵�<BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow>
ע�ͣ�
<input type=text size=50 name="readme"> ��ʾ�������ϵ�titleע��,Ҳ�ǲ���İ�Ȩ��Ϣ<BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow>
ģʽ��
<select name="windowtype" size="1">
<option value="0">ԭ����</option>
<option value="1">�´���</option>
<option value="2">�̶���С����</option>
<option value="3">ȫ��</option>
</select>&nbsp;&nbsp;
���ڿ�<input type=text name="windowwidth" value=0 size=5>
&nbsp;&nbsp;
���ڸߣ�<input type=text name="windowheight" value=0 size=5>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow>
���ӣ�
<input type=text size=50 name="url"> <BR>
</td>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow>
��̨�������ӣ�
<input type=text size=50 name="plus_adminpage"> <BR>
</td>
</tr>
<tr><th height="25" colspan="2" >�����������</th></tr>
<tr><td height="23" colspan="2" class=forumrow> ���ID&nbsp;&nbsp;<input type=text name="plusID" value="newplus1" size="20">����������Ψһ�ı�ʶ��ע�ⲻ�����ظ��ġ�
<a href="http://bbs.dvbbs.net/Dv_plusInfo.asp" target="_blank" title="�������ٷ�վ��ѯ���ڲ������Ϣ" >��ò��ID����Ϣ</a> </td></tr>
<tr><td height="23" colspan="2" class=forumrow>�Ƿ�ʱ����&nbsp;&nbsp; ��  <Input type="radio" name="useTime" value="1">  ��  <input type="radio" name="useTime" value="0" checked></td></tr>
<tr><td height="23" colspan="2" class=forumrow>��ʱ������ֹʱ��&nbsp;&nbsp; <input type=text name="timesetting" value="0|24" size=10></td></tr>
<tr><td height="23" colspan="2" class=forumrow>��ʹ�ò�����û���&nbsp;&nbsp; <input type=text name="Groupsetting" value="1@2@3@4@5@6@7@8" size=30>�������ÿ���ʹ�ò�����û���</td></tr>
<tr><td height="23" colspan="2" class=forumrow>������Ա&nbsp;&nbsp; <textarea  name="plusmaster" cols="50" rows="5" ></textarea><br>�������ò���Ĺ���Ա��ÿ���û����ûس��ָ���.ϵͳĬ����̳����Ա���Թ����������������Ҫ�������ù���Ա��������Բ��</td></tr>
<tr><th height="25" colspan="2" >������</th></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò������������&nbsp;&nbsp;<input type=text name="Plus_UserPost" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò������ͽ�Ǯ&nbsp;&nbsp;<input type=text name="Plus_userWealth" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò������;���&nbsp;&nbsp;<input type=text name="Plus_UserEP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò�����������&nbsp;&nbsp;<input type=text name="Plus_UserCP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ��ʹ�ò�����������&nbsp;&nbsp;<input type=text name="Plus_UserPower" value=0 size=5></td></tr>
<tr><th height="25" colspan="2" >������</th></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò����Ǯ�仯&nbsp;&nbsp;<input type=text name="Plus_ADDuserWealth" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò������仯&nbsp;&nbsp;<input type=text name="Plus_ADDUserEP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò�������仯&nbsp;&nbsp;<input type=text name="Plus_ADDUserCP" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow> ÿ��ʹ�ò�������仯&nbsp;&nbsp;<input type=text name="Plus_ADDUserPower" value=0 size=5></td></tr>
<tr><td height="23" colspan="2" class=forumrow>ע�⣬����û����еĿ�����(IDΪ7)����Ϊ�ɽ��룬��ô���е���������Ч��<Br>������������ǿ��˿�ʹ�ã��Կ�����Ч</td></tr>
<tr><th height="25" colspan="2" >����Զ�����չ����</th></tr>
<tr><td height="23" colspan="2" class=forumrow>
<textarea name="Plus_Setting" cols="80" rows="20">
�����ֶ�1=0
�����ֶ�2=0
�����ֶ�3=0
�����ֶ�4=0
�����ֶ�5=0
�����ֶ�6=0
�����ֶ�7=0
�����ֶ�8=0
�����ֶ�9=0
�����ֶ�10=0
�����ֶ�11=0
�����ֶ�12=0
�����ֶ�13=0
�����ֶ�14=0
�����ֶ�15=0
�����ֶ�16=0
�����ֶ�17=0
�����ֶ�18=0
�����ֶ�19=0

</textarea>
</td></tr>
<tr><td height="25" colspan="2" >˵��������ÿ����������ò�������ȫһ���������ֶεĶ���Ҳ��һ������Щ������������������޸��ˡ�</td></tr>
<tr>
<td height="23" colspan="2" class=forumrowHighLight>
<input type=submit name=submit value="�ύ">
</td></tr>
</FORM>
<%end if%>
</table><BR>
<%
End Sub 
sub gradeinfo()
dim trs
Dim PSetting
%>

<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><th width="100%" class="tableHeaderText" colspan=5>��̳����˵�����
</th>
</tr>
<tr>
<td height="23" colspan="5" class=forumrow>ע�⣺������������ݽ��Զ���ʾ����̳ǰ̨�Ķ����˵�</td>
</tr>
<tr>
<td height="23" colspan="5" class=forumrow><a href="admin_plus.asp?action=AddNew">�½��˵�</a> | <a href="admin_plus.asp?action=posttodb">�������ģ������</a> | <a href="admin_plus.asp?action=getfromdb">������ģ������</a></td>
</tr>
<tr>
<th height="23">����</th>
<th>����</th>
<th>��������</th>
<th>�Ƿ���ʾ</th>
<th>����</th>
</tr>
<%
Set Rs=Dvbbs.Execute("Select * From Dv_Plus Order by ID Desc")
Do While Not Rs.Eof
PSetting=Split(Rs("Plus_Setting"),"|")
%>
<tr>
<td height="23" class=forumrow><%=Rs("Plus_Name")%></td>
<td class=forumrow>
<%If Rs("Plus_type")=0 Then%>
һ���˵�
<%Else%>
<%
Set tRs=Dvbbs.Execute("Select * From Dv_Plus Where id="&Rs("Plus_Type")&"")
If tRs.Eof And tRs.Bof Then
	Response.Write "�ò˵�����������༭����"
Else
	Response.Write tRs("Plus_name")
End If
%>
<%End If%>
</td>
<td class=forumrow>
<%
Select Case PSetting(0)
Case 0
	Response.Write "��ǰ����"
Case 1
	Response.Write "�´���"
Case 2
	Response.Write "�̶���С���ڣ���"&PSetting(1)&"���� "&PSetting(2)&""
Case 3
	Response.Write "ȫ��"
End Select
%>
</td>
<td class=forumrow align=center><%If Rs("isuse")=1 Then%>Yes<%Else%>No<%End If%></td>

<td class=forumrow align=center><%
If Rs("plus_adminpage") <> "" Then  
%>
<a href="<%=Rs("plus_adminpage")%>">����</a> | 
<%
End If
%>
<a href="?action=edit&id=<%=rs("id")%>">�༭</a> | <a href="?action=del&id=<%=rs("id")%>">ɾ��</a></td>
</tr>
<%
Rs.MoveNext
Loop
Set Rs=Nothing
end sub

sub savenew()
	Dim plusID,plus_adminpage,Isuse
	plusID=Trim(Request("plusID"))
	If InStr(plusID,"'") >0 Then 
		Response.Write "���ID�в������е�����"
		exit sub
	End If
	If request("title")="" then
		Response.Write "������˵��ı��⣡"
		exit sub
	End If
	If plusID="" Then
		Response.Write "���趨���ID"
		exit sub
	End If
	SQL="Select count(*) From Dv_plus where plus_ID='"&plusID&"'"
	Set rs=Dvbbs.execute(SQL)
	If Rs(0) >0 Then
		Response.Write "�����õĲ��ID�Ѿ����ڣ����������á�"
		exit sub
	End If
	Isuse=Request("Isuse")
	If Isuse<>"1" And Isuse<>"0" Then Isuse="1"
	Isuse=CInt(Isuse)
	plus_adminpage=Dvbbs.checkStr(Request("plus_adminpage"))
	Dim Plus_SettingData,Plus_Setting,i,tmpstr
	Plus_Setting=Request("Plus_Setting")
	Plus_Setting=Split(Plus_Setting,vbCrLf)
	Plus_SettingData=""
	For i=0 to UBound(Plus_Setting)
		Plus_Setting(i)=Split(Plus_Setting(i),"=")
		If UBound(Plus_Setting(i))=1 Then
			If Plus_SettingData="" Then 
				Plus_SettingData=Trim(Plus_Setting(i)(1))
				tmpstr=Trim(Plus_Setting(i)(0))
			Else
				Plus_SettingData=Plus_SettingData&","&Trim(Plus_Setting(i)(1))
				tmpstr=tmpstr&","&Trim(Plus_Setting(i)(0))
			End If
		End If
	Next
	Plus_SettingData=Plus_SettingData&"|||"&tmpstr&"|||"
	Dim plusmaster,masterlist
	plusmaster=Request("plusmaster")
	plusmaster=split(plusmaster,vbCrLf)
	masterlist=""
	For i=0 to UBound(plusmaster)
		If Trim(plusmaster(i)) <>"" Then
			If masterlist="" Then
				masterlist=plusmaster(i)
			Else
				masterlist=masterlist&"|"&plusmaster(i)
			End If
		End If
	Next
	Dim useTime,timesetting,Groupsetting,Plus_UserPost,Plus_userWealth,Plus_UserEP,Plus_UserCP
	Dim Plus_UserPower,Plus_ADDuserWealth,Plus_ADDUserEP,Plus_ADDUserCP,Plus_ADDUserPower,guestuse
	useTime=Request("useTime")
	If useTime="" Then useTime=0
	timesetting=Trim(Request("timesetting"))
	If timesetting="" Then timesetting="0|24"
	Groupsetting=Trim(Request("Groupsetting"))
	Plus_UserPost=Trim(Request("Plus_UserPost"))
	If Plus_UserPost="" Then Plus_UserPost=0
	Plus_userWealth=Trim(Request("Plus_userWealth"))
	If Plus_userWealth="" Then Plus_userWealth=0
	Plus_UserEP=Trim(Request("Plus_UserEP"))
	If Plus_UserEP="" Then Plus_UserEP=0
	Plus_UserCP=Trim(Request("Plus_UserCP"))
	If Plus_UserCP="" Then Plus_UserCP=0
	Plus_UserPower=Trim(Request("Plus_UserPower"))
	If Plus_UserPower="" Then Plus_UserPower=0
	Plus_ADDuserWealth=Trim(Request("Plus_ADDuserWealth"))
	If Plus_ADDuserWealth="" Then Plus_ADDuserWealth=0
	Plus_ADDUserEP=Trim(Request("Plus_ADDUserEP"))
	If Plus_ADDUserEP="" Then Plus_ADDUserEP=0
	Plus_ADDUserCP=Trim(Request("Plus_ADDUserCP"))
	If Plus_ADDUserCP="" Then Plus_ADDUserCP=0
	Plus_ADDUserPower=Trim(Request("Plus_ADDUserPower"))
	If Plus_ADDUserPower="" Then Plus_ADDUserPower=0
	guestuse=Request("guestuse")
	tmpstr=useTime&","&timesetting&","&Groupsetting&","&masterlist&","&plus_UserPost&","
	tmpstr=tmpstr&Plus_userWealth&","&Plus_UserEP&","&Plus_UserCP&","&Plus_UserPower&","
	tmpstr=tmpstr&Plus_ADDuserWealth&","&Plus_ADDUserEP&","&Plus_ADDUserCP&","&Plus_ADDUserPower&","&guestuse
	Plus_SettingData=Plus_SettingData&tmpstr
	set rs=server.createobject("adodb.recordset")
	sql="select * from dv_plus"
	rs.open sql,conn,1,3
	rs.addnew
	Rs("plus_ID")=plusID
	rs("plus_type")=request("stype")
	rs("plus_name")=replace(request("title"),CHR(34),"")
	rs("isuse")=Isuse
	rs("IsShowMenu")=1
	rs("Mainpage")=replace(request("url"),CHR(34),"")
	rs("plus_Copyright")=replace(request("readme"),CHR(34),"")
	rs("Plus_Setting")=request("windowtype") & "|" & request("windowwidth") & "|" & request("windowheight")&"|||"&Plus_SettingData
	Rs("plus_adminpage")=plus_adminpage
	rs.update
	rs.close
	set rs=nothing
	dv_suc("�½���̳�˵��ɹ�")
	DelForumPlusMenuCache
End sub
sub savedit()
	Dim plusID,plus_adminpage,Isuse
	plusID=Trim(Request("plusID"))
	If InStr(plusID,"'") >0 Then 
		Response.Write "���ID�в������е�����"
		exit sub
	End If
	If request("title")="" then
		Response.Write "������˵��ı��⣡"
		exit sub
	End If
	If plusID="" Then
		Response.Write "���趨���ID"
		exit sub
	End If
	SQL="Select count(*) From Dv_plus where plus_ID='"&plusID&"' and id<>"&request("id")&""
	Set rs=Dvbbs.execute(SQL)
	If Rs(0) >0 Then
		Response.Write "�����õĲ��ID�Ѿ����ڣ����������á�"
		exit sub
	End If
	Isuse=Request("Isuse")
	If Isuse<>"1" And Isuse<>"0" Then Isuse="1"
	Isuse=CInt(Isuse)
	plus_adminpage=Dvbbs.Checkstr(Request("plus_adminpage"))
	Dim Plus_SettingData,Plus_Setting,i,tmpstr
	Plus_Setting=Request("Plus_Setting")
	Plus_Setting=Split(Plus_Setting,vbCrLf)
	Plus_SettingData=""
	For i=0 to UBound(Plus_Setting)
		Plus_Setting(i)=Split(Plus_Setting(i),"=")
		If UBound(Plus_Setting(i))=1 Then
			If Plus_SettingData="" Then 
				Plus_SettingData=Trim(Plus_Setting(i)(1))
				tmpstr=Trim(Plus_Setting(i)(0))
			Else
				Plus_SettingData=Plus_SettingData&","&Trim(Plus_Setting(i)(1))
				tmpstr=tmpstr&","&Trim(Plus_Setting(i)(0))
			End If
		End If
	Next
	Plus_SettingData=Plus_SettingData&"|||"&tmpstr&"|||"
	Dim plusmaster,masterlist
	plusmaster=Request("plusmaster")
	plusmaster=split(plusmaster,vbCrLf)
	masterlist=""
	For i=0 to UBound(plusmaster)
		If Trim(plusmaster(i)) <>"" Then
			If masterlist="" Then
				masterlist=plusmaster(i)
			Else
				masterlist=masterlist&"|"&plusmaster(i)
			End If
		End If
	Next
	Dim useTime,timesetting,Groupsetting,Plus_UserPost,Plus_userWealth,Plus_UserEP,Plus_UserCP
	Dim Plus_UserPower,Plus_ADDuserWealth,Plus_ADDUserEP,Plus_ADDUserCP,Plus_ADDUserPower,guestuse
	useTime=Request("useTime")
	If useTime="" Then useTime=0
	timesetting=Trim(Request("timesetting"))
	If timesetting="" Then timesetting="0|24"
	Groupsetting=Trim(Request("Groupsetting"))
	Plus_UserPost=Trim(Request("Plus_UserPost"))
	If Plus_UserPost="" Then Plus_UserPost=0
	Plus_userWealth=Trim(Request("Plus_userWealth"))
	If Plus_userWealth="" Then Plus_userWealth=0
	Plus_UserEP=Trim(Request("Plus_UserEP"))
	If Plus_UserEP="" Then Plus_UserEP=0
	Plus_UserCP=Trim(Request("Plus_UserCP"))
	If Plus_UserCP="" Then Plus_UserCP=0
	Plus_UserPower=Trim(Request("Plus_UserPower"))
	If Plus_UserPower="" Then Plus_UserPower=0
	Plus_ADDuserWealth=Trim(Request("Plus_ADDuserWealth"))
	If Plus_ADDuserWealth="" Then Plus_ADDuserWealth=0
	Plus_ADDUserEP=Trim(Request("Plus_ADDUserEP"))
	If Plus_ADDUserEP="" Then Plus_ADDUserEP=0
	Plus_ADDUserCP=Trim(Request("Plus_ADDUserCP"))
	If Plus_ADDUserCP="" Then Plus_ADDUserCP=0
	Plus_ADDUserPower=Trim(Request("Plus_ADDUserPower"))
	If Plus_ADDUserPower="" Then Plus_ADDUserPower=0
	guestuse=Request("guestuse")
	tmpstr=useTime&","&timesetting&","&Groupsetting&","&masterlist&","&plus_UserPost&","
	tmpstr=tmpstr&Plus_userWealth&","&Plus_UserEP&","&Plus_UserCP&","&Plus_UserPower&","
	tmpstr=tmpstr&Plus_ADDuserWealth&","&Plus_ADDUserEP&","&Plus_ADDUserCP&","&Plus_ADDUserPower&","&guestuse
	Plus_SettingData=Plus_SettingData&tmpstr
	set rs=server.createobject("adodb.recordset")
	sql="select * from dv_plus where id="&request("id")&""
	rs.open sql,conn,1,3
	Rs("plus_ID")=plusID
	rs("plus_type")=request("stype")
	rs("plus_name")=replace(request("title"),CHR(34),"")
	rs("isuse")=Isuse
	rs("IsShowMenu")=1
	rs("Mainpage")=replace(request("url"),CHR(34),"")
	rs("plus_Copyright")=replace(request("readme"),CHR(34),"")
	rs("Plus_Setting")=request("windowtype") & "|" & request("windowwidth") & "|" & request("windowheight")&"|||"&Plus_SettingData
	Rs("plus_adminpage")=plus_adminpage
	rs.update
	rs.close
	set rs=nothing
	dv_suc("�޸���̳�˵��ɹ�")
	DelForumPlusMenuCache
end sub


sub del()
	Dvbbs.Execute("Delete From Dv_Plus Where ID="&Request("id")&"")
	dv_suc("ɾ����̳�˵��ɹ�")
	DelForumPlusMenuCache
end Sub
Sub DelForumPlusMenuCache()
	Dim Templateslist
	Dvbbs.Name="Templateslist"
	If Dvbbs.ObjIsEmpty() Then  Dvbbs.ReloadTemplateslist()
	Templateslist= Dvbbs.Value
	Templateslist=split(Templateslist,"@@@")
	Dim i
	For i=0 to UBound(Templateslist)
		Templateslist(i)=split(Templateslist(i),"|||")
		Dvbbs.DelCahe("ForumPlusMenu"&Templateslist(i)(0))
	Next
	Dvbbs.DelCahe("Plus_Settingts")
End Sub
Sub FixPlusTable()
	Dim Rs,SQL
	SQL="select * From Dv_plus"
	Set Rs=Dvbbs.Execute(SQL)
	If Rs.Fields.Count < 10 Then
		Set Rs=Nothing
		Dvbbs.Execute("alter table [Dv_plus] add plus_adminpage varchar(100)")
		Dvbbs.Execute("alter table [Dv_plus] add plus_id varchar(100)")
		Set Rs=Dvbbs.Execute(SQL)
		If Not Rs.Eof Then
			Do While Not Rs.Eof
				Dvbbs.execute("update [Dv_plus] set plus_id='newplus"&Rs(0)&"' Where ID="&Rs(0)&"")
				Rs.MoveNext
			Loop
		End If
	End If
	Set Rs=Nothing
End Sub
%>