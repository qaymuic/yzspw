<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!-- #include file="inc/DvADChar.asp" -->
<%
Head()
Dim admin_flag
admin_flag=",20,"


Dim StyleID,StyleName,Style_Pic,Stype
Dim Forum_emotNum,Forum_userfaceNum,Forum_PostFaceNum
Dim Forum_PostFace,Forum_userface,Forum_emot
Dim face_id,Count
Dim newnum,newfilename
Dim bbspicmun,bbspicurl,picfilename,actname,connfile,upconfig
Dim TempForum_PostFace,TempForum_userface,TempForum_emot

If IsNumeric(Request("Stype")) and Request("Stype")<>"" Then
	Stype = Cint(Request("Stype"))		'1=���飬2=����em��3=ͷ��
Else
	Stype=4
End If

If Request.QueryString("StyleID")<>"" and IsNumeric(Request.QueryString("StyleID")) Then
	StyleID=Cint(Request("StyleID"))
Else
	StyleID=Dvbbs.cachedata(17,0)
End If

If StyleID="" Then StyleID=1

If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then
	Founderr=true
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	dvbbs_error()
Else
	GetNum()
End If

If Founderr=false Then
	Select Case Stype
	case 1
	'skins/default/topicface/face1.gif
		bbspicmun=Forum_PostFaceNum-1
		If not isarray(Forum_PostFace) Then
			bbspicurl="Skins/default/topicface/"
		Else
			bbspicurl=Forum_PostFace(0)
		End If
		connfile=Forum_PostFace
		actname="��������ͼƬ"
		picfilename="face"
	case 2
	'Skins/Default/emot/em01.gif	'Forum_emot
        bbspicmun=Forum_emotNum-1
		If not isarray(Forum_emot) Then
			bbspicurl="Skins/Default/emot/"
		Else
			bbspicurl=Forum_emot(0)
		End If
		connfile=Forum_emot
		actname="��������ͼƬ"
		picfilename="em"
	case 3
	'Images/userface/image1.gif
		bbspicmun=Forum_userfaceNum-1
		If not isarray(Forum_userface) Then
			bbspicurl="Images/userface/"
		Else
			bbspicurl=Forum_userface(0)
		End If
		connfile=Forum_userface
		actname="ע��ͷ��"
		picfilename="image"
	case else
	'Images/userface/image1.gif
		bbspicmun=Forum_userfaceNum-1
		If not isarray(Forum_userface) Then
			bbspicurl="Images/userface/"
		Else
			bbspicurl=Forum_userface(0)
		End If
		connfile=Forum_userface
		actname=""
		picfilename="image"
	End Select

	if trim(Request("newfilename"))<>"" then
		newfilename=trim(request("newfilename"))
	else
		newfilename=picfilename
	end if 

	if bbspicmun<0 then 
	count=1
	else
	count=bbspicmun+1
	end if

	if REQUEST("Newnum")<>"" and request("Newnum")<>0 then
		newnum=REQUEST("Newnum")
	else
		newnum=0
	end if

	if request("Submit")="��������" then
		call saveconst()
	elseif request("Submit")="�ָ�Ĭ������" then
		call savedefault()
	ElseIf request("Submit")="�ָ�Ĭ��������" then
		Stype=4
		call savedefault()
	else
		call consted()
	end if
End If
if Founderr then dvbbs_error()
Footer()
sub consted()
dim sel
%>
<form method="POST"  action=?Stype=<%=request("Stype")%>  name="bbspic" >
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
<tr> 
<td height="23" colspan="4" ><B>˵��</B>��<br>�١�����ͼƬ����������̳<%=bbspicurl%>Ŀ¼�У���Ҫ����Ҳ�뽫ͼƬ���ڸ�Ŀ¼<br>�ڡ��ұ߸�ѡ��Ϊɾ��ѡ����ѡ���㱣�����ã���ɾ����ӦͼƬ<BR>�ۡ�������޸��ļ����������޸���Ӧѡ���ֱ�ӵ���������ö�����ѡ���ұ߸�ѡ��
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="4" align=left><%=actname%>�������� ��Ŀǰ����<%=count%>��<%=actname%>ͼƬ���ļ��У�<%=bbspicurl%>��</th>
</tr>
<tr> 
<td width="20%"  align=left class=forumrow>��ǰģ�����ƣ�</td>
<td width="80%"  align=left class=forumrow colspan="3"><%=StyleName%>
</td>
</tr>
<tr> 
<td width="20%"  align=left class=forumrow>���ӵ��ļ�����</td>
<td width="80%"  align=left class=forumrow colspan="3"><input  type="text" name="NEWFILENAME" value="<%=newfilename%>">��<font color=red>�������Ĭ�ϣ����Ӻ����Ӧ���ļ����ϴ�����Ŀ¼�¡�</font>��
</td>
</tr>
<tr> 
<td width="20%"  align=left class=forumrow>����������Ŀ��</td>
<td width="80%"  align=left class=forumrow colspan="3"><input  type="text" name="NEWNUM" value="<%=newnum%>">
<input type="submit" name="Submit" value="����">
</td>
</tr>
<tr> 
<td width="20%"  align=left class=forumrow>��������ģ�壺</td>
<td width="80%"  align=left class=forumrow colspan="3">��<input type=radio name=coverall value=1 >��<input type=radio name=coverall value=0 checked>
</td>
</tr>
<%
Dim TempName 
IF REQUEST("Submit")="����" and REQUEST("Newnum")<>"" and request("Newnum")<>0 then
newnum=REQUEST("Newnum")

for i=count to count+newnum-1
if stype=2 and i<10 Then
TempName = newfilename&"0"&i
Else
TempName = newfilename&i
End If
%>
<tr>
<td width="20%" class=forumRowHighlight><%=actname%>ID��<input type=hidden name="face_id<%=i%>" size="10" value="<%=i%>"><%=i%></td>
<td width="75%" class=forumRowHighlight colspan="2">�����ӵ��ļ���<input  type="text" name="userface<%=i%>" value="<%=TempName%>.gif"></td>
<td width="5%" class=forumRowHighlight> 
<input type="checkbox" name="delid<%=i%>" value="<%=i%>">
</td>
</tr>
<% next 
end if
%>
<tr>
<th width="20%" class=forumrow>�ļ�</th>
<th width="45%" class=forumrow>�ļ���</th>
<th width="30%" class=forumrow> 
ͼƬ
<th width="5%" class=forumrow> 
ɾ��
</th>
</tr>
<tr>
<td width="20%" class=forumrow>�ļ�Ŀ¼��<input type=hidden  name="face_id0" size="10" ></td>
<td width="45%" class=forumrow>&nbsp;<input  type="text" name="userface0" value="<%=bbspicurl%>"></td>
<td width="30%" class=forumrow></td>
<td width="5%" class=forumrow></td>
</tr>
<% for i=1 to bbspicmun %>
<tr>
<td width="20%" class=forumrow>�ļ�����<input type=hidden  name="face_id<%=i%>" size="10" value="<%=i%>"></td>
<td width="45%" class=forumrow>&nbsp;<input  type="text" name="userface<%=i%>" value="<%=connfile(i)%>"></td>
<td width="30%" class=forumrow> 
&nbsp;&nbsp;<img src=<%=bbspicurl%><%=connfile(i)%>>
<td width="5%" class=forumrow> 
<input type="checkbox" name="delid<%=i%>" value="<%=i+1%>">
</td>
</tr>
<% next %>
<tr> 
<td  colspan="4" class=forumrow> 
<B>ע��</B>���ұ߸�ѡ��Ϊɾ��ѡ����ѡ���㱣�����ã���ɾ����ӦͼƬ<BR>������޸��ļ����������޸���Ӧѡ���ֱ�ӵ���������ö�����ѡ���ұ߸�ѡ��
</td>
</tr>
<tr> 
<td  colspan="4" class=forumrow> 
<div align="center"> 
 ɾ��ѡ�ɾ����ѡ��ʵ���ļ���<font color=red>��ҪFSO֧�ֹ���</font>������<input type=radio name=setfso value=1 >��<input type=radio name=setfso value=0 checked> ��ѡ��Ҫɾ�����ļ���<input type="checkbox" name=chkall value=on onclick="CheckAll(this.form)">ȫѡ <BR>
<input type="submit" name="Submit" value="��������">
<input type="submit" name="Submit" value="�ָ�Ĭ������">
<input type="submit" name="Submit" value="�ָ�Ĭ��������">
</div>
</td>
</tr>
</table><BR><BR>
</form>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>
<%
end sub

sub saveconst()
dim f_userface,formname,d_elid,faceid
dim filepaths,objFSO,upface
For i=0 to count+newnum-1
	faceid="face_id"&i
	d_elid="delid"&i
	formname="userface"&i
	If CInt(request.Form(d_elid))=0 Then
		f_userface=f_userface&request.Form(formname)&"|||"
	Else
		upface=bbspicurl&Request.Form(formname)
		upface=replace(upface,"..","")
		upface=replace(upface,"\","")
		If  request("setfso")=1 Then
			filepaths=Server.MapPath(""&upface&"")
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			If objFSO.fileExists(filepaths) Then
				'objFSO.DeleteFile(filepaths)
				response.write "ɾ��"&filepaths
			Else
				response.write "δ�ҵ�"&filepaths
			End If
		End If
	End If
Next
Set objFSO=Nothing
''1=���飬2=����em��3=ͷ��
'Style_Pic=TempForum_userface+"@@@"+TempForum_PostFace+"@@@"+TempForum_emot
	f_userface=replace(f_userface,"@@@","")
	Select Case Stype
	Case 1
		upconfig=TempForum_userface+"@@@"+f_userface+"@@@"+TempForum_emot
	Case 2
		upconfig=TempForum_userface+"@@@"+TempForum_PostFace+"@@@"+f_userface
	Case 3
		upconfig=f_userface+"@@@"+TempForum_PostFace+"@@@"+TempForum_emot
    End Select
	upconfig=Dvbbs.checkstr(upconfig)


	if Request.form("coverall")=1 Then
		sql = "update Dv_Style  set Style_Pic='"&upconfig&"'"
	Else
		sql = "update Dv_Style  set Style_Pic='"&upconfig&"' where id="&styleId
	End If
	Dvbbs.execute(sql)
	Dvbbs.DelCahe "Style_Pic"&StyleID
	Dv_suc(actname&"���óɹ���")
End Sub

sub savedefault()
dim userface,upconfig
userface=""
select case Stype
case 1     
        for i=1 to 18
        userface=userface&"face"&i&".gif|||"
        next
		userface="Skins/default/topicface/|||"+userface
		upconfig=TempForum_userface+"@@@"+userface+"@@@"+TempForum_emot
case 2
        for i=1 to 9
        userface=userface&"em0"&i&".gif|||"
        next
        for i=10 to 49
        userface=userface&"em"&i&".gif|||"
        next
		userface="Skins/Default/emot/|||"+userface
		upconfig=TempForum_userface+"@@@"+TempForum_PostFace+"@@@"+userface
case 3
        for i=1 to 60
        userface=userface&"image"&i&".gif|||"
        next
		userface="Images/userface/|||"+userface
		upconfig=userface+"@@@"+TempForum_PostFace+"@@@"+TempForum_emot
case else
		''ͷ��---------------------------------------
        for i=1 to 60
        userface=userface&"image"&i&".gif|||"
        next
		userface="Images/userface/|||"+userface
		upconfig=userface+"@@@"
		''����---------------------------------------
		userface=""
        for i=1 to 18
        userface=userface&"face"&i&".gif|||"
        next
		userface="Skins/default/topicface/|||"+userface
		upconfig=upconfig+userface+"@@@"
		''����---------------------------------------
		userface=""
        for i=1 to 9
        userface=userface&"em0"&i&".gif|||"
        next
        for i=10 to 49
        userface=userface&"em"&i&".gif|||"
        next
		userface="Skins/Default/emot/|||"+userface
		upconfig=upconfig+userface
end select
upconfig=Dvbbs.checkstr(upconfig)
if Request.form("coverall")=1 Then
sql = "update Dv_Style  set Style_Pic='"&upconfig&"'"
Else
sql = "update Dv_Style  set Style_Pic='"&upconfig&"' where id="&styleId
End If
Dvbbs.DelCahe "Style_Pic"&StyleID
Dvbbs.execute(sql)
Dv_suc(actname&"�ָ����óɹ���")
end sub

'������Dv_Style
'�ֶ�����Style_Pic
'@@@,|||

Sub GetNum()
	Dim NRs
	SQL=" Select id,StyleName,Style_Pic from Dv_Style where id="&styleId
	Set NRs=Dvbbs.Execute (SQL)
	If not NRs.eof Then
		StyleId=NRs(0)
		StyleName=NRs(1)
		Style_Pic=NRs(2)
	Else
		Errmsg=ErrMsg + "<li>"+"ģ��δ�ҵ��������ѱ�ɾ����������ѡȡ��ȷģ�棡"
		Founderr=True
		Exit Sub
	End if
	Rs.close:Set Rs=Nothing
	Style_Pic=Split(Style_Pic,"@@@")	'ģ�������@@@�ָС����|||�ָ�;
	TempForum_userface=Style_Pic(0)			'�û�ͷ��
	TempForum_PostFace=Style_Pic(1)			'��������
	TempForum_emot=Style_Pic(2)				'�������� EM

	Forum_PostFace=split(TempForum_PostFace,"|||")
	Forum_userface=split(TempForum_userface,"|||")
	Forum_emot=split(TempForum_emot,"|||")
	Forum_emotNum=UBound(Forum_emot)
	Forum_userfaceNum=UBound(Forum_userface)
	Forum_PostFaceNum=UBound(Forum_PostFace)
End Sub 
%>

