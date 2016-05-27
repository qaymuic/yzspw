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
	Stype = Cint(Request("Stype"))		'1=表情，2=心情em，3=头像
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
	Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
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
		actname="发贴表情图片"
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
		actname="发贴心情图片"
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
		actname="注册头像"
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

	if request("Submit")="保存设置" then
		call saveconst()
	elseif request("Submit")="恢复默认设置" then
		call savedefault()
	ElseIf request("Submit")="恢复默认总设置" then
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
<td height="23" colspan="4" ><B>说明</B>：<br>①、以下图片均保存于论坛<%=bbspicurl%>目录中，如要更换也请将图片放于该目录<br>②、右边复选框为删除选项，如果选择后点保存设置，则删除相应图片<BR>③、如仅仅修改文件名，可在修改相应选项后直接点击保存设置而不用选择右边复选框
</td>
</tr>
</table>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="4" align=left><%=actname%>管理设置 （目前共有<%=count%>个<%=actname%>图片在文件夹：<%=bbspicurl%>）</th>
</tr>
<tr> 
<td width="20%"  align=left class=forumrow>当前模版名称：</td>
<td width="80%"  align=left class=forumrow colspan="3"><%=StyleName%>
</td>
</tr>
<tr> 
<td width="20%"  align=left class=forumrow>增加的文件名：</td>
<td width="80%"  align=left class=forumrow colspan="3"><input  type="text" name="NEWFILENAME" value="<%=newfilename%>">（<font color=red>建议采用默认，增加后把相应的文件名上传到该目录下。</font>）
</td>
</tr>
<tr> 
<td width="20%"  align=left class=forumrow>批量增加数目：</td>
<td width="80%"  align=left class=forumrow colspan="3"><input  type="text" name="NEWNUM" value="<%=newnum%>">
<input type="submit" name="Submit" value="增加">
</td>
</tr>
<tr> 
<td width="20%"  align=left class=forumrow>覆盖所有模板：</td>
<td width="80%"  align=left class=forumrow colspan="3">是<input type=radio name=coverall value=1 >否<input type=radio name=coverall value=0 checked>
</td>
</tr>
<%
Dim TempName 
IF REQUEST("Submit")="增加" and REQUEST("Newnum")<>"" and request("Newnum")<>0 then
newnum=REQUEST("Newnum")

for i=count to count+newnum-1
if stype=2 and i<10 Then
TempName = newfilename&"0"&i
Else
TempName = newfilename&i
End If
%>
<tr>
<td width="20%" class=forumRowHighlight><%=actname%>ID：<input type=hidden name="face_id<%=i%>" size="10" value="<%=i%>"><%=i%></td>
<td width="75%" class=forumRowHighlight colspan="2">新增加的文件：<input  type="text" name="userface<%=i%>" value="<%=TempName%>.gif"></td>
<td width="5%" class=forumRowHighlight> 
<input type="checkbox" name="delid<%=i%>" value="<%=i%>">
</td>
</tr>
<% next 
end if
%>
<tr>
<th width="20%" class=forumrow>文件</th>
<th width="45%" class=forumrow>文件名</th>
<th width="30%" class=forumrow> 
图片
<th width="5%" class=forumrow> 
删除
</th>
</tr>
<tr>
<td width="20%" class=forumrow>文件目录：<input type=hidden  name="face_id0" size="10" ></td>
<td width="45%" class=forumrow>&nbsp;<input  type="text" name="userface0" value="<%=bbspicurl%>"></td>
<td width="30%" class=forumrow></td>
<td width="5%" class=forumrow></td>
</tr>
<% for i=1 to bbspicmun %>
<tr>
<td width="20%" class=forumrow>文件名：<input type=hidden  name="face_id<%=i%>" size="10" value="<%=i%>"></td>
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
<B>注意</B>：右边复选框为删除选项，如果选择后点保存设置，则删除相应图片<BR>如仅仅修改文件名，可在修改相应选项后直接点击保存设置而不用选择右边复选框
</td>
</tr>
<tr> 
<td  colspan="4" class=forumrow> 
<div align="center"> 
 删除选项：删除所选的实际文件（<font color=red>需要FSO支持功能</font>）：是<input type=radio name=setfso value=1 >否<input type=radio name=setfso value=0 checked> 请选择要删除的文件，<input type="checkbox" name=chkall value=on onclick="CheckAll(this.form)">全选 <BR>
<input type="submit" name="Submit" value="保存设置">
<input type="submit" name="Submit" value="恢复默认设置">
<input type="submit" name="Submit" value="恢复默认总设置">
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
				response.write "删除"&filepaths
			Else
				response.write "未找到"&filepaths
			End If
		End If
	End If
Next
Set objFSO=Nothing
''1=表情，2=心情em，3=头像
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
	Dv_suc(actname&"设置成功。")
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
		''头像---------------------------------------
        for i=1 to 60
        userface=userface&"image"&i&".gif|||"
        next
		userface="Images/userface/|||"+userface
		upconfig=userface+"@@@"
		''表情---------------------------------------
		userface=""
        for i=1 to 18
        userface=userface&"face"&i&".gif|||"
        next
		userface="Skins/default/topicface/|||"+userface
		upconfig=upconfig+userface+"@@@"
		''心情---------------------------------------
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
Dv_suc(actname&"恢复设置成功。")
end sub

'表名：Dv_Style
'字段名：Style_Pic
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
		Errmsg=ErrMsg + "<li>"+"模块未找到，可能已被删除，请重新选取正确模版！"
		Founderr=True
		Exit Sub
	End if
	Rs.close:Set Rs=Nothing
	Style_Pic=Split(Style_Pic,"@@@")	'模版大类以@@@分割；小类以|||分割;
	TempForum_userface=Style_Pic(0)			'用户头像
	TempForum_PostFace=Style_Pic(1)			'发贴表情
	TempForum_emot=Style_Pic(2)				'发贴心情 EM

	Forum_PostFace=split(TempForum_PostFace,"|||")
	Forum_userface=split(TempForum_userface,"|||")
	Forum_emot=split(TempForum_emot,"|||")
	Forum_emotNum=UBound(Forum_emot)
	Forum_userfaceNum=UBound(Forum_userface)
	Forum_PostFaceNum=UBound(Forum_PostFace)
End Sub 
%>

