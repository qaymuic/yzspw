<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<!-- #include file="inc/GroupPermission.asp" -->
<!--#include file=inc/md5.asp-->
<%
	Head()
	Server.ScriptTimeout=999999
	dim Str
	dim admin_flag
	admin_flag="9,10"
	founderr=False 
	if not Dvbbs.master or instr(","&session("flag")&",",",9,")=0 or instr(","&session("flag")&",",",10,")=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		call main()
		footer()
	end if

	sub main()
%>
<table width="95%" border="0" cellspacing="0" cellpadding="0"  align=center class="tableBorder">
<tr> 
<th width="100%" class="tableHeaderText" colspan=2 height=25>��̳����
</th>
</tr>
<tr>
<td class="forumRowHighlight" colspan=2>
<p><B>ע��</B>��<BR>��ɾ����̳ͬʱ��ɾ������̳���������ӣ�ɾ������ͬʱɾ��������̳���������ӣ� ����ʱ��������д����Ϣ��<BR>�����ѡ��<B>��λ���а���</B>�������а��涼����Ϊһ����̳�����ࣩ����ʱ����Ҫ���¶Ը���������й����Ļ������ã�<B>��Ҫ����ʹ�øù���</B>�����������˴�������ö��޷���ԭ����֮��Ĺ�ϵ�������ʱ��ʹ�ã���������Ҳ����ֻ���ĳ��������и�λ����(������ĸ�����������˵�)�������뿴����˵��<BR><font color=blue>ÿ������ĸ��������������˵�������ǰ����ϸ�Ķ�˵�������������˵��бȱ�İ��������˷�������ͷ��ิλ����</font>
</td>
</tr>
<tr>
<td class="forumRowHighlight" height=25>
<B>��̳����ѡ��</B></td>
<td class="forumRowHighlight"><a href="admin_board.asp">��̳������ҳ</a> | <a href="admin_board.asp?action=add">�½���̳����</a> | <a href="?action=settemplates">ģ������������</a> | <a href="?action=orders">һ����������</a> | <a href="?action=boardorders">N����������</a> | <a href="?action=RestoreBoard" onclick="{if(confirm('��λ���а��潫�����а���ָ���Ϊһ������࣬��λ��Ҫ�����а������½��й����Ļ������ã������ز�����ȷ����λ��?')){return true;}return false;}">��λ���а���</a> | <a href="?action=RestoreBoardCache" onclick="{if(confirm('��ʱ��������̳������޸���ǰ̨�������޸�Ч������ܿ�������Ӧ����Ļ���û����Ч���£������ｫ�ؽ����а���Ļ��棬������İ���ܶ࣬�⽫������һ����ʱ�䣬ȷ����?')){return true;}return false;}">�ؽ����滺��</a>
</td>
</tr>
</table>
<p></p>
<%
select case Request("action")
case "add"
	call add()
case "edit"
	call edit()
case "savenew"
	call savenew()
case "savedit"
	call savedit()
case "del"
	call del()
case "orders"
	call orders()
case "updatorders"
	call updateorders()
case "boardorders"
	call boardorders()
case "updatboardorders"
	call updateboardorders()
case "addclass"
	call addclass()
case "saveclass"
	call saveclass()
case "del1"
	call del1()
case "mode"
	call mode()
case "savemod"
	call savemod()
case "permission"
	call boardpermission()
case "editpermission"
	call editpermission()
case "RestoreBoard"
	call RestoreBoard()
Case "RestoreBoardCache"
	Call RestoreBoardCache()
Case "clearDate"
	Call clearDate
Case "delDate"
	Call delDate
Case "RestoreClass"
	Call RestoreClass
Case "handorders"
	Call handorders
Case "savehandorders"
	Call savehandorders
Case "savesid"
	Call savesid
Case "upallsid"
	Call upallsid
Case "settemplates"
	Call Settemplates
Case else
	call boardinfo()
end select
end Sub
Sub upallsid()
	Dim Sid,cid
	SID= Request("Sid")
	Cid=Request("cid")
	Dvbbs.Execute("Update Dv_board set Sid="&CLng(SID)&",cid="&cid&"")
	Call Dvbbs.LoadBoardsInfo()
	Dv_suc("��̳ģ��ͳһ���óɹ�!")
End Sub
Sub savesid
	Dim i,boardid,TempStr
	Dim Templateslist,sid,j,bid,cid
	sid=""
	For Each TempStr in Request.form("upboardid")
		If Bid="" Then
			Bid=TempStr
		Else
			Bid=Bid&","&TempStr
		End If 
	Next
	Bid=split(Bid,",")
	For i=0 to UBound(bid)
		If sid="" Then
			sid=Request("sid"&bid(i))
			cid=Request("cid"&bid(i))
		Else
			sid=sid&","&Request("sid"&bid(i))
			cid=cid&","&Request("cid"&bid(i))
		End If
	Next
	sid=split(sid,",")
	Cid=split(cid,",")
	Dvbbs.Name="Templateslist"
	If Dvbbs.ObjIsEmpty() Then  Dvbbs.ReloadTemplateslist()
	Templateslist= Dvbbs.Value
	Templateslist=split(Templateslist,"@@@")
	For i=0 to UBound(Templateslist)
		templateslist(i)=split(Templateslist(i),"|||")
		boardid=""
		For J=0 to UBound(Bid)
			If CLng(Templateslist(i)(0))=CLng(Sid(j)) Then
				If boardid="" Then 
					boardid=bid(j)
				Else
					boardid=boardid&","&bid(j)
				End If
			End If	
		Next
		If boardid<>"" Then
			'����SID
			'Response.Write "Update Dv_board set Sid="&CLng(Templateslist(i)(0))&" Where BoardId In("&Boardid&") "		
			Dvbbs.Execute("Update Dv_board set Sid="&CLng(Templateslist(i)(0))&" Where BoardId In("&Boardid&") ")
		End If
	Next
	'����cid
	For i=0 to UBound(bid)	
		Dvbbs.Execute("Update Dv_board set cid="&CLng(cid(i))&" Where BoardId="&bid(i)&" ")
	Next 
	Call Dvbbs.LoadBoardsInfo()
	Dv_suc("��̳ģ���������óɹ�!")
End Sub 
Sub Settemplates
Dim reBoard_Setting,MoreMenu,i
Dim Templateslist
Dvbbs.Name="Templateslist"
If Dvbbs.ObjIsEmpty() Then  Dvbbs.ReloadTemplateslist()
Templateslist= Dvbbs.Value
Templateslist=split(Templateslist,"@@@")
For i=0 to UBound(Templateslist)
	templateslist(i)=split(Templateslist(i),"|||")
Next
%>
<form action ="admin_board.asp?action=upallsid" method=post name="dv">
<table cellspacing="0" cellpadding="0" align=center Class="tableBorder" style="width:98%" >
<tr> 
<th colspan="2" class="tableHeaderText" align=center height=25>ģ �� ͳ һ �� ��
</th>
</tr>
<tr>
<td width=300 align=Left  class="forumRowHighlight" ><B>������̳����Ϊ��</b>&nbsp; ģ��
<script language="javascript">
<%
Dim cssdata
Response.Write "var StyleId="&Dvbbs.cachedata(17,0)&";"
Response.Write "var Cssid="&Dvbbs.cachedata(30,0)&";"
For i=0 to UBound(Templateslist)
	Dvbbs.SkinID=Templateslist(i)(0)
	Dvbbs.name="Forum_CSS"&Templateslist(i)(0)
	If Dvbbs.ObjIsEmpty() Then Dvbbs.TemplatesToCache ("Forum_CSS")
	cssdata=Dvbbs.value
	cssdata=Split(cssdata,"@@@")
	Response.Write "var css_Option"&Templateslist(i)(0)&"='"&cssdata(0)&"';"
	Response.Write chr(10)
	
Next 
%>
</script>
<select name="sid" onChange="Changeoption(this.value)" >
<%
For i=0 to UBound(Templateslist)
Response.Write "<option value="""&Templateslist(i)(0)&""""
If CLng(Templateslist(i)(0)) = CLng(Dvbbs.cachedata(17,0)) Then 
	Response.Write " selected"
End If 
Response.Write ">"&Templateslist(i)(1)&"</option>"
Next 
%>
</Select>
</td>
<td width=300 align=Left  class="forumRowHighlight" >
&nbsp;��� 
<select name=cid >
<option value="" >ѡ����Ƥ��</option>
</select>
<Input type="submit" name="Submit" value="�� ��"></td>
</tr>
</table><BR>
</form>
<SCRIPT LANGUAGE="JavaScript">

<!--
function Changeoption(sid)
{
var NewOption=eval("css_Option"+sid).split("|||");
var j=eval('document.dv.cid.length;');
	for (i=0;i<j;i++){
		eval('document.dv.cid.options[j-i]=null;')
	}
	for (i=0;i<NewOption.length-1;i++){
		tempoption=new Option(NewOption[i],i);
		eval('document.dv.cid.options[i]=tempoption;');
		if (Cssid==i&&sid==StyleId){
		eval('document.dv.cid.options[i].selected=true;');
		}
	}
}
var forum_sid=eval('document.dv.sid.value;');
Changeoption(forum_sid);
//-->
</SCRIPT>
<form action ="admin_board.asp?action=savesid" method=post name="dv1">
<table cellspacing="0" cellpadding="0" align=center Class="tableBorder" style="width:98%" >
<tr> 
<th width="70%" class="tableHeaderText" align=Left height=25>��̳����
</th>
<th width="30%" class="tableHeaderText" align=Left height=25>����ģ��
</th>
</tr>
<%
dim classrow
sql="select * from dv_board order by rootid,orders"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
do while not rs.eof
reBoard_Setting=split(rs("Board_setting"),",")
if classrow="forumRowHighlight" then
	classrow="forumRow"
else
	classrow="forumRowHighlight"
end if
%>
<tr> 
<td height="25"  class="<%=classrow%>">
<%if rs("depth")>0 then%>
<%for i=1 to rs("depth")%>
&nbsp;
<%next%>
<%end if%>
<%if rs("child")>0 then%><img src="skins/default/plus.gif"><%else%><img src="skins/default/nofollow.gif"><%end if%>
<%if rs("parentid")=0 then%><b><%end if%><%=rs("boardtype")%><%if rs("child")>0 then%>(<%=rs("child")%>)<%end if%>
<%if rs("parentid")=0 then%></b><%end if%>
</td>
<td align=Left  class="<%=classrow%>" >
<select name="sid<%=Rs("BoardID")%>" onChange="Changeoption<%=RS("BoardID")%>(this.value)" >
<%
For i=0 to UBound(Templateslist)
Response.Write "<option value="""&Templateslist(i)(0)&""""
If CLng(Templateslist(i)(0)) = Rs("Sid") Then 
	Response.Write " selected"
End If 
Response.Write ">"&Templateslist(i)(1)&"</option>"
Next 
%>
</select>
<select name=cid<%=Rs("BoardID")%> >
<option value="" >ѡ����Ƥ��</option>
</select>
<Input type="hidden" name="upboardid" value="<%=rs("boardid")%>">
</td></tr>
<script language="javascript">
<%
Response.Write "var StyleId"&Rs("BoardID")&"="&Rs("Sid")&";"
Response.Write "var Cssid"&Rs("BoardID")&"="&Rs("Cid")&";"
%>
</script>
<SCRIPT LANGUAGE="JavaScript">
<!--
function Changeoption<%=Rs("BoardID")%>(sid)
{
var NewOption=eval("css_Option"+sid).split("|||");
var j=eval('document.dv1.cid<%=Rs("BoardID")%>.length;');
	for (i=0;i<j;i++){
		eval('document.dv1.cid<%=Rs("BoardID")%>.options[j-i]=null;')
	}
	for (i=0;i<NewOption.length-1;i++){
		tempoption=new Option(NewOption[i],i);
		eval('document.dv1.cid<%=Rs("BoardID")%>.options[i]=tempoption;');
		if (Cssid<%=Rs("BoardID")%>==i&&sid==StyleId<%=Rs("BoardID")%>){
		eval('document.dv1.cid<%=Rs("BoardID")%>.options[i].selected=true;');
		}
	}
}
var forum_sid=eval('document.dv1.sid<%=Rs("BoardID")%>.value;');
Changeoption<%=Rs("BoardID")%>(StyleId);
//-->
</SCRIPT>
<%
Rs.movenext
loop
set rs=nothing
%>
<tr>
<td width=300 align=Left  class="forumRowHighlight" ></td>
<td width=300 align=Left  class="forumRowHighlight" ><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</table><BR><BR>
</form>

<%
End Sub 
sub boardinfo()
Dim reBoard_Setting,MoreMenu
Dim classrow,iii
%>
<table width="95%" cellspacing="0" cellpadding="0" align=center class="tableBorder">
<tr> 
<th width="35%" class="tableHeaderText" height=25>��̳����
</th>
<th width="35%" class="tableHeaderText" height=25>����
</th>
</tr>
<%
SQL="select boardid,boardtype,parentid,depth,child,Board_setting from dv_board order by rootid,orders"
SET Rs = Conn.Execute(SQL)
If Rs.eof Then
	Rs.close:Set Rs = Nothing
Else

SQL=Rs.GetRows(-1)
Rs.close:Set Rs = Nothing
For iii=0 To Ubound(SQL,2)
	reBoard_Setting=split(SQL(5,iii),",")
	if classrow="forumRowHighlight" then
		classrow="forumRow"
	else
		classrow="forumRowHighlight"
	end if
	Response.Write "<tr>"
	Response.Write "<td height=""25"" width=""35%""  class="
	Response.Write classrow 
	Response.Write ">"
	if SQL(3,iii)>0 then
		for i=1 to SQL(3,iii)
			Response.Write "&nbsp;&nbsp;"
		next
	end if
	if SQL(4,iii)>0 then
		Response.Write "<img src=""skins/default/plus.gif"">"
	else
		Response.Write "<img src=""skins/default/nofollow.gif"">"
	end if
	if SQL(2,iii)=0 then
		Response.Write "<b>"
	end if
	Response.Write SQL(1,iii)
	if SQL(4,iii)>0 then
		Response.Write "("
		Response.Write SQL(4,iii)
		Response.Write ")"
	end if
%>
</td>
<td width=65% align=center class="<%=classrow%>">
<a href="admin_board.asp?action=add&editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��Ӱ���</U></font></a> | <a href="admin_board.asp?action=edit&editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��������</U></font></a> | <a href="admin_BoardSetting.asp?editid=<%=SQL(0,iii)%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>�߼�����</U></font></a>
<%
if reBoard_Setting(2)=0 then
	MoreMenu=MoreMenu & "<div class=menuitems><a href=admin_vipboard.asp?boardid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><U>VIP��̳����</U></font></a></div>"
elseif reBoard_Setting(2)=0 and reBoard_Setting(46)>0 then
	MoreMenu=MoreMenu & "<div class=menuitems><a href=admin_vipboard.asp?boardid="&SQL(0,iii)&"&action=reinstall><font color="&Dvbbs.mainsetting(3)&"><U>����VIP��̳</U></font></a></div>"
elseif reBoard_Setting(2)=1 and reBoard_Setting(46)>0 then
	MoreMenu=MoreMenu & "<div class=menuitems><a href=admin_vipboard.asp?action=showvipuser&boardid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><U>�鿴VIP�û�</U></font></a></div>"
end if

if reBoard_Setting(2)=1 then
	MoreMenu=MoreMenu & "<div class=menuitems><a href=admin_board.asp?action=mode&boardid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><U>��֤�û�</U></font></a></div>"
end if

MoreMenu=MoreMenu & "<div class=menuitems><a href=admin_update.asp?action=updat&submit=������̳����&boardid="&SQL(0,iii)&" title=�������ظ������������ظ���><font color="&Dvbbs.mainsetting(3)&"><U>��������</U></font></a></div><div class=menuitems><a href=# onclick=alertreadme(\'��ս���������̳�����������ڻ���վ��ȷ�������?\',\'admin_update.asp?action=delboard&boardid="&SQL(0,iii)&"\')><font color="&Dvbbs.mainsetting(3)&"><U>��հ�������</U></font></a></div>"

if SQL(4,iii)=0 then
MoreMenu=MoreMenu & "<div class=menuitems><a href=# onclick=alertreadme(\'ɾ������������̳���������ӣ�ȷ��ɾ����?\',\'admin_board.asp?action=del&editid="&SQL(0,iii)&"\')><font color="&Dvbbs.mainsetting(3)&"><U>ɾ������</U></font></a></div>"
else
MoreMenu=MoreMenu & "<div class=menuitems><a href=# onclick=alertreadme(\'����̳����������̳��������ɾ����������̳����ɾ������̳��\',\'#\')><font color="&Dvbbs.mainsetting(3)&"><U>ɾ������</U></font></a></div>"
end if
MoreMenu=MoreMenu & "<div class=menuitems><a href=admin_Board.asp?action=clearDate&boardid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><u>��������</u></font></a></div>"
If SQL(2,iii)=0 Then
	MoreMenu=MoreMenu & "<div class=menuitems><a href=# onclick=alertreadme(\'��λ�÷��ཫ��Ѹ÷����µ����а��涼��λ�ɶ������棬����ԭ���Ķ༶���඼����λ�ɶ������棬�����ز�����ȷ����λ��?\',\'?action=RestoreClass&classid="&SQL(0,iii)&"\')><font color="&Dvbbs.mainsetting(3)&"><u>��λ�÷���</u></font></a></div><div class=menuitems><a href=?action=handorders&classid="&SQL(0,iii)&"><font color="&Dvbbs.mainsetting(3)&"><u>��������(�ֶ�)</u></font></a></div>"
End If
%>
 | <a href="#" onMouseOver="showmenu(event,'<%=MoreMenu%>')" style="CURSOR:hand"><font color=<%=Dvbbs.mainsetting(3)%>><u>�������</u></font></a>
</td></tr>
<%
MoreMenu=""
Next
End If
%>
</table><BR><BR>
<SCRIPT LANGUAGE="JavaScript">
<!--
function alertreadme(str,url){
{if(confirm(str)){
location.href=url;
return true;
}return false;}
}
//-->
</SCRIPT>
<%
end sub

sub add()
dim rs_c
Dim forum_sid,forum_cid,Style_Option,TempOption
set rs_c= server.CreateObject ("adodb.recordset")
sql = "select * from dv_board order by rootid,orders"
rs_c.open sql,conn,1,1
	dim boardnum
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select Max(boardid) from dv_board"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
	boardnum=1
	else
	boardnum=rs(0)+1
	end if
	if isnull(boardnum) then boardnum=1
	if boardnum=444 then boardnum=445
	if boardnum=777 then boardnum=778
	rs.close
%>
<form action ="admin_board.asp?action=savenew" method=post name=theform>
<input type="hidden" name="newboardid" value=<%=boardnum%>>
<table width="85%" border="0" cellspacing="1" cellpadding="0"  align=center class="tableBorder">
<tr> 
<th height=24 colspan=2><B>�������̳</th>
</tr>
<tr> 
<td width="100%" height=30 class="forumrowHighLight" colspan=2>
˵����<BR>1�������̳�������ص����þ�ΪĬ�����ã��뷵����̳���������ҳ�����б�ĸ߼����������ø���̳����Ӧ���ԣ��������Ը���̳���������Ȩ�����ã��뵽<A HREF="admin_board.asp?action=permission"><font color=blue>��̳Ȩ�޹���</font></A>��������Ӧ�û����ڸð����Ȩ�ޡ�<BR>
2��<font color=blue>�������ӵ�����̳����</font>��ֻ��Ҫ������������ѡ����Ϊ��̳���༴�ɣ�<font color=blue>�������ӵ�����̳����</font>����Ҫ������������ȷ����ѡ�����̳������ϼ�����
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow">��̳����</td>
<td width="60%" class="forumrow"> 
<input type="text" name="boardtype" size="35">
</td>
</tr>
<tr> 
<td width="40%" height=24 class="forumrow">����˵��<BR>����ʹ��HTML����</td>
<td width="60%" class="forumrow"> 
<textarea name="Readme" cols="40" rows="3"></textarea>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>�������</U></td>
<td width="60%" class="forumrow"> 
<select name=class>
<option value="0">��Ϊ��̳����</option>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <%if request("editid")<>"" and clng(request("editid"))=rs_c("boardid") then%>selected<%end if%>>
<%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>

<%next%>
<%end if%><%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>ʹ������ģ��</U><BR>���ģ���а�����̳��ɫ��ͼƬ<BR>����Ϣ</td>
<td width="60%" class="forumrow">
<%
	set rs_c= server.CreateObject ("adodb.recordset")
	sql = "select id,StyleName,Forum_CSS from dv_style"
	rs_c.open sql,conn,1,1
	if rs_c.eof and rs_c.bof then
		response.write "������ӷ��"
	else
		sql=rs_c.GetRows(-1)
		forum_sid=SQL(0,0)
		forum_cid=0
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">"
		Response.Write chr(10)
		Response.Write "var StyleId="&forum_sid&";"
		Response.Write "var Cssid="&forum_cid&";"
		Response.Write chr(10)
		For i=0 To Ubound(SQL,2)
			Style_Option=Style_Option+"<option value="
			Style_Option=Style_Option&SQL(0,i)
			If forum_sid=SQL(0,i) Then Style_Option=Style_Option+" selected "
			Style_Option=Style_Option+" >"+SQL(1,i)+"</option>"
			TempOption=Split(SQL(2,i),"@@@")
			Response.Write "var css_Option"&SQL(0,i)&"='"&TempOption(0)&"';"
			Response.Write chr(10)
		Next
		Response.Write "</SCRIPT>"
	End If
	rs_c.close:Set rs_c=Nothing
%>
ģ�壺<select name=sid onChange="Changeoption(this.value)">
<%
Response.Write Style_Option
%>
</select>
 ���<select name=cid onChange="">
<option value="" >ѡ����Ƥ��</option>
</select>
<SCRIPT LANGUAGE="JavaScript">
<!--
function Changeoption(sid)
{
var NewOption=eval("css_Option"+sid).split("|||");
var j=eval('document.theform.cid.length;');
	for (i=0;i<j;i++){
		eval('document.theform.cid.options[j-i]=null;')
	}
	for (i=0;i<NewOption.length-1;i++){
		tempoption=new Option(NewOption[i],i);
		eval('document.theform.cid.options[i]=tempoption;');
		if (Cssid==i&&sid==StyleId){
		eval('document.theform.cid.options[i].selected=true;');
		}
	}
}
var forum_sid=eval('document.theform.sid.value;');
Changeoption(forum_sid);
//-->
</SCRIPT>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>��̳����</U><BR>������������|�ָ����磺ɳ̲С��|wodeail</td>
<td width="60%" class="forumrow"> 
<input type="text" name="boardmaster" size="35">
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>��ҳ��ʾ��̳ͼƬ</U><BR>��������ҳ��̳����������<BR>��ֱ����дͼƬURL</td>
<td width="60%" class="forumrow">
<input type="text" name="indexIMG" size="35">
</td>
</tr>
<tr> 
<td width="40%" height=24 class="forumRow">&nbsp;</td>
<td width="60%" class="forumRow"> 
<input type="submit" name="Submit" value="�����̳">
</td>
</tr>
</table>
</form>
<%
set rs_c=nothing
set rs=nothing
end sub

sub edit()
dim rs_c,reBoard_Setting
Dim forum_sid,forum_cid,Style_Option,TempOption
sql = "select * from dv_board order by rootid,orders"
set rs_c=Dvbbs.Execute(sql)
sql = "select * from dv_board where boardid="&request("editid")
set rs=Dvbbs.Execute(sql)
reBoard_Setting=split(rs("Board_setting"),",")

forum_sid=rs("sid")
forum_cid=rs("cid")
%>        
<form action ="admin_board.asp?action=savedit" method=post name=theform>       
<input type="hidden" name=editid value="<%=Request("editid")%>">
<table width="85%" border="0" cellspacing="1" cellpadding="0"  align=center class="tableBorder">
<tr> 
<th height=24 colspan=2>�༭��̳��<%=rs("boardtype")%></th>
</tr>
<tr> 
<td width="100%" height=30 class="forumrowHighLight" colspan=2>
˵����<BR>1�������̳�������ص����þ�ΪĬ�����ã��뷵����̳���������ҳ�����б�ĸ߼����������ø���̳����Ӧ���ԣ��������Ը���̳���������Ȩ�����ã��뵽<A HREF="admin_board.asp?action=permission"><font color=blue>��̳Ȩ�޹���</font></A>��������Ӧ�û����ڸð����Ȩ�ޡ�<BR>
2��<font color=blue>�������ӵ�����̳����</font>��ֻ��Ҫ������������ѡ����Ϊ��̳���༴�ɣ�<font color=blue>�������ӵ�����̳����</font>����Ҫ������������ȷ����ѡ�����̳������ϼ�����
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow">��̳����</td>
<td width="60%" class="forumrow"> 
<input type="text" name="boardtype" size="35"  value="<%=Server.htmlencode(rs("boardtype"))%>" >
</td>
</tr>
<tr> 
<td width="40%" height=24 class="forumrow">����˵��<BR>����ʹ��HTML����</td>
<td width="60%" class="forumrow"> 
<textarea name="Readme" cols="40" rows="3"><%=server.HTMLEncode(Rs("readme")&"")%></textarea>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>�������</U><BR>������̳����ָ��Ϊ��ǰ����<BR>������̳����ָ��Ϊ��ǰ�����������̳</td>
<td width="60%" class="forumrow"> 
<select name=class>
<option value="0">��Ϊ��̳����</option>
<% do while not rs_c.EOF%>
<option value="<%=rs_c("boardid")%>" <% if cint(rs("parentid")) = rs_c("boardid") then%> selected <%end if%>><%if rs_c("depth")>0 then%>
<%for i=1 to rs_c("depth")%>
��
<%next%>
<%end if%><%=rs_c("boardtype")%></option>
<%
rs_c.MoveNext 
loop
rs_c.Close 
%>
</select>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>ʹ������ģ��</U><BR>���ģ���а�����̳��ɫ��ͼƬ<BR>����Ϣ</td>
<td width="60%" class="forumrow">
<%
	set rs_c= server.CreateObject ("adodb.recordset")
	sql = "select id,StyleName,Forum_CSS from dv_style"
	rs_c.open sql,conn,1,1
	if rs_c.eof and rs_c.bof then
		response.write "������ӷ��"
	else
		sql=rs_c.GetRows(-1)
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">"
		Response.Write chr(10)
		Response.Write "var StyleId="&forum_sid&";"
		Response.Write "var Cssid="&forum_cid&";"
		Response.Write chr(10)
		For i=0 To Ubound(SQL,2)
			Style_Option=Style_Option+"<option value="
			Style_Option=Style_Option&SQL(0,i)
			If forum_sid=SQL(0,i) Then Style_Option=Style_Option+" selected "
			Style_Option=Style_Option+" >"+SQL(1,i)+"</option>"
			TempOption=Split(SQL(2,i),"@@@")
			Response.Write "var css_Option"&SQL(0,i)&"='"&TempOption(0)&"';"
			Response.Write chr(10)
		Next
		Response.Write "</SCRIPT>"
	End If
	rs_c.close:Set rs_c=Nothing
%>
ģ�壺<select name=sid onChange="Changeoption(this.value)">
<%
Response.Write Style_Option
%>
</select>
 ���<select name=cid onChange="">
<option value="" >ѡ����Ƥ��</option>
</select>
<SCRIPT LANGUAGE="JavaScript">
<!--
function Changeoption(sid)
{
var NewOption=eval("css_Option"+sid).split("|||");
var j=eval('document.theform.cid.length;');
	for (i=0;i<j;i++){
		eval('document.theform.cid.options[j-i]=null;')
	}
	for (i=0;i<NewOption.length-1;i++){
		tempoption=new Option(NewOption[i],i);
		eval('document.theform.cid.options[i]=tempoption;');
		if (Cssid==i&&sid==StyleId){
		eval('document.theform.cid.options[i].selected=true;');
		}
	}
}
var forum_sid=eval('document.theform.sid.value;');
Changeoption(forum_sid);
//-->
</SCRIPT>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>��̳����</U><BR>������������|�ָ����磺ɳ̲С��|wodeail</td>
<td width="60%" class="forumrow"> 
<input type="text" name="boardmaster" size="35"  value='<%=rs("boardmaster")%>'>
<input type="hidden" name="oldboardmaster" value='<%=rs("boardmaster")%>'>
</td>
</tr>
<tr> 
<td width="40%" height=30 class="forumrow"><U>��ҳ��ʾ��̳ͼƬ</U><BR>��������ҳ��̳����������<BR>��ֱ����дͼƬURL</td>
<td width="60%" class="forumrow">
<input type="text" name="indexIMG" size="35" value="<%=enfixjs(rs("indexIMG"))%>">
</td>
</tr>
<tr> 
<td width="40%" height=24 class="forumrow">&nbsp;</td>
<td width="60%" class="forumrow"> 
<input type="submit" name="Submit" value="�ύ�޸�">
</td>
</tr>
<tr> 
<td width="100%" height=30 class="forumrowHighLight" colspan=2 align=right>
<a href="admin_board.asp?action=add&editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��Ӱ���</U></font></a> | <a href="admin_board.asp?action=edit&editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��������</U></font></a> | <a href="admin_BoardSetting.asp?editid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>�߼�����</U></font></a>
<%if reBoard_Setting(2)=1 then%>
| <a href="admin_board.asp?action=mode&boardid=<%=Request("editid")%>"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��֤�û�</U></font></a>
<%end if%>
| <a href="admin_update.asp?action=updat&submit=������̳����&boardid=<%=Request("editid")%>" title="�������ظ������������ظ���"><font color="<%=Dvbbs.mainsetting(3)%>"><U>��������</U></font></a> | <a href="admin_update.asp?action=delboard&boardid=<%=Request("editid")%>" onclick="{if(confirm('��ս���������̳�����������ڻ���վ��ȷ�������?')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>���</U></font></a> | <%if rs("child")=0 then%><a href="admin_board.asp?action=del&editid=<%=Request("editid")%>" onclick="{if(confirm('ɾ������������̳���������ӣ�ȷ��ɾ����?')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>ɾ��</U></a><%else%><a href="#" onclick="{if(confirm('����̳����������̳��������ɾ����������̳����ɾ������̳��')){return true;}return false;}"><font color="<%=Dvbbs.mainsetting(3)%>"><U>ɾ��</U></a><%end if%>
| <a href="admin_Board.asp?action=clearDate&boardid=<%=Request("editid")%>"> <font color="<%=Dvbbs.mainsetting(3)%>"><u>��������</u></a>
</td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
set rs_c=nothing
end sub
sub mode()
dim boarduser
%>
<form action ="admin_board.asp?action=savemod" method=post>
<table width="95%" class="tableBorder" cellspacing="1" cellpadding="1" align="center">
<tr> 
<th width="52%" height=22>˵����</th>
<th width="48%">������</th>
</tr>
<tr> 
<td width="52%" height=22 class=forumrow><B>��̳����</B></td>
<td width="48%" class=forumrow> 
<%
set rs= server.CreateObject ("adodb.recordset")
sql="select boardid,boardtype,boarduser from dv_board where boardid="&request("boardid")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write "�ð��沢�����ڻ��߸ð��治�Ǽ��ܰ��档"
Else
response.write rs(1)
response.write "<input type=hidden value="&rs(0)&" name=boardid>"
boarduser=rs(2)
end if
rs.close
set rs=nothing
%>
</td>
</tr>
<tr> 
<td width="52%" class=forumrow><B>��֤�û�</B>��<br>
ֻ���趨Ϊ��֤��̳����̳��Ҫ��д�ܹ�����ð�����û���ÿ����һ���û���ȷ���û�������̳�д��ڣ�ÿ���û�����<B>�س�</B>�ֿ�</font>
</td>
<td width="48%" class=forumrow> 
<textarea cols=35 rows=6 name="vipuser">
<%if not isnull(boarduser) or boarduser<>"" then
	response.write Replace(boarduser,",",Chr(10))
end if%></textarea>
</td>
</tr>
<tr> 
<td width="52%" height=22 class=forumrow>&nbsp;</td>
<td width="48%" class=forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</table>
</form>
<%
End Sub 

'����༭��̳��֤�û���Ϣ
'��ڣ��û��б��ַ���
sub savemod()
dim boarduser
dim boarduser_1
dim userlen
dim updateinfo

If trim(request("vipuser"))<>"" then
	boarduser=request("vipuser")
	boarduser=split(boarduser,chr(13)&chr(10))
	for i = 0 to ubound(boarduser)
	if not (boarduser(i)="" or boarduser(i)=" ") then
		boarduser_1=""&boarduser_1&""&boarduser(i)&","
	end if
	next
	userlen=len(boarduser_1)
	if boarduser_1<>"" then
		boarduser=left(boarduser_1,userlen-1)
		updateinfo=" boarduser='"&boarduser&"' "
		Dvbbs.Execute("update dv_board set "&updateinfo&" where boardid="&request("boardid"))
		Dv_suc("��̳���óɹ�!<LI>�ɹ������֤�û���"&boarduser&"<LI><a href=""?action=RestoreBoardCache"" >��ִ���ؽ����滺�������Ч</a><br>")
	else
		response.write "<p><font color=red>��û�������֤�û�</font><br><br>"
		Exit Sub
	end if
Else
	response.write "<p><font color=red>��û�������֤�û�</font><br><br>"
End If

End Sub

'���������̳��Ϣ
sub savenew()
if request("boardtype")="" then
	Errmsg=Errmsg+"<br>"+"<li>��������̳���ơ�"
	founderr=true
end if
if request("class")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ѡ����̳���ࡣ"
	founderr=true
end if
if request("readme")="" then
	Errmsg=Errmsg+"<br>"+"<li>��������̳˵����"
	founderr=true
end if
if founderr=true then
	dvbbs_error()
	exit sub
end if
dim boardid
dim rootid
dim parentid
dim depth
dim orders
dim Fboardmaster
dim maxrootid
dim parentstr
if request("class")<>"0" then
set rs=Dvbbs.Execute("select rootid,boardid,depth,orders,boardmaster,ParentStr from dv_board where boardid="&request("class"))
rootid=rs(0)
parentid=rs(1)
depth=rs(2)
orders=rs(3)
if depth+1>20 then
	Errmsg="����̳�������ֻ����20������"
	dvbbs_error()
	exit sub
end if
parentstr=rs(5)
else
set rs=Dvbbs.Execute("select max(rootid) from dv_board")
maxrootid=rs(0)+1
if isnull(MaxRootID) then MaxRootID=1
end if
sql="select boardid from dv_board where boardid="&request("newboardid")
set rs=Dvbbs.Execute(sql)
if not (rs.eof and rs.bof) then
	Errmsg="������ָ���ͱ����̳һ������š�"
	dvbbs_error()
	exit sub
else
	boardid=request("newboardid")
end if

dim trs,forumuser,setting
set trs=Dvbbs.Execute("select * from dv_setup")
Setting=Split(trs("Forum_Setting"),"|||")
forumuser=Setting(2)
set rs = server.CreateObject ("adodb.recordset")
sql = "select * from dv_board"
rs.Open sql,conn,1,3
rs.AddNew
if request("class")<>"0" then
rs("depth")=depth+1
rs("rootid")=rootid
rs("orders") = Request.form("newboardid")
rs("parentid") = Request.Form("class")
if ParentStr="0" then
rs("ParentStr")=Request.Form("class")
else
rs("ParentStr")=ParentStr & "," & Request.Form("class")
end if
else
rs("depth")=0
rs("rootid")=maxrootid
rs("orders")=0
rs("parentid")=0
rs("parentstr")=0
end if
rs("boardid") = Request.form("newboardid")
rs("boardtype") = request.form("boardtype")

rs("readme") =fixjs(Request.form("readme"))
rs("TopicNum") = 0
rs("PostNum") = 0
rs("todaynum") = 0
rs("child")=0
rs("LastPost")="$0$"&Now()&"$$$$$"
rs("Board_Setting")="0,0,0,0,1,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,0,0,100,30,10,9,12,1,10,10,0,0,0,0,1,4,0,1,4,0,0,0,200,0,0,0,0,0,0,0,1,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
rs("sid")=request.form("sid")
rs("cid")=request.form("cid")
rs("board_ads")=trs("forum_ads")
rs("board_user")=forumuser
if Request("boardmaster")<>"" then
	rs("boardmaster") = Request.form("boardmaster")
end if
if request.form("indexIMG")<>"" then
	rs("indexIMG")=fixjs(request.form("indexIMG"))
end if
rs.Update 
rs.Close
if Request("boardmaster")<>"" then call addmaster(Request("boardmaster"),"none",0)
if request("class")<>"0" then
if depth>0 then
	'���ϼ�������ȴ���0��ʱ��Ҫ�����丸�ࣨ����ĸ��ࣩ�İ��������������
	for i=1 to depth
		'�����丸�������
		if parentid<>"" then
		Dvbbs.Execute("update dv_board set child=child+1 where boardid="&parentid)
		end if
		'�õ��丸��ĸ���İ���ID
		set rs=Dvbbs.Execute("select parentid from dv_board where boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
		end if
		'��ѭ����������1�������е����һ��ѭ����ʱ��ֱ�ӽ��и���
		if i=depth and parentid<>"" then
		Dvbbs.Execute("update dv_board set child=child+1 where boardid="&parentid)
		end if
	next
	'���¸ð��������Լ����ڱ���Ҫ��ͬ�ڱ������µİ����������
	Dvbbs.Execute("update dv_board set orders=orders+1 where rootid="&rootid&" and orders>"&orders)
	Dvbbs.Execute("update dv_board set orders="&orders&"+1 where boardid="&Request.form("newboardid"))
else
	'���ϼ��������Ϊ0��ʱ��ֻҪ�����ϼ�����������͸ð���������ż���
	Dvbbs.Execute("update dv_board set child=child+1 where boardid="&request("class"))
	set rs=Dvbbs.Execute("select max(orders) from dv_board where boardid="&Request.form("newboardid"))
	Dvbbs.Execute("update dv_board set orders="&rs(0)&"+1 where boardid="&Request.form("newboardid"))
end if
end if
dv_suc("��̳��ӳɹ���<br>����̳Ŀǰ�߼�����ΪĬ��ѡ�������������̳���������������ø���̳�ĸ߼�ѡ�<A HREF=admin_BoardSetting.asp?editid="&Request.form("newboardid")&">����˴�����ð���߼�����</A><br>" & str)
set rs=nothing
trs.close
set trs=nothing
Dvbbs.ReloadAllBoardInfo()
Dvbbs.Name="setup"
Dvbbs.ReloadSetup
Dvbbs.CacheData=Dvbbs.value

Set tRs=Dvbbs.Execute("Select RootID From Dv_Board Where BoardID="&Request.form("newboardid"))
Dim UpdateRootID
UpdateRootID = tRs(0)
Set tRs=dvbbs.Execute("Select BoardID From Dv_Board Where RootID="&UpdateRootID&" Order By Orders")
Do While Not tRs.Eof
	Dvbbs.ReloadBoardInfo(tRs(0))
tRs.Movenext
Loop
Set tRs=Nothing

Dvbbs.DelCahe "BoardJumpList"
Dvbbs.DelCahe "MyAllBoardList"
end sub

'����༭��̳��Ϣ
sub savedit()
if clng(request("editid"))=clng(request("class")) then
	Errmsg="������̳����ָ���Լ�"
	dvbbs_error()
	exit sub
end if
dim newboardid,maxrootid
dim parentid,boardmaster,depth,child,ParentStr,rootid,iparentid,iParentStr
dim trs,brs,mrs
Dim iii
set rs = server.CreateObject ("adodb.recordset")
sql = "select * from dv_board where boardid="&request("editid")
rs.Open sql,conn,1,3
newboardid=rs("boardid")
parentid=rs("parentid")
iparentid=rs("parentid")
boardmaster=rs("boardmaster")
ParentStr=rs("ParentStr")
depth=rs("depth")
child=rs("child")
rootid=rs("rootid")
'�ж���ָ������̳�Ƿ���������̳
if ParentID=0 then
	if clng(request("class"))<>0 then
	set trs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("class"))
	if rootid=trs(0) then
		errmsg="������ָ���ð����������̳��Ϊ������̳1"
		dvbbs_error()
		exit sub
	end if
	end if
else
	set trs=Dvbbs.Execute("select boardid from dv_board where ParentStr like '%"&ParentStr&","&newboardid&"%' and boardid="&request("class"))
	if not (trs.eof and trs.bof) then
		errmsg="������ָ���ð����������̳��Ϊ������̳2"
		dvbbs_error()
		exit sub
	end if
end if
if parentid=0 then
	parentid=rs("boardid")
	iparentid=0
end if
rs("boardtype") = Request.Form("boardtype")	'ȡ��JS���ˡ�
'rs("parentid") = Request.Form("class")
rs("boardmaster") = Request("boardmaster")
rs("readme") = fixjs(Request("readme"))
rs("indexIMG")=Simencodejs(request.form("indexIMG"))
rs("sid")=Cint(request.form("sid"))
rs("cid")=Cint(request.form("cid"))
rs.Update 
rs.Close
set rs=nothing
if request("oldboardmaster")<>Request("boardmaster") then call addmaster(Request("boardmaster"),request("oldboardmaster"),1)

set mrs=Dvbbs.Execute("select max(rootid) from dv_board")
Maxrootid=mrs(0)+1
mrs.close:set mrs=nothing
'���������������̳
'��Ҫ������ԭ������������Ϣ��������ȡ�����ID�������������򡢼̳а���������
'��Ҫ���µ�ǰ����������Ϣ
'�̳а���������Ҫ��д�������и���--ȡ������ǰ̨����boardid in parentstr�����
dim k,nParentStr,mParentStr
dim ParentSql,boardcount
if clng(parentid)<>clng(request("class")) and not (iparentid=0 and cint(request("class"))=0) then
	'���ԭ������һ������ĳ�һ������
	if iparentid>0 and cint(request("class"))=0 then
		'���µ�ǰ��������
		Dvbbs.Execute("update dv_board set depth=0,orders=0,rootid="&maxrootid&",parentid=0,parentstr='0' where boardid="&newboardid)
		ParentStr=ParentStr & ","
		set rs=Dvbbs.Execute("select count(*) from dv_board where ParentStr like '%"&ParentStr&"%'")
		boardcount=rs(0)
		if isnull(boardcount) then
		boardcount=1
		else
		boardcount=boardcount+1
		end if
		'������ԭ��������̳������
		Dvbbs.Execute("update dv_board set child=child-"&boardcount&" where boardid="&iparentid)
		'������ԭ��������̳���ݣ������൱�ڼ�֦�����迼��
		for i=1 to depth
			'�õ��丸��ĸ���İ���ID
			set rs=Dvbbs.Execute("select parentid from dv_board where boardid="&iparentid)
			if not (rs.eof and rs.bof) then
				iparentid=rs(0)
				Dvbbs.Execute("update dv_board set child=child-"&boardcount&" where boardid="&iparentid)
			end if
		next
		if child>0 then
		'������������̳����
		'��������̳�������迼�ǣ�����������̳��Ⱥ�һ������ID(rootid)����
		'���µ�ǰ��������
		'ParentStr=ParentStr & ","		
		i=0
		set rs=Dvbbs.Execute("select * from dv_board where ParentStr like '%"&ParentStr&"%'")
		do while not rs.eof
		i=i+1
		mParentStr=replace(rs("ParentStr"),ParentStr,"")
		Dvbbs.Execute("update dv_board set depth=depth-"&depth&",rootid="&maxrootid&",ParentStr='"&mParentStr&"' where boardid="&rs("boardid"))
		rs.movenext
		loop
		end if
	elseif iparentid>0 and cint(request("class"))>0 then
	'��һ������̳�ƶ�����������̳��
	'�����ָ������̳�������Ϣ
	set trs=Dvbbs.Execute("select * from dv_board where boardid="&request("class"))
	'�õ�������������
	ParentStr=ParentStr & ","
	set rs=Dvbbs.Execute("select count(*) from dv_board where ParentStr like '%"&ParentStr & newboardid&"%'")
	boardcount=rs(0)
	if isnull(boardcount) then boardcount=1
	'�ڻ���ƶ������İ����������������ָ����̳֮�����̳��������
	Dvbbs.Execute("update dv_board set orders=orders + "&boardCount&" + 1 where rootid="&trs("rootid")&" and orders>"&trs("orders")&"")
	'���µ�ǰ��������
	If trs("parentstr")="0" Then
	Dvbbs.Execute("update dv_board set depth="&trs("depth")&"+1,orders="&trs("orders")&"+1,rootid="&trs("rootid")&",ParentID="&request("class")&",ParentStr='" & trs("boardid") & "' where boardid="&newboardid)
	Else
	Dvbbs.Execute("update dv_board set depth="&trs("depth")&"+1,orders="&trs("orders")&"+1,rootid="&trs("rootid")&",ParentID="&request("class")&",ParentStr='" & trs("parentstr") & "," & trs("boardid") & "' where boardid="&newboardid)
	End If
	i=1
	'����������������������
	'���Ϊԭ����ȼ��ϵ�ǰ������̳�����
	set rs=Dvbbs.Execute("select * from dv_board where ParentStr like '%"&ParentStr & newboardid&"%' order by orders")
	do while not rs.eof
	i=i+1
	If trs("parentstr")="0" Then
	iParentStr=trs("boardid") & "," & replace(rs("parentstr"),ParentStr,"")
	Else
	iParentStr=trs("parentstr") & "," & trs("boardid") & "," & replace(rs("parentstr"),ParentStr,"")
	End If
	Dvbbs.Execute("update dv_board set depth=depth+"&trs("depth")&"-"&depth&"+1,orders="&trs("orders")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&iParentStr&"' where boardid="&rs("boardid"))
	rs.movenext
	loop
	ParentID=request("class")
	if rootid=trs("rootid") then
	'��ͬһ�������ƶ�
	'������ָ����ϼ���̳��������iΪ�����ƶ������İ�����
	'�����丸�������
	Dvbbs.Execute("update dv_board set child=child+"&i&" where (not ParentID=0) and boardid="&parentid)
	for k=1 to trs("depth")
		'�õ��丸��ĸ���İ���ID
		set rs=Dvbbs.Execute("select parentid from dv_board where (not ParentID=0) and boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			'�����丸��ĸ��������
			Dvbbs.Execute("update dv_board set child=child+"&i&" where (not ParentID=0) and  boardid="&parentid)
		end if
	next
	'������ԭ���������
	Dvbbs.Execute("update dv_board set child=child-"&i&" where (not ParentID=0) and boardid="&iparentid)
	'������ԭ��������̳����
	'response.write iparentid & "<br>"
	for k=1 to depth
		'�õ���ԭ����ĸ���İ���ID
		set rs=Dvbbs.Execute("select parentid from dv_board where (not ParentID=0) and boardid="&iparentid)
		if not (rs.eof and rs.bof) then
			iparentid=rs(0)
			'response.write iparentid & "<br>"
			'������ԭ����ĸ��������
			Dvbbs.Execute("update dv_board set child=child-"&i&" where (not ParentID=0) and  boardid="&iparentid)
		end if
	next
	else
	'������ָ����ϼ���̳��������iΪ�����ƶ������İ�����
	'�����丸�������
	Dvbbs.Execute("update dv_board set child=child+"&i&" where boardid="&parentid)
	for k=1 to trs("depth")
		'�õ��丸��ĸ���İ���ID
		set rs=Dvbbs.Execute("select parentid from dv_board where boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			'�����丸��ĸ��������
			Dvbbs.Execute("update dv_board set child=child+"&i&" where boardid="&parentid)
		end if
	next
	'������ԭ���������
	Dvbbs.Execute("update dv_board set child=child-"&i&" where boardid="&iparentid)
	'������ԭ��������̳����
	for k=1 to depth
		'�õ���ԭ����ĸ���İ���ID
		set rs=Dvbbs.Execute("select parentid from dv_board where boardid="&iparentid)
		if not (rs.eof and rs.bof) then
			iparentid=rs(0)
			'������ԭ����ĸ��������
			Dvbbs.Execute("update dv_board set child=child-"&i&" where boardid="&iparentid)
		end if
	next
	end if 'end if rootid=trs("rootid") then
	else
	'���ԭ����һ����̳�ĳ�������̳��������̳
	'�õ���ָ������̳�������Ϣ
	set trs=Dvbbs.Execute("select * from dv_board where boardid="&request("class"))
	set rs=Dvbbs.Execute("select count(*) from dv_board where rootid="&rootid)
	boardcount=rs(0)
	Rs.Close
	'������ָ����ϼ���̳��������iΪ�����ƶ������İ�����
	ParentID=request("class")
	'�����丸�������
	Dvbbs.Execute("update dv_board set child=child+"&boardcount&" where boardid="&parentid)
	'response.write parentid & "-"&boardcount&"<br>"
	for k=1 to trs("depth")
		'�õ��丸��ĸ���İ���ID
		set rs=Dvbbs.Execute("select parentid from dv_board where boardid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			'�����丸��ĸ��������
			Dvbbs.Execute("update dv_board set child=child+"&boardcount&" where boardid="&parentid)
		end if
		Rs.Close
	'response.write parentid & "-"&boardcount&"<br>"
	next
	'�ڻ���ƶ������İ����������������ָ����̳֮�����̳��������
	Dvbbs.Execute("update dv_board set orders=orders + "&boardCount&" + 1 where rootid="&trs("rootid")&" and orders>"&trs("orders")&"")
	i=0
	SQL = "select boardid,parentid,ParentStr from dv_board where rootid="&rootid&" order by orders"
	SET Rs = Dvbbs.Execute(SQL)
	If Not Rs.eof Then
		SQL=Rs.GetRows(-1)
		Rs.close:Set Rs = Nothing
		For iii=0 To Ubound(SQL,2)
		i=i+1
		if SQL(1,iii)=0 then
			if trs("ParentStr")="0" then
			parentstr=trs("boardid")
			else
			parentstr=trs("parentstr") & "," & trs("boardid")
			end if
			Dvbbs.Execute("update dv_board set depth=depth+"&trs("depth")&"+1,orders="&trs("orders")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&ParentStr&"',parentid="&request("class")&" where boardid="&SQL(0,iii))
		else
			if trs("ParentStr")="0" then
			parentstr=trs("boardid") & "," & SQL(2,iii)
			else
			parentstr=trs("parentstr") & "," & trs("boardid") & "," & SQL(2,iii)
			end if
			Dvbbs.Execute("update dv_board set depth=depth+"&trs("depth")&"+1,orders="&trs("orders")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&ParentStr&"' where boardid="&SQL(0,iii))
		end if
		Next
	Else
		Rs.close:Set Rs = Nothing
	End If
	End If
End If
dv_suc("��̳�޸ĳɹ���<br>" & str)

set trs=nothing
'cache��������
Dvbbs.ReloadAllBoardInfo()
Dvbbs.Name="setup"
Dvbbs.ReloadSetup
Dvbbs.CacheData=Dvbbs.value

Dim iUpdateRootID
'ԭ����RootID
iUpdateRootID=RootID
'���ڵ�RootID
Set tRs=Dvbbs.Execute("Select RootID From Dv_Board Where BoardID="&request("editid"))
Dim UpdateRootID
UpdateRootID = tRs(0)
Set tRs=Dvbbs.Execute("Select BoardID From Dv_Board Where RootID="&UpdateRootID&" Order By Orders")
Do While Not tRs.Eof
	Dvbbs.ReloadBoardInfo(tRs(0))
tRs.Movenext
Loop
'����༭���������һ�����࣬��RootID��һ���������ԭRootID��������̳����
If iUpdateRootID<>UpdateRootID Then
	Set tRs=Dvbbs.Execute("Select BoardID From Dv_Board Where RootID="&iUpdateRootID&" Order By Orders")
	Do While Not tRs.Eof
		Dvbbs.ReloadBoardInfo(tRs(0))
	tRs.Movenext
	Loop
End If
Set tRs=Nothing

Dvbbs.DelCahe "BoardJumpList"
Dvbbs.DelCahe "MyAllBoardList"
'end cache
end sub

'ɾ�����棬ɾ���������ӣ���ڣ�����ID
sub del()
Dim trs
'�������ϼ�������̳�����������̳�����¼���̳������ɾ��
Set tRs=Dvbbs.Execute("Select RootID From Dv_Board Where BoardID="&request("editid"))
Dim UpdateRootID
UpdateRootID = tRs(0)
set rs=Dvbbs.Execute("select ParentStr,child,depth from dv_board where boardid="&Request("editid"))
if not (rs.eof and rs.bof) then
if rs(1)>0 then
	response.write "����̳����������̳����ɾ����������̳���ٽ���ɾ������̳�Ĳ���"
	exit sub
end if
'������ϼ����棬���������
if rs(2)>0 then
	Dvbbs.Execute("update dv_board set child=child-1 where boardid in ("&rs(0)&")")
	Dim UpdateBoardID
	UpdateBoardID=Split(Rs(0),",")
	For i=0 To Ubound(UpdateBoardID)
		Dvbbs.ReloadBoardInfo(UpdateBoardID(i))
	Next
end if
sql = "delete from dv_board where boardid="&Request("editid")
Dvbbs.Execute(sql)
for i=0 to ubound(AllPostTable)
sql = "delete from "&AllPostTable(i)&" where boardid="&Request("editid")
Dvbbs.Execute(sql)
next
Dvbbs.Execute("delete from dv_topic where boardid="&Request("editid"))
Dvbbs.Execute("delete from dv_besttopic where boardid="&Request("editid"))
Dvbbs.Execute("delete from dv_upfile where f_boardid="&Request("editid"))
end if
set rs=nothing
Dvbbs.ReloadAllBoardInfo()
Dvbbs.Name="setup"
Dvbbs.ReloadSetup
Dvbbs.CacheData=Dvbbs.value

Set tRs=dvbbs.Execute("Select BoardID From Dv_Board Where RootID="&UpdateRootID&" Order By Orders")
Do While Not tRs.Eof
	Dvbbs.ReloadBoardInfo(tRs(0))
tRs.Movenext
Loop
Set tRs=Nothing
'Dvbbs.ReloadBoardInfo(request("editid"))

Dvbbs.DelCahe "BoardJumpList"
Dvbbs.DelCahe "MyAllBoardList"
response.write "<p>��̳ɾ���ɹ���"
end sub

sub orders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">��̳һ���������������޸�(������Ӧ��̳������������������Ӧ���������)
	</th>
	</tr>

	<tr>
	<td class="Forumrow"><table width="50%">
<%
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from dv_Board where ParentID=0 order by RootID"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "��û����Ӧ����̳���ࡣ"
	else
		do while not rs.eof
		response.write "<form action=admin_board.asp?action=updatorders method=post><tr><td width=""50%"">"&rs("boardtype")&"</td>"
		response.write "<td width=""50%""><input type=text name=""OrderID"" size=4 value="""&rs("rootid")&"""><input type=hidden name=""cID"" value="""&rs("rootid")&""">&nbsp;&nbsp;<input type=submit name=Submit value=�޸�></td></tr></form>"
		rs.movenext
		loop
%>
</table>
<BR>&nbsp;<font color=red>��ע�⣬����һ��<B>������д��ͬ�����</B>������ǳ����޸���</font>
<%
	end if
	rs.close
	set rs=nothing
%>
	</td>
	</tr>
</table>
<%
end sub

sub updateorders()
	dim cID,OrderID,ClassName
	'response.write request.form("cID")(1)
	'response.end
	cID=replace(request.form("cID"),"'","")
	OrderID=replace(request.form("OrderID"),"'","")
	set rs=Dvbbs.Execute("select boardid from dv_board where rootid="&orderid)
	if rs.eof and rs.bof Then
		Dv_suc("����")
	Dvbbs.Execute("update dv_board set rootid="&OrderID&" where rootid="&cID)
	else
	response.write "�벻Ҫ��������̳������ͬ�����"
	end if

	Dvbbs.ReloadAllBoardInfo()
	Dim Forum_Boards
	Dvbbs.Name="setup"
	Dvbbs.ReloadSetup
	Dvbbs.CacheData=Dvbbs.value
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
end sub

Sub Boardorders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">��̳N���������������޸�(������Ӧ��̳������������������Ӧ���������)
	</th>
	</tr>
	<tr>
	<td class="Forumrow"><table width="90%">
<%
	Dim Trs, Uporders, Doorders
	Set Rs = Server.CreateObject ("Adodb.recordset")
	Sql = "SELECT Depth, Child, Parentid, Boardtype, Orders, BoardId FROM Dv_Board ORDER BY RootID, Orders"
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof And Rs.Bof Then
		Response.Write "��û����Ӧ����̳���ࡣ"
	Else
		Sql = Rs.GetRows(-1)
		Dim Bn
		Rs.Close:Set Rs = Nothing
		For Bn = 0 To Ubound(Sql,2)
			Response.Write "<form action=admin_board.asp?action=updatboardorders method=post><tr><td width=""50%"">"
			If Sql(0,Bn) > 0 Then
				For i = 1 To Sql(0,Bn)
					Response.Write "&nbsp;"
				Next
			End If
			If Sql(1,Bn) > 0 Then
				Response.Write "<img src=skins/default/plus.gif>"
			Else
				Response.Write "<img src=skins/default/nofollow.gif>"
			End If
			If Sql(2,Bn) = 0 Then
				Response.Write "<b>"
			End If
			Response.Write Sql(3,Bn)
			If Sql(1,Bn) > 0 Then
				Response.Write "(" & Sql(1,Bn) & ")"
			End If
			Response.Write "</td><td width=""50%"">"
			If Sql(2,Bn) > 0 Then
				'�����ͬ��ȵİ�����Ŀ���õ��ð�������ͬ��ȵİ���������λ�ã�֮�ϻ���֮�µİ�������
				'��������������ӦΪFor i=1 to �ð�֮�ϵİ�����
				Set Trs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_Board WHERE ParentID = " & Sql(2,Bn) & " AND ORDERS < " & Sql(4,Bn) &"")
				Uporders = Trs(0)
				If Isnull(Uporders) Then Uporders = 0
				'���ܽ���������ӦΪFor i=1 to �ð�֮�µİ�����
				Set Trs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_Board WHERE ParentID = " & Sql(2,Bn) &" AND ORDERS > " & Sql(4,Bn) &"")
				Doorders = Trs(0)
				If Isnull(doorders) Then Doorders = 0
				If Uporders > 0 Then
					Response.Write "<select name=uporders size=1><option value=0>�����ƶ�</option>"
					For i = 1 To Uporders
						Response.Write "<option value=" & i & ">" & i & "</option>"
					Next
					Response.Write "</select>"
				End If
				If Doorders > 0 Then
					If uporders > 0 Then Response.Write "&nbsp;"
					Response.Write "<select name=doorders size=1><option value=0>�����ƶ�</option>"
					For i = 1 To Doorders
						Response.Write "<option value=" & i & ">" & i & "</option>"
					Next
					Response.Write "</select>"
				End If
				If Doorders > 0 Or Uporders > 0 Then
					Response.Write "<input type=hidden name=""editID"" value=""" & Sql(5,Bn) & """>&nbsp;<input type=submit name=Submit value=�޸�>"
				End If
			End If
			Response.Write "</td></tr></form>"
			Uporders = 0
			Doorders = 0
		Next
		Response.Write "</table>"
	End If
%>
	</td>
	</tr>
</table>
<%
End Sub

Sub Updateboardorders()
	Dim ParentID, Orders, ParentStr, Child
	Dim Uporders, Doorders, Oldorders, Trs, ii
	Dim Bn
	If Not Isnumeric(Request("EditID")) Then
		Response.Write "�Ƿ��Ĳ�����"
		Exit Sub
	End If
	If Request("Uporders") <> "" And Not Cint(Request("Uporders")) = 0 Then
		If Not Isnumeric(Request("Uporders")) Then
			Response.Write "�Ƿ��Ĳ�����"
			Exit Sub
		Elseif Cint(Request("Uporders")) = 0 Then
			Response.Write "��ѡ��Ҫ���������֣�"
			Exit Sub
		End If
		'�����ƶ�
		'Ҫ�ƶ�����̳��Ϣ shinzeal����rootid��depth��Ϊ����������ذ��������
		Dim Rootid, Depth
		Set Rs = Dvbbs.Execute("SELECT ParentID, Orders, ParentStr, Child, Rootid, Depth FROM Dv_Board WHERE Boardid = " & Request("EditID"))
		ParentID = Rs(0)
		Orders = Rs(1)
		ParentStr = Rs(2) & "," & Request("EditID")
		Child = Rs(3)
		Rootid = Rs(4)
		Depth = Rs(5)
		i = 0
		Set Rs = Nothing
		If Child > 0 Then
			Set Rs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_Board WHERE ParentStr LIKE '%" & ParentStr & "%' AND Rootid = " & Rootid)
			Oldorders = Rs(0)
		Else
			Oldorders = 0
		End If
		'shinzeal�������shin��shinlast��¼���º��������orders
		Dim Shin, Shinlast
		Shin = 0
		Shinlast = 0
		'�͸���̳ͬ������������֮�ϵ���̳��������������ĩ��Ϊ��ǰ��̳�����
		Set Rs = Nothing
		Set Rs = Dvbbs.Execute("SELECT Boardid, Orders, Child, ParentStr FROM Dv_Board WHERE ParentID = " & ParentID & " AND Orders < " & Orders & " ORDER BY Orders DESC")
		If Not(Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
			Rs.Close:Set Rs = Nothing
			For Bn = 0 To Ubound(Sql,2)
				i = i + 1
				If Cint(Request("Uporders")) >= i Then
					'response.write "UPDATE Dv_Board SET Orders = " & orders & " WHERE Boardid = " & Sql(0,Bn) & "<br>"
					If Sql(2,Bn) > 0 Then
						ii = 0
						Set Trs = Dvbbs.Execute("SELECT Boardid, Orders FROM Dv_Board WHERE ParentStr like '%" & Sql(3,Bn) & "," & Sql(0,Bn) & "%' ORDER BY Orders")
						If Not (Trs.Eof And Trs.Bof) Then
							Do While Not Trs.Eof
								ii = ii + 1
								Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & Oldorders & " + " & ii & " WHERE Boardid = " & Trs(0))
								Shin = Orders + Oldorders + ii
								If Shin > Shinlast Then Shinlast = Shin
								Trs.Movenext
							Loop
							Trs.Close:Set Trs = Nothing
						End If
					End If
					Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & Oldorders & " WHERE Boardid = " & Sql(0,Bn))
					Shin = Orders + Oldorders
					If Shin > Shinlast Then Shinlast = Shin
					If Cint(Request("Uporders")) = i Then Uporders = Sql(1,Bn)
				End If
				Orders = Sql(1,Bn)
			Next
		End If
		'response.write "update dv_board set orders="&uporders&" where boardid="&request("editID")
		'������Ҫ�������̳�����
		Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Uporders & " WHERE Boardid = " & Request("EditID"))
		'�����������̳���������������̳����
		If Child > 0 Then
			i = Uporders
			Set Rs = Dvbbs.Execute("SELECT Boardid FROM Dv_Board WHERE ParentStr Like '%" & ParentStr & "%' AND Depth > " & Depth & " ORDER BY ORDERS")
			If Not(Rs.Eof And Rs.Bof) Then
				Sql = Rs.GetRows(-1)
				Rs.Close:Set Rs = Nothing
				For Bn = 0 To Ubound(Sql,2)
					i = i + 1
					Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & i & " WHERE Boardid = " & Sql(0,Bn))
					Shin = i
					If Shin > Shinlast Then Shinlast = Shin
				Next
			End If
		End If
		'shinzeal������±�������̳֮������ͬ����̳��orders������ͱ�������̳���¼���̳�ظ�
		Dim Shin1, Shinlast1
		Shin1 = 0
		Shinlast1 = 0
		Set Rs = Dvbbs.Execute("SELECT Boardid, Orders, Child, ParentStr FROM Dv_Board WHERE ParentID = " & ParentID & " AND Orders > " & Uporders & " ORDER BY Orders")
		If Not(Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
			Rs.Close:Set Rs = Nothing
			For Bn = 0 To Ubound(Sql,2)
				If Sql(2,Bn) > 0 Then
					ii = 0
					Set Trs = Dvbbs.Execute("SELECT Boardid, Orders FROM Dv_Board WHERE ParentStr LIKE '%" & Sql(3,Bn) & ", " & Sql(0,Bn) & "%' ORDER By Orders")
					If Not (Trs.Eof And Trs.Bof) Then
						Do While Not Trs.Eof
							ii = ii + 1
							'response.write "update dv_board set orders="&orders&"+"&ii&" where boardid="&trs(0)&"��a<br>"
							Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & ii & " + " & Shinlast & " WHERE Boardid = " & Trs(0))
							Shin1 = Orders + Oldorders + ii + Shinlast
							If Shin1 > Shinlast1 Then Shinlast1 = Shin1
							Trs.Movenext
						Loop
						Trs.Close:Set Trs = Nothing
					End If
				End If
				'response.write "update dv_board set orders="&orders&" where boardid="&rs(0)&"<br>"
				Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & Shinlast & " WHERE Boardid = " & Sql(0,Bn))
				Shin1 = Orders + Oldorders + Shinlast
				If Shin1 > Shinlast1 Then Shinlast1 = Shin1
				Orders = Sql(1,Bn)
			Next
		End If
		'shinzeal������±�������̳��һ����orders�ڱ�������̳֮�����̳orders����ֹorders���ཻ��
		Set Rs = Dvbbs.Execute("SELECT Boardid, Orders, Child, ParentStr FROM Dv_Board WHERE RootID = " & RootID & " AND Orders> " & Uporders & " AND Depth < " & Depth & " ORDER BY Orders")
		If Not (Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
			Rs.Close:Set Rs = Nothing
			For Bn = 0 To Ubound(Sql,2)
				If Sql(2,Bn) > 0 Then
					ii = 0
					Set Trs = Dvbbs.Execute("SELECT Boardid, Orders FROM Dv_Board WHERE ParentStr LIKE '%" & Sql(3,Bn) & ", " & Sql(0,Bn) & "%' ORDER BY Orders")
					If Not (Trs.Eof And Trs.Bof) Then
						Do While Not Trs.Eof
							ii = ii + 1
							'response.write "update dv_board set orders="&orders&"+"&ii&" where boardid="&trs(0)&"��a<br>"
							Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & ii & " + " & Shinlast1 & " WHERE Boardid = " & Trs(0))
							Trs.Movenext
						Loop
						Trs.Close:Set Trs = Nothing
					End If
				End If
				'response.write "update dv_board set orders="&orders&" where boardid=" & Sql(0,Bn) & "<br>"
				Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & Shinlast1 & " WHERE Boardid = " & Sql(0,Bn))
				Orders = Sql(1,Bn)
			Next
		End If
		'shinzeal��������̳�ĸ��½���
	Elseif Request("Doorders") <> "" Then
		If Not Isnumeric(Request("Doorders")) Then
			Response.Write "�Ƿ��Ĳ�����"
			Exit Sub
		Elseif Cint(Request("doorders")) = 0 Then
			Response.Write "��ѡ��Ҫ�½������֣�"
			Exit Sub
		End If
		'Ҫ�ƶ�����̳��Ϣ shinzeal����rootid��depth��Ϊ����������ذ��������
		Set Rs = Dvbbs.Execute("SELECT ParentID, Orders, ParentStr, Child, Rootid, Depth FROM Dv_Board WHERE Boardid = " & Request("EditID"))
		ParentID = Rs(0)
		Orders = Rs(1)
		ParentStr = Rs(2) & "," & Request("EditID")
		Child = Rs(3)
		Rootid = Rs(4)
		Depth = Rs(5)
		i = 0
		Shin = 0
		Shinlast = 0
		Set Rs = Nothing
		Set Rs = Dvbbs.Execute("SELECT Boardid, Orders, Child, ParentStr FROM Dv_Board WHERE ParentID = " & ParentID & " AND Orders > " & Orders & " ORDER BY Orders")
		If Not (Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
			Rs.Close:Set Rs = Nothing
			For Bn = 0 To Ubound(Sql,2)
				i = i + 1
				If Cint(Request("Doorders")) >= i Then
					If Sql(2,Bn) > 0 Then
						ii = 0
						Set Trs = Dvbbs.Execute("SELECT Boardid, Orders FROM Dv_Board WHERE ParentStr LIKE '%" & Sql(3,Bn) & " , " & Sql(0,Bn) & "%' ORDER BY Orders")
						If Not (Trs.Eof And Trs.Bof) Then
							Do While Not Trs.Eof
								ii = ii + 1
								'response.write "update dv_board set orders="&orders&"+"&ii&" where boardid="&trs(0)&"��a<br>"
								Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & ii & " WHERE Boardid = " & Trs(0))
								Shin = Orders + ii
								If Shin > Shinlast Then Shinlast = Shin
								Trs.Movenext
							Loop
							Trs.Close:Set Trs = Nothing
						End If
					End If
					'response.write "update dv_board set orders="&orders&" where boardid=" & Sql(0,Bn) & "<br>"
					Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " WHERE Boardid = " & Sql(0,Bn))
					Shin = Orders
					If Shin > Shinlast Then Shinlast = Shin
					If Cint(Request("doorders")) = i Then Doorders = Sql(1,Bn)
				End If
				Orders = Sql(1,Bn)
			Next
		End If
		'response.write "update dv_board set orders="&doorders&" where boardid="&request("editID")&"<br>"
		Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Shinlast + 1 & " WHERE Boardid = " & Request("EditID"))
		'�����������̳���������������̳����
		If Child > 0 Then
			i = Shinlast + 1
			Set Rs = Dvbbs.Execute("SELECT Boardid FROM Dv_Board WHERE ParentStr LIKE '%" & ParentStr & "%' AND Depth > " & Depth & " ORDER BY Orders")
			If Not (Rs.Eof And Rs.Bof) Then
				Sql = Rs.GetRows(-1)
				Rs.Close:Set Rs = Nothing
				For Bn = 0 To Ubound(Sql,2)
					i = i + 1
					'response.write "update dv_board set orders="&i&" where boardid=" & Sql(0,Bn) & "��b<br>"
					Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & i & " WHERE Boardid = " & Sql(0,Bn))
					shin = i
					If Shin > Shinlast Then Shinlast = Shin
				Next
			End If
		End If
		'shinzeal������±��½���̳֮������ͬ����̳��orders������ͱ�������̳���¼���̳�ظ�
		Shin1 = 0
		Shinlast1 = 0
		Set Rs = Dvbbs.Execute("SELECT Boardid, Orders, Child, ParentStr FROM Dv_Board WHERE ParentID = " & ParentID & " AND Orders > " & Doorders & " ORDER BY Orders")
		If Not (Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
			Rs.Close:Set Rs = Nothing
			For Bn = 0 To Ubound(Sql,2)
				Orders = Sql(1,Bn)
				If Sql(2,Bn) > 0 Then
					ii = 0
					Set Trs = Dvbbs.Execute("SELECT Boardid, Orders FROM Dv_board WHERE ParentStr LIKE '%" & Sql(3,Bn) & "," & Sql(0,Bn) & "%' ORDER BY Orders")
					If Not (Trs.Eof And Trs.Bof) Then
						Do While Not Trs.Eof
							ii = ii + 1
							'response.write "update dv_board set orders="&orders&"+"&ii&" where boardid="&trs(0)&"��a<br>"
							Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & ii & " + " & Shinlast & " WHERE Boardid = " & Trs(0))
							Shin1 = Orders + ii + Shinlast
							If Shin1 > Shinlast1 Then Shinlast1 = Shin1
							Trs.Movenext
						Loop
						Trs.Close:Set Trs = Nothing
					End If
				End If
				'response.write "update dv_board set orders="&orders&" where boardid=" & Sql(0,Bn) & "<br>"
				Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & Shinlast & " WHERE Boardid = " & Sql(0,Bn))
				Shin1 = Orders + Shinlast
				If Shin1 > Shinlast1 Then Shinlast1 = Shin1
			Next
		End If
		'shinzeal������±��½���̳��һ����orders�ڱ�������̳֮�����̳orders����ֹorders���ཻ��
		Set Rs = Dvbbs.Execute("SELECT BoardId, Orders, Child, ParentStr FROM Dv_Board WHERE RootID = " & RootID & " AND Orders > " & Doorders & " AND Depth < " & Depth & " ORDER BY Orders")
		If Not (Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
			Rs.Close:Set Rs = Nothing
			For Bn = 0 To Ubound(Sql,2)
				Orders = Sql(1,Bn)
				If Sql(2,Bn) > 0 Then
					ii = 0
					Set Trs = Dvbbs.Execute("SELECT Boardid, Orders FROM Dv_Board WHERE ParentStr LIKE '%" & Sql(3,Bn) & "," & Sql(0,Bn) & "%' Order BY Orders")
					If Not (Trs.Eof And Trs.Bof) Then
						Do While Not Trs.Eof
							ii = ii + 1
							'response.write "update dv_board set orders="&orders&"+"&ii&" where boardid="&trs(0)&"��a<br>"
							Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & ii & " + " & Shinlast1 & " WHERE Boardid = " & Trs(0))
							Trs.Movenext
						Loop
						Trs.Close:Set Trs = Nothing
					End If
				End If
				'response.write "update dv_board set orders="&orders&" where boardid=" & Sql(0,Bn) & "<br>"
				Dvbbs.Execute("UPDATE Dv_Board SET Orders = " & Orders & " + " & Shinlast1 & " WHERE Boardid = " & Sql(0,Bn))
			Next
		End If
		'shinzeal���½���̳�ĸ��½���
	End If
	Dvbbs.ReloadAllBoardInfo()
	Dvbbs.Name = "Setup"
	Dvbbs.ReloadSetup
	Dvbbs.CacheData = Dvbbs.Value
	Set tRs = Dvbbs.Execute("SELECT RootID FROM Dv_Board WHERE BoardID = " & Request("Editid"))
	Dim UpdateRootID
	UpdateRootID = tRs(0)
	Set tRs = Dvbbs.Execute("SELECT BoardID FROM Dv_Board WHERE RootID = " & UpdateRootID & " ORDER BY Orders")
	Do While Not tRs.Eof
		Dvbbs.ReloadBoardInfo(tRs(0))
		tRs.Movenext
	Loop
	Set tRs = Nothing
	'Dvbbs.ReloadBoardInfo(request("editid"))
	Dvbbs.DelCahe "BoardJumpList"
	Dvbbs.DelCahe "MyAllBoardList"
	Response.Redirect "admin_board.asp?action=boardorders"
End Sub

Sub Addmaster(s,o,n)
	Dim Arr, Pw, Oarr
	Dim Classname, Titlepic
	Set Rs = Dvbbs.Execute("SELECT Usertitle, GroupPic FROM Dv_Usergroups WHERE Usergroupid = 3 ORDER BY Minarticle DESC")
	If Not (Rs.Eof And Rs.Bof) Then
		Classname = Rs(0)
		Titlepic = Rs(1)
	End If
	Randomize
	Pw = Cint(Rnd * 9000) + 1000
	Arr = Split(s,"|")
	Oarr = Split(o,"|")
	Set Rs = Server.Createobject("Adodb.Recordset")
	For i = 0 To Ubound(Arr)
		Sql = "SELECT * FROM [Dv_User] WHERE Username = '" & Arr(i) & "'"
		Rs.Open Sql,Conn,1,3
		If Rs.Eof And Rs.Bof Then
			Rs.Addnew
			Rs("Username") = Arr(i)
			Rs("Userpassword") = Md5(Pw,16)
			Rs("Userclass") = Classname
			Rs("UserGroupID") = 3
			Rs("Titlepic") = Titlepic
			Rs("UserWealth") = 100
			Rs("Userep") = 30
			Rs("Usercp") = 30
			Rs("Userisbest") = 0
			Rs("Userdel") = 0
			Rs("Userpower") = 0
			Rs("Lockuser") = 0
			'�������ϸ����ʹ��¼����ʾ���ϲ������
			Rs("UserSex") = 1
			Rs("UserEmail") = Arr(i) & "@aspsky.net"
			Rs("UserFace") = "Images/userface/image1.gif"
			Rs("UserWidth") = 32
			Rs("UserHeight") = 32
			Rs("UserIM") = "||||||||||||||||||"
			Rs("UserFav") = "İ����,�ҵĺ���,������"
			Rs("LastLogin") = Now()
			Rs("JoinDate") = Now()
			Rs("Userpost") = 0
			Rs("Usertopic") = 0
			Rs.Update
			Str = Str & "������������û���<b>" & Arr(i) & "</b> ���룺<b>" & Pw & "</b><br><br>"
			Dvbbs.Execute("UPDATE Dv_Setup SET Forum_Usernum = Forum_Usernum + 1, Forum_Lastuser = '" & Arr(i) & "'")
		Else
			If Rs("UserGroupID") = 4 Then
				Rs("Userclass") = Classname
				Rs("UserGroupID") = 3
				Rs("Titlepic") = Titlepic
				Rs.Update
			End If
		End If
		Rs.Close
	Next
	Dvbbs.Name = "Setup"
	Dvbbs.ReloadSetup
	'�ж�ԭ���������������Ƿ񻹵��ΰ�������û�е����򳷻����û�ְλ��
	If n = 1 Then
		Dim Iboardmaster
		Dim UserGrade, Article
		Iboardmaster = False
		For i = 0 To Ubound(Oarr)
			Set Rs = Dvbbs.Execute("SELECT Boardmaster FROM Dv_Board")
			Do While Not Rs.Eof
				If Instr("|" & Trim(Rs("Boardmaster")) & "|","|" & Trim(Oarr(i)) & "|") > 0 Then
					Iboardmaster = True
					Exit Do
				End If
				Rs.Movenext
			Loop
			If Not Iboardmaster Then
				Set Rs = Dvbbs.Execute("SELECT Userid, UserGroupID, UserPost FROM [Dv_User] WHERE Username = '" & Trim(Oarr(i)) & "'")
				If Not (Rs.Eof And Rs.Bof) Then
					If Rs(1) > 2 Then
						If Not Isnumeric(Rs(2)) Or Rs(2) = "" Then
							Article = 0
						Else
							Article = Cstr(Rs(2))
						End If
						'ȡ��Ӧע���Ա�ĵȼ�
						Set UserGrade = Dvbbs.Execute("SELECT TOP 1 Usertitle, Grouppic FROM Dv_Usergroups WHERE Minarticle <= " & Article & " AND NOT MinArticle = -1 AND ParentGID = 4 ORDER BY MinArticle DESC")
						If Not (UserGrade.Eof And UserGrade.Bof) Then
							Dvbbs.Execute("UPDATE [Dv_User] SET UserGroupID = 4, Titlepic = '" & UserGrade(1) & "', Userclass = '" & UserGrade(0) & "' WHERE Userid = " & Rs(0))
						End If
						UserGrade.Close:Set UserGrade = Nothing
					End If
				End If
			End If
			Iboardmaster = False
		Next
	End If
	Set Rs = Nothing
End Sub

Rem �ְ����û�Ȩ������ ��д2004-5-2 Dvbbs.YangZheng
Sub BoardPerMission()
	Dim iUserGroupID(20), UserTitle(20)
	Dim Trs, Ars, k, ii
	Dim Bn
	Set Trs = Dvbbs.Execute("SELECT Usertitle, Usergroupid FROM Dv_UserGroups WHERE Issetting = 1 ORDER BY UserGroupId")
	If Not (Trs.Eof And Trs.Bof) Then
		Sql = Trs.GetRows(-1)
		Trs.Close:Set Trs = Nothing
		For ii = 0 To Ubound(Sql,2)
			UserTitle(ii) = Sql(0,ii)
			iUserGroupID(ii) = Sql(1,ii)
		Next
	End If
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr>
	<th height="25">�༭��̳Ȩ��</th>
	</tr>
	<tr>
	<td class=forumrow>�����������ò�ͬ�û����ڲ�ͬ��̳�ڵ�Ȩ�ޣ���ɫ��ʾΪ����̳���û���ʹ�õ����û���������<BR>�ڸ�Ȩ�޲��ܼ̳У�������������һ�������¼���̳�İ��棬��ôֻ�������õİ�����Ч��������������̳��Ч<BR>���������������Ч������������ҳ��<B>ѡ���Զ�������</B>��ѡ�����Զ������ú��������õ�Ȩ�޽�<B>����</B>���û������ã������û���Ĭ�ϲ��ܹ������ӣ������������˸��û���ɹ������ӣ���ô���û������������Ϳ��Թ�������
	</td>
	</tr>
</table><BR>
<table width="95%" cellspacing="1" cellpadding="1" align=center class="tableBorder">
<tr> 
<th width="35%" class="tableHeaderText" height=25>��̳����
</th>
<th width="35%" class="tableHeaderText" height=25>�����û���Ȩ��
</th>
</tr>
<%
	Sql = "SELECT Depth, Child, Parentid, BoardType, Boardid FROM Dv_Board ORDER BY Rootid, Orders"
	Set Rs = Dvbbs.Execute(Sql)
	If Not (Rs.Eof And Rs.Bof) Then
		Sql = Rs.GetRows(-1)
		Rs.Close:Set Rs = Nothing
		For Bn = 0 To Ubound(Sql,2)
			Response.Write "<tr><td height=25 width=40% class=forumrow>"
			If Sql(0,Bn) > 0 Then
				For i = 1 To Sql(0,Bn)
					Response.Write "&nbsp;"
				Next
			End If
			If Sql(1,Bn) > 0 Then
				Response.Write "<img src=skins/default/plus.gif>"
			Else
				Response.Write "<img src=skins/default/nofollow.gif>"
			End If
			If Sql(2,Bn) = 0 Then
				Response.Write "<b>"
			End If
			Response.Write Sql(3,Bn)
			If Sql(1,Bn) > 0 Then
				Response.Write "(" & Sql(1,Bn) & ")"
			End If
%>
</td>
<FORM METHOD=POST ACTION="?action=editpermission">
<td width=60% class="forumrow">&nbsp;
<select name="groupid" size=1>
<%
			For k = 0 To ii-1
				Set Ars = Dvbbs.Execute("SELECT Pid FROM Dv_BoardPerMission WHERE BoardID = " & Sql(4,Bn) & " AND GroupID = " & iUserGroupID(k))
				If Ars.Eof And Ars.Bof Then
					Response.Write "<option value=""" & iUserGroupID(k) & """>" & UserTitle(k) & "</option>"
				Else
					Response.Write "<option value=""" & iUserGroupID(k) & """>" & UserTitle(k) & "(�Զ���)</option>"
				End If
			Next
			Response.Write "</select><input type=hidden value="
			Response.Write Sql(4,Bn)
			Response.Write " name=reboardid><input type=submit name=submit value=����>"
			Dim Percount
			Set Trs = Dvbbs.Execute("SELECT COUNT(*) FROM Dv_BoardPermission WHERE Boardid = " & Sql(4,Bn))
			Percount = Trs(0)
			If Not Isnull(Percount) And Percount > 0 Then Response.Write "(���Զ������)"
			Response.Write "</td></FORM></tr>"
		Next
	End If
	Response.Write "</table><BR><BR>"
	Set Trs = Nothing
	Set Ars = Nothing
End Sub

sub editpermission()
if not isnumeric(request("groupid")) Or request("groupid")="" Or request("reBoardID")="" Or not isnumeric(request("reBoardID"))  then
response.write "����Ĳ������뷵�طְ���Ȩ��������ҳѡ����ȷ�����ã�"
exit sub
end if
if request("groupaction")="yes" then
	dim GroupSetting,rspid
	Dim IsGroupSetting,MyIsGroupSetting
	GroupSetting=GetGroupPermission
	Set rs= Server.CreateObject("ADODB.Recordset")
	if request("isdefault")=1 then
		Dvbbs.Execute("delete from dv_BoardPermission where BoardID="&request("reBoardID")&" and GroupID="&request("GroupID"))
		Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&request("reBoardID"))
		If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
			MyIsGroupSetting = Request("GroupID")
		Else
			IsGroupSetting = "," & Rs(0) & ","
			IsGroupSetting = Replace(IsGroupSetting,"," & Request("GroupID"),"")
			IsGroupSetting = Split(IsGroupSetting,",")
			For i=1 To Ubound(IsGroupSetting)-1
				If i=1 Then
					MyIsGroupSetting = IsGroupSetting(i)
				Else
					MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
				End If
			Next
		End If
		Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&request("reBoardID"))
		Set Rs=Nothing
		Dvbbs.ReloadBoardInfo(request("reBoardID"))
	else
		sql="Select pid from dv_BoardPermission where BoardID="&request("reBoardID")&" And GroupID="&request("GroupID")&""
		Set Rspid=Dvbbs.Execute(sql)
		If Not Rspid.eof And Not Rspid.bof then
			sql="update dv_BoardPermission set PSetting='"&GroupSetting&"' where pid="&Rspid(0)
		else
			sql="insert into dv_BoardPermission (BoardID,GroupID,PSetting) values ("&request("reBoardID")&","&request("GroupID")&",'"&GroupSetting&"')"
		End If 
		Set Rspid=Nothing
		Dvbbs.Execute(sql)
		Set Rs=Dvbbs.Execute("select IsGroupSetting From Dv_Board Where BoardID="&request("reBoardID"))
		If Trim(Rs(0))="" Or IsNull(Rs(0)) Then
			MyIsGroupSetting = Request("GroupID")
		Else
			IsGroupSetting = "," & Rs(0) & ","
			IsGroupSetting = Replace(IsGroupSetting,"," & Request("GroupID"),"")
			IsGroupSetting = IsGroupSetting & Request("GroupID") & ","
			IsGroupSetting = Split(IsGroupSetting,",")
			For i=1 To Ubound(IsGroupSetting)-1
				If i=1 Then
					MyIsGroupSetting = IsGroupSetting(i)
				Else
					MyIsGroupSetting = MyIsGroupSetting & "," & IsGroupSetting(i)
				End If
			Next
		End If
		Dvbbs.Execute("update dv_Board set IsGroupSetting='"&MyIsGroupSetting&"' Where BoardID="&request("reBoardID"))
		Set Rs=Nothing
		Dvbbs.ReloadBoardInfo(request("reBoardID"))
	End If

	Set Rs=Nothing
	Dv_suc("�޸ĳɹ���")
Else
Dim reGroupSetting,reBoardID,groupid
Dim Groupname,Boardname,founduserper
founduserper=false
if request("GroupID")<>"" then
set rs=Dvbbs.Execute("select * from dv_BoardPermission where boardid="&request("reBoardID")&" and GroupID="&request("GroupID"))
if rs.eof and rs.bof then
	founduserper=false
else
groupid=rs("groupid")
reGroupSetting=rs("PSetting")
reBoardID=rs("boardid")
set rs=Dvbbs.Execute("select usertitle from dv_UserGroups where usergroupid="&groupid)
groupname=rs("usertitle")
founduserper=true
end if
if not founduserper then
set rs=Dvbbs.Execute("select * from dv_usergroups where usergroupid="&request("groupid"))
if rs.eof and rs.bof then
response.write "δ�ҵ����û��飡"
exit sub
end if
groupid=request("groupid")
reGroupSetting=rs("GroupSetting")
reBoardID=request("reBoardID")
Groupname=rs("usertitle")
end if
end if
set rs=Dvbbs.Execute("select boardtype from dv_board where boardid="&reBoardID)
Boardname=rs("boardtype")
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<FORM METHOD=POST ACTION="?action=editpermission">
<input type=hidden name="groupid" value="<%=groupid%>">
<input type=hidden name="reBoardID" value="<%=reBoardID%>">
<input type=hidden name="pID" value="<%=request("pid")%>">
<tr> 
<th height="23" colspan="3" >�༭��̳�û���Ȩ��&nbsp;>> <%=boardname%>&nbsp;>> <%=groupname%></th>
</tr>
<tr> 
<td height="23" colspan="3" class=forumrow><input type=radio name="isdefault" value="1" <%if not founduserper then%>checked<%end if%>><B>ʹ���û���Ĭ��ֵ</B> (ע��: �⽫ɾ���κ�֮ǰ�������Զ�������)</td>
</tr>
<tr> 
<td height="23" colspan="3"  class=forumrow><input type=radio name="isdefault" value="0" <%if founduserper then%>checked<%end if%>><B>ʹ���Զ�������</B>&nbsp;(ѡ���Զ������ʹ����������Ч) </td>
</tr>
<%
GroupPermission(reGroupSetting)
%>
<input type=hidden value="yes" name="groupaction">
</FORM>
</table>
<%
end if
end sub

sub RestoreBoard()
'����Ŀǰ������ѭ��i��ֵ����rootid
'��ԭ���а����depth,orders,parentid,parentstr,childΪ0
i=0
set rs=Dvbbs.Execute("select boardid from dv_board order by rootid,orders")
do while not rs.eof
i=i+1
Dvbbs.Execute("update dv_board set rootid="&i&",depth=0,orders=0,ParentID=0,ParentStr='0',child=0 where boardid="&rs(0))
rs.movenext
loop
Set Rs=Nothing
Dv_suc("�뷵������̳�������á���λ")
Dvbbs.ReloadAllBoardInfo()
Dim Forum_Boards
Dvbbs.Name="setup"
Dvbbs.ReloadSetup
Dvbbs.CacheData=Dvbbs.value
Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
For i=0 To Ubound(Forum_Boards)
	Dvbbs.ReloadBoardInfo(Forum_Boards(i))
Next
End sub
Sub clearDate
	If Dvbbs.Boardid=0 Then
		errmsg=errmsg+"<br><li>��ѡ����̳����"
		dvbbs_error()
		Exit Sub
	End If
	Dim Rs,str1,str2,str3,str4
	Set Rs=Dvbbs.Execute("Select Count(*) from dv_topic where Boardid="& Dvbbs.boardid &"")
	str1=Rs(0)
	str3=0
	str4=0
	For i= 0 to UBound(AllPostTable)
		Set Rs=Dvbbs.Execute("Select Count(*) from "&AllPostTable(i)&" where Boardid="& Dvbbs.boardid &"")
		str2=str2&"������"&AllPostTable(i)&"��"&Rs(0)&"ƪ���£�"
		str3=str3+Rs(0)
		Set Rs=Dvbbs.Execute("Select Count(*) from "&AllPostTable(i)&" where Boardid="& Dvbbs.boardid &" and isbest=1")
		str4=str4+Rs(0)
	Next
	Response.Write"<br>"
	Response.Write"<table cellpadding=0 cellspacing=0 align=center class=""tableBorder"" style=""width:90%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" height=25 colspan=2>"
	Response.Write Dvbbs.BoardType
	Response.Write "-������Ϣ"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""ForumrowHighLight"" colspan=2>"
	Response.Write "<li>��������:<b>"
	Response.Write str1
	Response.Write "</b><li>��������:<b>"
	Response.Write str3
	Response.Write "</b><li>"
	Response.Write str2
	Response.Write "<li>��<B>"&str4&"</B>ƪ��������"
	Response.Write"</td></tr>"
	Response.Write "<form action =""?action=delDate&boardid="&Dvbbs.boardid&""" method=post>"
	Response.Write"<tr>"
	Response.Write"<td class=""ForumrowHighLight"" valign=middle colspan=2 align=left><li>  ���<b>"
	Response.Write Dvbbs.BoardType
	Response.Write "</b>�� "
	Response.Write "<select name=""tablelist""><option value=""all"">�������ݱ�</option>"
	For i= 0 to UBound(AllPostTable)
		Response.Write "<option value="""&AllPostTable(i)&""">"
		Response.Write 	AllPostTableName(i)
		Response.Write "</option>"
	Next 
	Response.Write "</select>"
	Response.Write " �� <input type=text name=dd value=365 size=5 > ��ǰ������"
	Response.Write " <input type=""submit"" name=""Submit"" value=""ִ ��""> <b>ע��:�˲������ɻָ���</b>���о��������ᱻɾ����<BR><BR>���������̳�����ڶִ࣬�д˲��������Ĵ����ķ�������Դ��ִ�й��������ĵȺ����ѡ��ҹ���������ٵ�ʱ����¡�"
	Response.Write "</td></tr>"
	Response.Write "</form>"
	Response.Write"</table>"
End Sub
Sub delDate
	If Dvbbs.Boardid=0 Then
		errmsg=errmsg+"<br><li>��ѡ����̳����"
		dvbbs_error()
		Exit Sub
	End If
	Dim tablelist
	If request.form("tablelist")<>"all" Then
		tablelist=Dvbbs.checkstr(request.form("tablelist"))
	Else
		For i= 0 to UBound(AllPostTable)
		If i=0 Then
			tablelist=AllPostTable(i)
		Else
			tablelist=tablelist&","&AllPostTable(i)
		End If
		Next
	End If
	tablelist=split(tablelist,",")
	Dim SqlTopic
	Dim k
	k=0
	For i= 0 to UBound(tablelist)
		'ɾ�����ݱ��¼
		If IsSqlDataBase=1 Then
		SqlTopic="Select TopicID,isvote,PollID from dv_Topic where Boardid="&Dvbbs.boardid&" and isbest=0 and PostTable='"&tablelist(i)&"' and Datediff(d,LastPostTime,"&SqlNowString&") > "& CLng(request.form("dd"))&" "
		Else
		SqlTopic="Select TopicID,isvote,PollID from dv_Topic where Boardid="&Dvbbs.boardid&" and isbest=0 and PostTable='"&tablelist(i)&"' and Datediff('d',LastPostTime,"&SqlNowString&") > "& CLng(request.form("dd"))&" "
		End If
		Set rs=Dvbbs.Execute(SqlTopic)
		Do While Not Rs.Eof
			Sql="Delete from "&tablelist(i)&" where Boardid="&Dvbbs.boardid&" and rootid="&RS(0)&""
			Dvbbs.Execute(Sql) 
			If Rs(1)=1 And Not IsNull(Rs(2)) Then
				Sql="Delete from dv_vote where voteid="&RS(2)&""
				Dvbbs.Execute(Sql)
			End If 
			Rs.movenext
		k=k+1
		Loop 
		'ɾ��������¼
		If IsSqlDataBase=1 Then
		SqlTopic="Delete from dv_Topic where Boardid="&Dvbbs.boardid&" and isbest=0 and PostTable='"&tablelist(i)&"' and Datediff(d,LastPostTime,"&SqlNowString&") > "& CLng(request.form("dd"))&" "
		Else
		SqlTopic="Delete from dv_Topic where Boardid="&Dvbbs.boardid&" and isbest=0 and PostTable='"&tablelist(i)&"' and Datediff('d',LastPostTime,"&SqlNowString&") > "& CLng(request.form("dd"))&" "
		End If
		Dvbbs.Execute(SqlTopic) 
		Set rs=Nothing 	
	Next
	Response.Write "ɾ����"&k&"�����⡣"
End Sub

Sub RestoreClass()
	Dim ClassID,RootID,RootIDNum,ParentID
	ClassID=Request("ClassID")
	If Not IsNumeric(ClassID) Or ClassID="" Then
		Response.Write "����İ��������"
		Exit Sub
	Else
		ClassID=Clng(ClassID)
	End If
	Set Rs=Dvbbs.Execute("Select RootID,BoardID From Dv_Board Where BoardID="&ClassID)
	If Rs.Eof And Rs.Bof Then
		Response.Write "����İ��������"
		Exit Sub
	Else
		RootID=Rs(0)
		ParentID=Rs(1)
	End If
	i=0
	Set Rs=Dvbbs.Execute("Select BoardID,ParentID From Dv_Board Where RootID="&RootID&" Order By ParentID,Orders,Depth")
	Do While Not Rs.Eof
		If Rs(1)=0 Then
			Dvbbs.Execute("UpDate Dv_Board Set Orders="&i&" Where BoardID="&Rs(0))
		Else
			Dvbbs.Execute("UpDate Dv_Board Set Orders="&i&",ParentID="&ParentID&",ParentStr='"&ParentID&"',Depth=1,child=0 Where BoardID="&Rs(0))
		End If
		i=i+1
	Rs.MoveNext
	Loop
	Set Rs=Dvbbs.Execute("Select Count(*) From Dv_Board Where RootID="&RootID)
	RootIDNum=Rs(0)
	If IsNull(RootIDNum) Or RootIDNum="" Then
		RootIDNum=0
	Else
		RootIDNum=RootIDNum-1
	End If
	Dvbbs.Execute("UpDate Dv_Board Set Child="&RootIDNum&" Where BoardID="&ClassID)
	dv_suc("��λ����ɹ���")
	RestoreBoardCache()
	Set Rs=Nothing
End Sub

Sub handorders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<tr> 
	<th height="22">��̳�������������޸�(������Ӧ��̳������������������Ӧ���������)
	</th>
	</tr>
	<tr>
	<td class="Forumrow">
	<B>ע��</B>��<BR>
1�����ڱ���̳�����㷨���ǵݹ飬��������ȷ�����������ţ�����������̳��ʾ��������<font color=red>�����δ��ȷ�˽�˵�����벻Ҫ�������</font><BR>
2��һ�������������Ϊ0������ȷ���룬<font color=blue>�������������������</font><BR>
3���������Ϊ�����������ں��棬<font color=blue>�����ﲻ����������ָ��ĳ�������������������</font>������Ϊ��ȷ���������뷽ʽ��<BR>
<B>����</B> 0<BR>
--��������A 1<BR>
--��������B 2<BR>
----��������A 3<BR>
----��������B 4<BR>
----��������C 5<BR>
--��������C 6<BR>
A.<font color=blue>Ҫ����������C�ᵽ��������A����</font>�����������룺����(0)-����A(1)-����B(2)-����A(<font color=red>4</font>)-����B(<font color=red>5</font>)-����C(<font color=red>3</font>)-����C(6)<BR>
B.<font color=blue>Ҫ�Ѷ�������C�ᵽ��������B����</font>�����������룺����(0)-����A(1)-����B(<font color=red>3</font>)-����A(<font color=red>4</font>)-����B(<font color=red>5</font>)-����C(<font color=red>6</font>)-����C(<font color=red>2</font>)<BR>
B.<font color=blue>Ҫ�Ѷ�������B�ᵽ��������A����</font>�����������룺����(0)-����A(<font color=red>5</font>)-����B(<font color=red>1</font>)-����A(<font color=red>2</font>)-����B(<font color=red>3</font>)-����C(<font color=red>4</font>)-����C(6)
	</td></tr>
<form action="admin_board.asp?action=savehandorders" method=post>
	<tr>
	<td class="Forumrow"><table width="90%">
<%
dim trs,uporders,doorders,RootID
Set Rs=Dvbbs.Execute("Select RootID From Dv_Board Where BoardID="&Request("classid"))
If Rs.eof And Rs.bof Then
	response.write "��û����Ӧ����̳���ࡣ"
	exit sub
Else
	RootID=Rs(0)
End If
set rs = server.CreateObject ("Adodb.recordset")
sql="select * from dv_Board Where RootID="&RootID&" order by RootID,orders"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "��û����Ӧ����̳���ࡣ"
else
	do while not rs.eof
	response.write "<tr><td width=""50%"">"
	if rs("depth")>0 then
	for i=1 to rs("depth")
		response.write "&nbsp;"
	next
	end if
	if rs("child")>0 then
		response.write "<img src=skins/default/plus.gif>"
	else
		response.write "<img src=skins/default/nofollow.gif>"
	end if
	if rs("parentid")=0 then
		response.write "<b>"
	end if
	response.write rs("boardtype")
	if rs("child")>0 then
		response.write "("&rs("child")&")"
	end if
	response.write "</td><td width=""50%"">"
	Response.Write "<input type=hidden value="""&rs("boardid")&""" name=getboard>"
	Response.Write "<input type=text size=5 value="""&rs("orders")&""" name=orders>"
	response.write "</td></tr>"
	uporders=0
	doorders=0
	rs.movenext
	loop
	Response.Write "<tr><td class=Forumrow><input type=submit name=submit value=�ύ></td></tr>"
	response.write "</table>"
end if
rs.close
set rs=nothing
%>
	</td>
	</tr></form>
</table>
<%
End Sub

Sub savehandorders()
	dim cID,OrderID,ClassName
	cID=replace(request.form("getboard"),"'","")
	OrderID=replace(request.form("Orders"),"'","")
	for i=1 to request.form("getboard").count
		cID=request.form("getboard")(i)
		OrderID=request.form("Orders")(i)
		Dvbbs.Execute("Update Dv_Board Set Orders="&OrderID&" Where BoardID="&cID)
	next

	Dv_suc("���ķ�������ɹ���")
	Dvbbs.ReloadAllBoardInfo()
	Dim Forum_Boards
	Dvbbs.Name="setup"
	Dvbbs.ReloadSetup
	Dvbbs.CacheData=Dvbbs.value
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
End Sub

Sub RestoreBoardCache()
	Dim Forum_Boards,i
	Dvbbs.ReloadAllBoardInfo()
	Dvbbs.Name="setup"
	Dvbbs.ReloadSetup
	Dvbbs.CacheData=Dvbbs.value
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
	dv_suc("�ؽ����а��滺��ɹ���")
End Sub

Function SimEncodeJS(str)
	If Not IsNull(str) Then
   		'str = replace(str, "<", "&lt;")
		str = replace(str, "\", "\\")
		str = replace(str, chr(34), "\""")
		str = replace(str, chr(39), "\'")
		str = Replace(str, chr(10), "\n")
		str = Replace(str, chr(13), "\r")
		SimEncodeJS=str
	End If
End Function
%>