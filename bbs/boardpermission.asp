<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dvbbs.LoadTemplates("help_permission")
Dim orders
If Request("Action")="Myinfo" Then
	Dvbbs.stats=template.Strings(4)
Else
	Dvbbs.stats=template.Strings(0)
End If
Dvbbs.nav()
If Dvbbs.BoardID=0 then
	Dvbbs.Head_var 2,0,"",""
Else
	Dvbbs.Head_var 1,Dvbbs.Board_Data(4,0),"",""
End If

If Not (Request("Action")="Myinfo" And Dvbbs.UserID=0) Then
	If Cint(Dvbbs.GroupSetting(39))=0 And Not Dvbbs.master Then Dvbbs.AddErrCode(55)
End If
Dvbbs.ShowErr

If Not IsNumeric(request("orders")) or request("orders")="" Then
	orders=1
Else
	orders=request("orders")
End If

permission()
Dvbbs.activeonline()
Dvbbs.footer()


Sub permission()
	Response.Write Replace(Replace(Replace(template.html(0),"{$boardid}",Dvbbs.BoardID),"{$alertcolor}",Dvbbs.mainsetting(1)),"{$action}",Request("Action"))
	Response.Write "<Script Language=JavaScript>"
	dim trs,ars,rs
	If Request("Action")="Myinfo" Then
		Dim myper_1,myper_2,myper_3
		Dim UserTitle,MyGroupSetting
		myper_1=false
		myper_2=false
		myper_3=false
		Set Rs=Dvbbs.Execute("Select Uc_userid,uc_Setting From Dv_UserAccess Where Uc_boardid="&Dvbbs.Boardid&" And Uc_userid="&Dvbbs.Userid)
		If Not(Rs.Eof And Rs.Bof) Then
			myper_1=true
			MyGroupSetting = Rs(1)
			UserTitle = template.Strings(1)
		End If
		If not myper_1 Then
			Set Rs=Dvbbs.Execute("Select Pid,PSetting From Dv_BoardPermission Where Boardid="&Dvbbs.boardid&" and GroupID="&Dvbbs.UserGroupID)
			If Not(Rs.Eof And Rs.Bof) Then
				myper_2=true
				MyGroupSetting = Rs(1)
				UserTitle = template.Strings(2)
			End If
		End If
		If not(myper_1 and myper_2) Then
			Set Rs=Dvbbs.Execute("Select UserGroupID,GroupSetting,Usertitle From Dv_UserGroups Where UserGroupID="&Dvbbs.UserGroupID)
			If Not(Rs.Eof And Rs.Bof) Then
				myper_3=true
				MyGroupSetting = Rs(1)
				UserTitle = Rs(2) & template.Strings(3)
			End If
		End If
		Set Rs=Nothing
		Response.Write "groupname[0]='"
		Response.Write UserTitle
		Response.Write "';"
		Response.Write "GroupSetting[0]='"
		Response.Write MyGroupSetting
	    Response.Write "';"
	Else
		Set trs=dvbbs.execute("select * from dv_usergroups Where IsSetting=1 order by usergroupid")
		Dim i
		i=0
		Do While Not trs.EOF
		Response.Write "groupname["
		Response.Write i
	    Response.Write "]='"
		Response.Write Trim(trs("usertitle"))   
        Response.Write "';"    
		Set ars=dvbbs.Execute("select * from dv_BoardPermission where BoardID="&Dvbbs.boardid&" and GroupID="&trs("UserGroupID"))
        If  Not ars.EOF  Then
	    	Response.Write "GroupSetting["
			Response.Write i
       		Response.Write "]='"
        	Response.Write ars("PSetting")
	    	Response.Write "';"
		Else
       		Response.Write "GroupSetting["
        	Response.Write i
	    	Response.Write "]='"
			Response.Write trs("GroupSetting")
        	Response.Write "';"
	    End If
		i=i+1
        trs.MoveNext  
		Loop
		set trs=Nothing 
		set ars=Nothing
	End If
	Response.Write "showtoptable("&orders&")"
	Response.Write "</script>"
End Sub  
%>
<!--����Ȩ�����������Js����-->
<SCRIPT LANGUAGE="JavaScript">
<!--
var groupname=new Array();
var GroupSetting=new Array();
function showtoptable(orders)
{
	document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
	document.write ('<tr>');
	document.write ('<th height="25" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=1 >���Ȩ��<a></th>');
	document.write ('<th height="23" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=2 >����Ȩ��<a></th>');
	document.write ('<th height="23" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=3 >���ӹ���Ȩ��<a></th>');
	document.write ('<th height="23" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=4 >����Ȩ��<a></th>');
	document.write ('<th height="23" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=5 >����Ȩ��<a></th>');
	document.write ('<th height="23" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=6 >����Ȩ��<a></th>');
	document.write ('</tr>');
	document.write ('</table>');
	switch (orders)
	{
		case 1:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=5 class=tablebody1>���Ȩ��</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=20% class=tablebody2>�û�����('+groupname.length+')</td>');
		document.write ('<td height="23" width=20% class=tablebody2>���������̳</td>');
		document.write ('<td height="23" width=20% class=tablebody2>���Բ鿴��Ա��Ϣ(����������Ա�����Ϻͻ�Ա�б�)</td>');
		document.write ('<td height="23" width=20% class=tablebody2>���Բ鿴�����˷���������</td>');
		document.write ('<td height="23" width=20% class=tablebody2>���������������</td>');
		document.write ('</tr>');
		break;
		case 2:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=13 class=tablebody1>����Ȩ��</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=16% class=tablebody2>�û�����('+groupname.length+')</td>');
		document.write ('<td height="23" width=7% class=tablebody2>���Է���������</td>');
		document.write ('<td height="23" width=7% class=tablebody2>���Իظ��Լ�������</td>');
		document.write ('<td height="23" width=7% class=tablebody2>���Իظ������˵�����</td>');
		document.write ('<td height="23" width=7% class=tablebody2>��������̳�������ֵ�ʱ���������(�ʻ��ͼ���)?</td>');
		document.write ('<td height="23" width=7% class=tablebody2>�������������Ǯ</td>');
		document.write ('<td height="23" width=7% class=tablebody2>�����ϴ�����</td>');
		document.write ('<td height="23" width=7% class=tablebody2>����ϴ��ļ�����</td>');
		document.write ('<td height="23" width=7% class=tablebody2>�ϴ��ļ���С����</td>');
		document.write ('<td height="23" width=7% class=tablebody2>���Է�����ͶƱ</td>');
		document.write ('<td height="23" width=7% class=tablebody2>���Բ���ͶƱ</td>');
		document.write ('<td height="23" width=7% class=tablebody2>���Է���С�ֱ�</td>');
		document.write ('<td height="23" width=7% class=tablebody2>����С�ֱ������Ǯ</td>');
		document.write ('</tr>');
		break;
		case 3:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=5 class=tablebody1>���ӹ���Ȩ��</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=20% class=tablebody2>�û�����('+groupname.length+')</td>');
		document.write ('<td height="23" width=20% class=tablebody2>���Ա༭�Լ�������</td>');
		document.write ('<td height="23" width=20% class=tablebody2>����ɾ���Լ�������</td>');
		document.write ('<td height="23" width=20% class=tablebody2>�����ƶ��Լ������ӵ�������̳</td>');
		document.write ('<td height="23" width=20% class=tablebody2>���Դ�/�ر��Լ�����������</td>');
		document.write ('</tr>');
		break;
		case 4:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=8 class=tablebody1>����Ȩ��</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=12.5% class=tablebody2>�û�����('+groupname.length+')</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>����������̳</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>����ʹ��\'���ͱ�ҳ������\'����</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>�����޸ĸ�������</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>�����п��Բ鿴������ͷ��</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>�����п��Բ鿴������ͷ��</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>�����п��Կ��Բ鿴������ǩ��</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>���������̳�¼�</td>');
		document.write ('</tr>');
		break;
		case 5:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=18 class=tablebody1>����Ȩ��</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=10% class=tablebody2>�û�����('+groupname.length+')</td>');
		document.write ('<td height="23" width=5% class=tablebody2>����ɾ������������</td>');
		document.write ('<td height="23" width=5% class=tablebody2>�����ƶ�����������</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Դ�/�ر�����������</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Թ̶�/����̶�����</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Խ���/�ͷ������û�</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Խ���/�ͷ��û�</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Ա༭����������</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Լ���/�����������</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Է�������</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Թ�����</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Թ���С�ֱ�</td>');
		document.write ('<td height="23" width=5% class=tablebody2>��������/����/��������û�</td>');
		document.write ('<td height="23" width=5% class=tablebody2>����ɾ���û�1��10������������</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Բ鿴����IP����Դ</td>');
		document.write ('<td height="23" width=5% class=tablebody2>�����޶�IP����</td>');
		document.write ('<td height="23" width=5% class=tablebody2>���Թ����û�Ȩ��</td>');
		document.write ('<td height="23" width=5% class=tablebody2>��������ɾ�����ӣ�ǰ̨��</td>');
		document.write ('</tr>');
		break;
		case 6:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=5 class=tablebody1>����Ȩ��</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=12.5% class=tablebody2>�û�����('+groupname.length+')</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>���Է��Ͷ���</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>��෢���û�</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>�������ݴ�С����</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>�����С����</td>');
		document.write ('</tr>');
		break ;
		case "":
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=5 class=tablebody1>����Ȩ��</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=20% class=tablebody2>�û�����('+groupname.length+')</td>');
		document.write ('<td height="23" width=20% class=tablebody2>���������̳</td>');
		document.write ('<td height="23" width=20% class=tablebody2>���Բ鿴��Ա��Ϣ(����������Ա�����Ϻͻ�Ա�б�)</td>');
		document.write ('<td height="23" width=20% class=tablebody2>���Բ鿴�����˷���������</td>');
		document.write ('<td height="23" width=20% class=tablebody2>���������������</td>');
		document.write ('</tr>');
		baeak;
	}
	for (i=0;i<groupname.length;i++)
	{
	GroupSetting[i]=GroupSetting[i].split(",")
	
	switch (orders)
	{	
		case 1:
		document.write ('<tr>');
		document.write ('<td height="23"  class=tablebody1>'+groupname[i]+'</td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][0])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][1])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][2])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][41])+'</B></td>');
		document.write ('</tr>');
		break;
		case 2:
		document.write ('<tr>');
		document.write ('<td height="23"  class=tablebody1>'+groupname[i]+'</td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][3])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][4])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][5])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][6])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+GroupSetting[i][47]+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][7])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+GroupSetting[i][40]+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+GroupSetting[i][44]+'</B> KB</td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][8])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][9])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][17])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+GroupSetting[i][46]+'</B></td>');
		document.write ('</tr>');
		
		break;
		case 3:
		document.write ('<tr>');
		document.write ('<td height="23"  class=tablebody1>'+groupname[i]+'</td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][10])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][11])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][12])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][13])+'</B></td>');
		document.write ('</tr>');

		break;
		case 4:
		document.write ('<tr>');
		document.write ('<td height="23"  class=tablebody1>'+groupname[i]+'</td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][14])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][15])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][16])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][36])+'</td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][37])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][38])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][39])+'</B></td>');
		document.write ('</tr>');

		break;
		case 5:
		document.write ('<tr>');
		document.write ('<td height="23"  class=tablebody1>'+groupname[i]+'</td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][18])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][19])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][20])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][21])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][22])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][43])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][23])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][24])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][25])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][26])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][27])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][28])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][29])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][30])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][31])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][42])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][45])+'</B></td>');
		document.write ('</tr>');

		break;
		case 6:
		document.write ('<tr>');
		document.write ('<td height="23"  class=tablebody1>'+groupname[i]+'</td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][32])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+GroupSetting[i][33]+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+GroupSetting[i][34]+'</B> byte</td>');
		document.write ('<td height="23"  class=tablebody1><B>'+GroupSetting[i][35]+'</B> KB</td>');
		document.write ('</tr>');

		break;
		case "":
		document.write ('<tr>');
		document.write ('<td height="23"  class=tablebody1>'+groupname[i]+'</td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][0])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][1])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][2])+'</B></td>');
		document.write ('<td height="23"  class=tablebody1><B>'+YesOrNo(GroupSetting[i][41])+'</B></td>');
		document.write ('</tr>');

		break;
	}
	}
	document.write ('</table>');
	
	
}
function YesOrNo(thiskey)
{	if (thiskey=='1')
	{
		return('��')
	}
	else
	{
		return('<font color={$alertcolor}>��</font>')
	}
	
	
}
//-->
</SCRIPT>