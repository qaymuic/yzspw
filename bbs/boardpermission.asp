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
<!--版面权限浏览，主体Js部分-->
<SCRIPT LANGUAGE="JavaScript">
<!--
var groupname=new Array();
var GroupSetting=new Array();
function showtoptable(orders)
{
	document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
	document.write ('<tr>');
	document.write ('<th height="25" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=1 >浏览权限<a></th>');
	document.write ('<th height="23" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=2 >发帖权限<a></th>');
	document.write ('<th height="23" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=3 >帖子管理权限<a></th>');
	document.write ('<th height="23" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=4 >其他权限<a></th>');
	document.write ('<th height="23" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=5 >管理权限<a></th>');
	document.write ('<th height="23" width=16% id=tabletitlelink><a href=?boardid={$boardid}&orders=6 >短信权限<a></th>');
	document.write ('</tr>');
	document.write ('</table>');
	switch (orders)
	{
		case 1:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=5 class=tablebody1>浏览权限</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=20% class=tablebody2>用户组名('+groupname.length+')</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以浏览论坛</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以查看会员信息(包括其他会员的资料和会员列表)</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以查看其他人发布的主题</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以浏览精华帖子</td>');
		document.write ('</tr>');
		break;
		case 2:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=13 class=tablebody1>发帖权限</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=16% class=tablebody2>用户组名('+groupname.length+')</td>');
		document.write ('<td height="23" width=7% class=tablebody2>可以发布新主题</td>');
		document.write ('<td height="23" width=7% class=tablebody2>可以回复自己的主题</td>');
		document.write ('<td height="23" width=7% class=tablebody2>可以回复其他人的主题</td>');
		document.write ('<td height="23" width=7% class=tablebody2>可以在论坛允许评分的时候参与评分(鲜花和鸡蛋)?</td>');
		document.write ('<td height="23" width=7% class=tablebody2>参与评分所需金钱</td>');
		document.write ('<td height="23" width=7% class=tablebody2>可以上传附件</td>');
		document.write ('<td height="23" width=7% class=tablebody2>最多上传文件个数</td>');
		document.write ('<td height="23" width=7% class=tablebody2>上传文件大小限制</td>');
		document.write ('<td height="23" width=7% class=tablebody2>可以发布新投票</td>');
		document.write ('<td height="23" width=7% class=tablebody2>可以参与投票</td>');
		document.write ('<td height="23" width=7% class=tablebody2>可以发布小字报</td>');
		document.write ('<td height="23" width=7% class=tablebody2>发布小字报所需金钱</td>');
		document.write ('</tr>');
		break;
		case 3:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=5 class=tablebody1>帖子管理权限</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=20% class=tablebody2>用户组名('+groupname.length+')</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以编辑自己的帖子</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以删除自己的帖子</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以移动自己的帖子到其他论坛</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以打开/关闭自己发布的主题</td>');
		document.write ('</tr>');
		break;
		case 4:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=8 class=tablebody1>其他权限</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=12.5% class=tablebody2>用户组名('+groupname.length+')</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>可以搜索论坛</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>可以使用\'发送本页给好友\'功能</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>可以修改个人资料</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>帖子中可以查看其它人头衔</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>帖子中可以查看其它人头像</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>帖子中可以可以查看其它人签名</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>可以浏览论坛事件</td>');
		document.write ('</tr>');
		break;
		case 5:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=18 class=tablebody1>管理权限</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=10% class=tablebody2>用户组名('+groupname.length+')</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以删除其它人帖子</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以移动其它人帖子</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以打开/关闭其它人帖子</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以固顶/解除固顶帖子</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以奖励/惩罚发贴用户</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以奖励/惩罚用户</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以编辑其它人帖子</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以加入/解除精华帖子</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以发布公告</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以管理公告</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以管理小字报</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以锁定/屏蔽/解除锁定用户</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以删除用户1－10天内所发帖子</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以查看来访IP及来源</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以限定IP来访</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以管理用户权限</td>');
		document.write ('<td height="23" width=5% class=tablebody2>可以批量删除帖子（前台）</td>');
		document.write ('</tr>');
		break;
		case 6:
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=5 class=tablebody1>短信权限</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=12.5% class=tablebody2>用户组名('+groupname.length+')</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>可以发送短信</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>最多发送用户</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>短信内容大小限制</td>');
		document.write ('<td height="23" width=12.5% class=tablebody2>信箱大小限制</td>');
		document.write ('</tr>');
		break ;
		case "":
		document.write ('<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>');
		document.write ('<tr  align=center>');
		document.write ('<td colspan=5 class=tablebody1>其他权限</td>');
		document.write ('</tr>');
		document.write ('<tr>');
		document.write ('<td height="25" width=20% class=tablebody2>用户组名('+groupname.length+')</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以浏览论坛</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以查看会员信息(包括其他会员的资料和会员列表)</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以查看其他人发布的主题</td>');
		document.write ('<td height="23" width=20% class=tablebody2>可以浏览精华帖子</td>');
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
		return('√')
	}
	else
	{
		return('<font color={$alertcolor}>×</font>')
	}
	
	
}
//-->
</SCRIPT>