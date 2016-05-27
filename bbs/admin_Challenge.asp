<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%	
	Head()
	dim admin_flag,rs_c
	admin_flag=",1,"
	If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then 
		Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
		Call dvbbs_error()
	Else
		if request("action")="save" then
		call saveconst()
		elseif request("action")="restore" then
		call restore()
		else
		call consted()
		end if
		Footer()
	end if


Sub consted()
Dim  sel
%>
<form method="POST" action="admin_Challenge.asp?action=save">
<table width="95%" border="0" cellspacing="0" cellpadding="3"  align=center class="tableBorder"> 
<th height=25 colspan=2 align=center id=tabletitlelink><a name="setting20"></a><b>论坛短信互动信息设置</b>[<a href="admin_Challenge.asp?action=restore">还原默认设置</a>]
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>是否开启阳光短信</U><br>阳光短信总开关，选择否则其他阳光短信设置均无效！</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(0)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(0))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(0)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(0))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
<td width="50%" class=Forumrow> <U>是否开启阳光广告</U><br>阳光广告总开关，选择否则阳光互动广告均不显示！</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(1)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(1))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(1)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(1))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>顶部通栏模式</U><br>阳光互动顶部广告显示模式。</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(2)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(2))=0 then%>checked<%end if%>>不显示&nbsp;
<input type=radio name="Forum_ChanSetting(2)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(2))=1 then%>checked<%end if%>>显示&nbsp;
<input type=radio name="Forum_ChanSetting(2)" value=2 <%if cint(Dvbbs.Forum_ChanSetting(2))=2 then%>checked<%end if%>>显示于下拉菜单中&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>顶部banner模式</U><br>论坛顶部所显示的广告。</td>
<td width="50%" class=Forumrow>
<input type=radio name="Forum_ChanSetting(3)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(3))=0 then%>checked<%end if%>>论坛广告&nbsp;
<input type=radio name="Forum_ChanSetting(3)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(3))=1 then%>checked<%end if%>>短信广告&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>尾部通栏模式</U><br>论坛底部显示的广告。</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(4)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(4))=0 then%>checked<%end if%>>论坛广告&nbsp;
<input type=radio name="Forum_ChanSetting(4)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(4))=1 then%>checked<%end if%>>短信广告&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>帖间广告模式</U><br>帖子间是否显示短信广告。</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(5)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(5))=0 then%>checked<%end if%>>否&nbsp;(显示论坛贴间广告)
<input type=radio name="Forum_ChanSetting(5)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(5))=1 then%>checked<%end if%>>是&nbsp;(显示短信贴间广告)
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>是否开启站内短信互发</U><br>　</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(6)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(6))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(6)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(6))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>是否开启主题订阅</U><br>　</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(7)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(7))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(7)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(7))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>是否开启VIP论坛</U><br>　</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(8)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(8))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(8)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(8))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>是否允许阳光会员注册、修改资料</U><br>　</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(9)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(9))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(9)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(9))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>是否允许无边界登录</U><br>　</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(10)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(10))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(10)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(10))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>是否开启同步广告功能</U><br>　</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(11)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(11))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(11)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(11))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>登录和注册成功是否出现阳光会员简介</U><br>　</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(12)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(12))=0 then%>checked<%end if%>>否&nbsp;
<input type=radio name="Forum_ChanSetting(12)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(12))=1 then%>checked<%end if%>>是&nbsp;
</td>
</tr>

<tr>
<td width="50%" class=Forumrow> &nbsp;</td>
<td width="50%" class=Forumrow>  
<div align="center"> 
<input type="submit" name="Submit" value="提 交">
</div>
</td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
dim Forum_ChanSetting,ChanSetting

'Forum_ChanSetting=request.form("Forum_ChanSetting(0)") & "," & request.form("Forum_ChanSetting(1)") & "," & request.form("Forum_ChanSetting(2)") & "," & request.form("Forum_ChanSetting(3)") & "," & request.form("Forum_ChanSetting(4)") & "," & request.form("Forum_ChanSetting(5)") & "," & request.form("Forum_ChanSetting(6)") & "," & request.form("Forum_ChanSetting(7)") & "," & request.form("Forum_ChanSetting(8)") & "," & request.form("Forum_ChanSetting(9)") & "," & request.form("Forum_ChanSetting(10)") & "," & request.form("Forum_ChanSetting(11)") & "," & request.form("Forum_ChanSetting(12)")
For i=0 To 60
	If Request.Form("Forum_ChanSetting("&i&")")="" Then
		ChanSetting = 1
	Else
		ChanSetting = Replace(Request.Form("Forum_ChanSetting("&i&")"),",","")
	End If
	If i = 0 Then
		Forum_ChanSetting = ChanSetting
	Else
		Forum_ChanSetting = Forum_ChanSetting & "," & ChanSetting
	End If
Next

sql="update Dv_setup set Forum_ChanSetting='"&Forum_ChanSetting&"'"
dvbbs.execute(sql)
Dvbbs.Name="setup"
Dvbbs.reloadsetup()
Dv_suc("设置短信互动功能成功")

end sub

'恢复默认设置
Sub restore()
	Dvbbs.Forum_ChanSetting="1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"
	Dvbbs.Execute("update Dv_setup set Forum_ChanSetting='"&dvbbs.Forum_ChanSetting&"'")
	Dv_suc("还原设置")
End Sub 
%>