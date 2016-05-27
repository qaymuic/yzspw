<%
Dim AllPostTable
Dim AllPostTableName
Dim FoundErr
FoundErr=False 
Dim ErrMsg
Dim Rs,sql,i
Dvbbs.LoadTemplates("")
Set Rs=Dvbbs.Execute("Select H_Content From Dv_Help Where H_ID=1")
template.value = Rs(0)
'页面错误提示信息
Sub dvbbs_error()
	Response.Write"<br>"
	Response.Write"<table cellpadding=3 cellspacing=1 align=center class=""tableBorder"" style=""width:75%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" height=25 colspan=2>错误信息"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""Forumrow"" colspan=2>"
	Response.Write ErrMsg
	Response.Write"</td></tr>"
	Response.Write"<tr>"
	Response.Write"<td class=""Forumrow"" valign=middle colspan=2 align=center><a href=""javascript:history.go(-1)""><<返回上一页</a></td></tr>"
	Response.Write"</table>"
	footer()
	Response.End 
End Sub 
Function AllPostTable1()
	Dim Trs
	Set Trs=Dvbbs.Execute("select * from [Dv_TableList]")
	AllPostTable=""
	Do While Not TRs.EOF
		If AllPostTable=""  Then 
			AllPostTable=TRs("TableName")
			AllPostTableName=TRs("TableType")
		Else
			AllPostTable=AllPostTable&"|"&TRs("TableName")
			AllPostTableName=AllPostTableName&"|"&TRs("TableType")
		End If
	TRs.MoveNext
	Loop 
	Trs.Close 
	
End Function 
AllPostTable1
AllPostTableName=Split(AllPostTableName,"|")
AllPostTable=Split(AllPostTable,"|")
Dim NowUseBbs
NowUseBbs=Dvbbs.NowUseBbs

Sub footer()
	Response.Write"<table align=center >"
	Response.Write "<tr align=center><td width=""100%"" class=copyright>"
	Response.Write"Dvbbs v7.0 , Copyright (c) 2001-2005 <a href=""http://www.aspsky.net"" target=""_blank""><font color=#708796><b>AspSky<font color=#CC0000>.Net</font></b></font></a>. All Rights Reserved ."
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"</table>"

	Response.Write "<div id=footjs style=""display:none"">"
	Dim Yrs,dupack
	set Yrs=Dvbbs.Execute("select * from Dv_setup")
	dupack=Split(Yrs("forum_pack"),"|||")
	If Cint(dupack(0))=0 Then
		Response.Write "<a href=""http://bbs.dvbbs.net"" target=_blank><font color="&Dvbbs.Mainsetting(1)&">官方讨论区</font></a>"
	Else
		Response.Write "<script src=""http://bbs.dvbbs.net/packjs.asp?a="&dupack(1)&"&b="&dupack(2)&"&c="&Dvbbs.Forum_info(0)&"&d="&Dvbbs.Get_ScriptNameUrl&"&e="&Yrs("Forum_usernum")&"&f="&Yrs("Forum_PostNum")&"&g="&Dvbbs.Forum_info(5)&"&h="&IsSqlDatabase&"&i=Dvbbs Version "&Dvbbs.Forum_Version&"""></script>"
	End if
	Yrs.close:set Yrs=nothing
	Response.Write "</div>"
	%>
	<script language="javascript">
		document.getElementById("packjs").innerHTML=document.getElementById("footjs").innerHTML;
	</script>
	<%
	Response.Write "</body>"
	Response.Write "</html>"
	SaveLog()
End Sub
Sub Head()
	Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
	Response.Write Chr(10)	
	Response.Write "<html>"
	Response.Write Chr(10)	
	Response.Write "<head>"
	Response.Write Chr(10)	
	Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
	Response.Write Chr(10)	
	Response.Write "<meta name=keywords content=""动网先锋,动网论坛,dvbbs"">"
	Response.Write Chr(10)	
	Response.Write "<meta name=""description"" content=""Design By www.dvbbs.net"">"
	Response.Write Chr(10)	
	Response.Write "<title>"& Dvbbs.Forum_info(0)&"-管理页面</title>"
%>
<!--默认风格--> 
<style type=text/css>
.menuskin {
	BORDER: #666666 1px solid; VISIBILITY: hidden; FONT: 12px Verdana;
	POSITION: absolute; 
	BACKGROUND-COLOR:#EFEFEF;
	background-image:url("skins/default/dvmenubg3.gif");
	background-repeat : repeat-y;
	}
.menuskin A {
	PADDING-RIGHT: 10px; PADDING-LEFT: 25px; COLOR: black; TEXT-DECORATION: none; behavior:url(inc/noline.htc);
	}
#mouseoverstyle {
	BACKGROUND-COLOR: #C9D5E7; margin:2px; padding:0px; border:#597DB5 1px solid;
	}
#mouseoverstyle A {
	COLOR: black
}
.menuitems{
	margin:2px;padding:1px;word-break:keep-all;
}


}
</style>
<%
	Response.Write Chr(10)	
	Response.Write template.html(1)
	Response.Write Chr(10)
	Response.Write "</head>"
	Response.Write "<script src=""images/manage/admin.js"" type=""text/javascript""></script><script src=""inc/main.js"" type=""text/javascript""></script>"
	Response.Write Chr(10)
	Response.Write "<body leftmargin=""0"" topmargin=""0"" marginheight=""0"" marginwidth=""0"">"
	Response.Write Chr(10)
%>
<div class=menuskin id=popmenu 
      onmouseover="clearhidemenu();highlightmenu(event,'on')" 
      onmouseout="highlightmenu(event,'off');dynamichide(event)" style="Z-index:100"></div>
<table cellpadding="3" cellspacing="0" border="0" align=center class="tableBorder1" style="width:100%">
<tr><td class="bodytitle" height=25>
<table height="100%" width="100%" border=0 cellpadding=0 cellspacing=0>
<tr valign=middle>
	<td width=40><img src="images/manage/i_home.gif">
	</td>
	<td width=150>
		动网论坛系统设置面板
	</td>
	<td width=40>
	</td>
	<td width=100>
		<a href="admin_admin.asp" target=main>修改管理员资料</a>
	</td>
	<%if Dvbbs.Forum_ChanSetting(0)=1 then%>
	<%
	set rs=Dvbbs.Execute("select top 1 * from dv_challengeinfo")
	%>
	<td width=40>
	</td>
	<td width=80>
		<a href="http://bbs.ray5198.com/login_new.jsp?username=<%=rs("D_username")%>&fourmid=<%=rs("D_ForumID")%>&css=Get_CSS.asp?skinid=1&url=<%=Dvbbs.Get_ScriptNameUrl%>" target=main><font color=blue>站长收益查询</font></a>
	</td>
	<%end if%>
	<td width=10>
	</td>
	<td width=120 align=right id=packjs></td>
	<td width="*" align=right>
		<a href="index.asp" target=_top>论坛首页</a>
	</td>
</tr>
</table>
</td></tr>
<tr><td height=10></td></tr>
</table>
<%
End Sub
Sub Dv_suc(info)
	Response.Write"<br>"
	Response.Write"<table cellpadding=0 cellspacing=0 align=center class=""tableBorder"" style=""width:75%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" height=25 colspan=2>成功信息"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""forumRowHighlight"" colspan=2 height=25>"
	Response.Write info
	Response.Write"</td></tr>"
	Response.Write"<tr>"
	Response.Write"<td class=""forumRowHighlight"" valign=middle colspan=2 align=center><a href="&Request.ServerVariables("HTTP_REFERER")&" ><<返回上一页</a></td></tr>"
	Response.Write"</table>"
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

Function Reload_All_Board_Cache()
	'更新版面列表缓存
	ReloadBoardListAll
	'更新单个版面缓存（循环）
	Dim BoardListAll,BoardListNum,myBoardID
	BoardListAll=myCache.value
	BoardListNum=Ubound(BoardListAll,2)
	For i=0 To BoardListNum
		myBoardID=BoardListAll(0,i)
		ReloadBoardInfo(myBoardID)
		Set rs=Dvbbs.Execute("Select ParentStr from board where boardid="&myBoardID)
		If not rs.eof Then
			Dvbbs.ReloadBoardParentStr(rs(0))
		End If
		Rs.close
		Set Rs=nothing
	Next
End Function
Sub SaveLog()
	On Error Resume Next
	Dim RequestStr
	RequestStr=lcase(Request.ServerVariables("Query_String"))
	If RequestStr<>"" Then 
		RequestStr=Dvbbs.checkStr(RequestStr)
		RequestStr=Left(RequestStr,250)
		sql="insert into [Dv_log] (l_touser,l_username,l_content,l_ip,l_type) values ('"&Dvbbs.ScriptName&"','"&Dvbbs.membername&"','"&RequestStr&"','"&Dvbbs.UserTrueIP&"',0)"		
		Dvbbs.Execute(sql)
	End If
	If request.form<>"" Then
		RequestStr=Dvbbs.checkStr(request.form)
		RequestStr=Left(RequestStr,250)
		sql="insert into [Dv_log] (l_touser,l_username,l_content,l_ip,l_type) values ('"&Dvbbs.ScriptName&"','"&Dvbbs.membername&"','"&RequestStr&"','"&Dvbbs.UserTrueIP&"',1)"		
		Dvbbs.Execute(sql)
	End If
End Sub
%>