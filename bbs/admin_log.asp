<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
Head()
Dim admin_flag
Dim action,actiontype
Dim sqlstr,l_type
action=request("action")
admin_flag=",3,"
If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
	Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	dvbbs_error()
Else 
Select Case action
Case "topic"
	actiontype="贴子管理日志：包括除固顶相关的对贴子的所有操作。"
	sqlstr=" where l_type=3 "
	l_type=3
	main
Case "istop"
	actiontype="固顶操作日志"
	sqlstr=" where l_type=4 "
	l_type=4
	main
Case "wealth"
	actiontype="用户奖惩日志"
	sqlstr=" where l_type=5 "
	l_type=5
	main
Case "users"
	actiontype="用户处理日志：包括屏蔽、锁定、封IP和解除。"
	sqlstr=" where l_type=6 "
	l_type=6
	main
Case "admin0"
	actiontype="后台日志0"
	sqlstr=" where l_type=0 "
	l_type=0
	main
Case "admin1"
	actiontype="后台日志1"
	sqlstr=" where l_type=1 "
	l_type=1
	main
Case "admin2"
	actiontype="后台日志2"
	sqlstr=" where l_type=2 "
	l_type=2
	main
Case "dellog"
	batch()
Case Else
	actiontype="全部日志"
	sqlstr=" "
	l_type=""
	main
End Select

If founderr then call dvbbs_error()
footer()
End If 
Sub main()
'日志分类：后台一般记录,l_type=1,后台重要记录,l_type=2,贴子一般操作:l_type=3，贴子固顶相关l_type=4,奖励惩罚，l_type=5,用户处理 l_type=6
Response.Write "<table width=""95%"" border=""0"" cellspacing=""0"" cellpadding=""0""  align=center class=""tableBorder"" >"
Response.Write "<tr>"
Response.Write "<th width=""100%"" colspan=""6"" class=""tableHeaderText""  height=25>论坛日志管理"
Response.Write "</th>"
Response.Write "</tr>"
Response.Write "<tr>"
Response.Write "<td align=""center"" width=""100%"" colspan=""6"" class=""tableHeaderText""  height=25>当前显示："
Response.Write actiontype
Response.Write "</td>"
Response.Write "</tr>"
Response.Write "<th width=""100%"" colspan=""6"" class=""tableHeaderText""  height=25 id=tabletitlelink >选择查看："
Response.Write " <a href=""?action="">全部日志</a> |"
Response.Write " <a href=""?action=topic"">贴子管理</a> |"
Response.Write " <a href=""?action=istop"">固顶操作</a> |"
Response.Write " <a href=""?action=wealth"">奖惩操作</a> |"
Response.Write " <a href=""?action=users"">用户处理</a> |"
Response.Write " <a href=""?action=admin0"">后台事件0</a> |"
Response.Write " <a href=""?action=admin1"">后台事件1</a> |"
Response.Write " <a href=""?action=admin2"">后台事件2</a> |"
Response.Write "</th>"
Response.Write "</tr>"
Response.Write "</table><br>"
Dim currentpage,page_count,Pcount,endpage
Dim sql,Rs,totalrec
currentPage=request("page")
If currentpage="" or not IsNumeric(currentpage) Then
	currentpage=1
Else
	currentpage=clng(currentpage)
End If
Dvbbs.Forum_Setting(11)=50
sql="select * from [dv_log] "&sqlstr&" order by l_addtime desc"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1

Response.Write "<table width=""95%"" border=""0"" cellspacing=""0"" cellpadding=""0""  align=center class=""tableBorder"" style=""word-break:break-all"" >"
Response.Write "<form action=admin_log.asp?action=dellog&l_type="&l_type&" method=post name=even>"
Response.Write "<tr align=left>"
Response.Write "<th height=25 width=""10%"" >"
Response.Write "对象"
Response.Write "</td>"
Response.Write "<th height=25 width=""55%"" >"
Response.Write "事件内容"
Response.Write "</td>"
Response.Write "<th height=25 width=""20%"">"
Response.Write "操作时间/IP"
Response.Write "</td>"
Response.Write "<th height=25 width=""10%"" >"
Response.Write "操作人"
Response.Write "</td>"
Response.Write "<th height=25 width=""5%"" >"
Response.Write "操作"
Response.Write "</th>"
Response.Write "</tr>"
If Not(Rs.eof or Rs.bof) Then
	rs.PageSize = Dvbbs.Forum_Setting(11)
	rs.AbsolutePage=currentpage
	page_count=0
    	totalrec=rs.recordcount
	While (Not Rs.EOF) And (Not page_count = Rs.PageSize)
	Response.Write "<tr align=left>"
	Response.Write "<td class=""forumrow""  width=""10%"" >"
	Response.Write "<a href=dispuser.asp?name="
	Response.Write Dvbbs.HTMLEncode(rs("l_touser"))
	Response.Write " target=_blank>"
	Response.Write Dvbbs.HTMLEncode(rs("l_touser"))
	Response.Write "</a>"
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""55%"" >"
	Response.Write Dvbbs.HTMLEncode(Rs("l_content"))
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""20%"">"
	Response.Write rs("l_addtime")
	Response.Write "<br>"
	Response.Write Rs("l_ip")
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""10%"">"
	Response.Write "<a href=dispuser.asp?name="&Dvbbs.HTMLEncode(rs("l_username"))&" target=_blank>"&Dvbbs.HTMLEncode(rs("l_username"))&"</a>"
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""5%"">"
	If Rs("l_type")<>2 Then
		Response.Write  "<input type=checkbox name=lid value="&rs("l_id")&">"
	End If
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td height=2></td></tr>"
	
	page_count = page_count + 1
	Rs.MoveNext
	Wend
	Response.Write "<tr><td class=forumrowHighLight colspan=6>请选择要删除的事件，<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">全选 <input type=submit name=act value=删除  onclick=""{if(confirm('您确定执行的操作吗?')){this.document.even.submit();return true;}return false;}"">"
	Response.Write "　<input type=submit name=act onclick=""{if(confirm('确定清除回收站所有的纪录吗?')){this.document.even.submit();return true;}return false;}"" value=清空日志></td></tr>"
	If totalrec mod Dvbbs.Forum_Setting(11)=0 Then
		Pcount= totalrec \ Dvbbs.Forum_Setting(11)
  	Else
  		Pcount= totalrec \ Dvbbs.Forum_Setting(11)+1
  	End If
  	Response.Write "<table border=0 cellpadding=0 cellspacing=3 width="""&Dvbbs.mainsetting(0)&""" align=center>"
  	Response.Write "<tr><td valign=middle nowrap>"
	Response.Write "页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"
	Response.Write "&nbsp;每页<b>"&Dvbbs.Forum_Setting(11)&"</b> 总数<b>"&totalrec&"</b></td>"
	Response.Write "<td valign=middle nowrap align=right>分页："
	If currentpage > 4 Then
		Response.Write "<a href=""?page=1&action="&action&""">[1]</a> ..."
	End If
	If Pcount>currentpage+3 Then
		endpage=currentpage+3
	Else
		endpage=Pcount
	End If
	For i=currentpage-3 to endpage
	If Not i<1 Then
		If i = clng(currentpage) Then
			response.write " <font color="&Dvbbs.mainsetting(1)&">["&i&"]</font>"
		Else
			Response.Write " <a href=""?page="&i&"&action="&action&""">["&i&"]</a>"
		End If
	End If
	Next
	If currentpage+3 < Pcount Then   
		Response.Write "... <a href=""?page="&Pcount&"&action="&action&""">["&Pcount&"]</a>"
	End If
	Response.Write "</td></tr></table>"
Else
	Response.Write "<tr align=center>"
	Response.Write "<td class=""forumrow"" width=""100%"" colspan=""6"" >"
	Response.Write "无相关记录。"
	Response.Write "</td>"
	Response.Write "</tr>"
End If
Response.Write "</form>"
Response.Write "</table>"
Rs.close
Set rs=Nothing
End Sub

Sub batch()
	Dim lid
	If request("act")="删除" Then
		If request.form("lid")="" Then
			DVbbs.AddErrmsg "请指定相关事件。"
		Else
			lid=replace(request.Form("lid"),"'","")
			lid=replace(lid,";","")
			lid=replace(lid,"--","")
			lid=replace(lid,")","")
		End If
	End if
	If request("act")="删除" Then
		Dvbbs.Execute("delete from dv_log where Datediff(""D"",l_addtime, "&SqlNowString&") > 2 and l_id in ("&lid&")")
	ElseIf request("act")="清空日志" Then
		If request("l_type")="" or IsNull(request("l_type")) Then 
			If IsSqlDataBase = 1 Then
			Dvbbs.Execute("delete from dv_log Where Datediff(D,l_addtime, "&SqlNowString&") > 2")
			else
			Dvbbs.Execute("delete from dv_log Where Datediff('D',l_addtime, "&SqlNowString&") > 2")
			end if
		Else
			If IsSqlDataBase = 1 Then
			Dvbbs.Execute("delete from dv_log where  Datediff(D,l_addtime, "&SqlNowString&") > 2 and l_type="&CInt(request("l_type"))&"")
			else
			Dvbbs.Execute("delete from dv_log where  Datediff('D',l_addtime, "&SqlNowString&") > 2 and l_type="&CInt(request("l_type"))&"")
			end if
		End If
	End If
	Dv_suc("成功删除日志。注意：两天内的日志会被系统保留。")
End Sub
%>
<script language="javascript">
function CheckAll(form)  
  {  
  for (var i=0;i<form.elements.length;i++)  
    {  
    var e = form.elements[i];  
    if (e.name != 'chkall')  
       e.checked = form.chkall.checked;  
    }  
  }  
</script>