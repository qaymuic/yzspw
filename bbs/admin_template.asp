<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!-- #include file="inc/DvADChar.asp" -->
<%
Head()
Dim admin_flag
admin_flag=",20,"
If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then
	Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	dvbbs_error()
End If
Dim action,SkinID,StyleID
StyleID=request("StyleID")
'取当前系统默认模板编号
SkinID=Dvbbs.cachedata(17,0)
If StyleID="" Then
	StyleID=SkinID
ElseIf Not IsNumeric(StyleID) Then
	StyleID=SkinID
End If
StyleID=CLng(StyleID)

action=Request("action")
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
Response.Write "<tr>"
Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=2 height=25>论坛模板管理"
Response.Write "</th>"
Response.Write "</tr>"
Response.Write "<tr>"
Response.Write "<td class=""forumRowHighlight"" colspan=2>"
Response.Write "<p><B>注意</B>：<BR>①在这里，您可以新建和修改模板，可以编辑论坛语言包和风格，可以新建模板页面，操作时请按照相关页面提示完整填写表单信息。<BR>②论坛当前正在使用的默认模板不能删除<br>③如果修改分模板页面名称或删除分模板页面请在关闭论坛之后操作,否则可能会影响论坛访问."
Response.Write "</td>"
Response.Write "</tr>"
Response.Write "<tr>"
Response.Write "<td class=""forumRowHighlight"" width=20% height=25 align=left>"
Response.Write "<B>论坛模板操作选项</B></td>"
Response.Write "<td class=""forumRowHighlight"" width=""80%""><a href=""admin_template.asp"">模板管理首页</a>"
Response.Write "</td>"
Response.Write "</tr>"
Response.Write "</table>"
Select Case action
	Case "edit"
		Call Edit() 
	Case "manage"
		If Request("mostyle")="编 辑" Then
			Main()
		ElseIf Request("mostyle") = "删 除" Then
			DelStyle()
		End If
	Case "saveedit"
		Call Saveedit()
	Case "addpage"
		addpage()
	Case "addstyle"
		addstyle()
	Case "Edit_Main"
		Edit_Main()
	Case "Save_Main"
		Save_Main()
	Case "rename"
		rename()
	Case "editcss"
		editcss()
	Case "savecss"
		savecss()
	Case "editmain"
		editmain()
	Case "savemain"
		Savemain() 
	Case "ghost"
		ghost()
	Case "delpage"
		delpage()
	Case "pagerename"
		pagerename()
	Case Else
		Main()
End Select

footer()
Sub Main()
	Dim NowOpenTemplate
	Response.Write "<p></p>"
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=2 height=25>当前论坛模板管理"
	Response.Write "</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<form method=post action=""?action=manage"">"
	Response.Write "<td class=""forumRowHighlight"" height=40 align=left>"
	Response.Write "请选择相关模板： "
	'利用系统缓存数据取得所有模板名称和ID
	Dim Templateslist
	Dvbbs.Name="Templateslist"
	If Dvbbs.ObjIsEmpty() Then  Dvbbs.ReloadTemplateslist()
	Templateslist= Dvbbs.Value
	Templateslist=split(Templateslist,"@@@")
	Response.Write "<select name=""StyleID"" size=1>"
	For i=0 to UBound(Templateslist)
		Templateslist(i)=split(Templateslist(i),"|||")
		Response.Write "<option value="""&Templateslist(i)(0)&""""
		If CLng(Templateslist(i)(0)) = CLng(StyleID) Then 
			Response.Write " selected"
			NowOpenTemplate=Templateslist(i)(1)
		End If 
		Response.Write ">"&Templateslist(i)(1)&"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;&nbsp;"
	Response.Write "<input type=submit value=""编 辑"" name=""mostyle"">&nbsp;&nbsp;&nbsp;"
	Response.Write "<input type=submit value=""删 除"" name=""mostyle"">"
	Response.Write "<br><br><b>说明：</b>删除操作将删除该模板所有数据，慎用。"
	Response.Write "</td>"
	Response.Write "</FORM>"
	Response.Write "<FORM METHOD=POST ACTION=""?action=addpage"">"
	Response.Write "<td class=""forumRowHighlight"" height=25 align=left>"
	Response.Write "新建分模板页面：&nbsp;"
	Response.Write "<input type=text size=25 name=""StylePageName"">&nbsp;&nbsp;"
	Response.Write "<input type=submit name=submit value=""添 加""> "
	Response.Write "</td>"
	Response.Write "</FORM>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=25 align=right colspan=2>"
	Response.Write "↑请输入ASP页面名（不要包含后缀），新建立的页面模板既是该页面的模板资源（包括图片、语言、风格）↑"
	Response.Write "</td></tr>"
	Response.Write "<tr>"
	Response.Write "<th height=25 colspan=2>"
	Response.Write NowOpenTemplate&"－－模板资源管理</th></tr><tr><td height=25 class=""bodytitle"" colspan=2>"
	Response.Write "通常来说，分页面模板就是论坛中每个页面的风格模板，括号中是字段名，字段的命名规则为：Page_页面名（不要后缀）"
	Response.Write "</td>"
	Response.Write "</tr>"
	Set Rs=Dvbbs.Execute("Select top 1 * From Dv_Style ")
	For i= 2 to Rs.Fields.Count-1
		If i = 13 Then
			Response.Write "<tr onmouseover=""this.style.backgroundColor='#B3CFEE';this.style.color='red'"" onmouseout=""this.style.backgroundColor='';this.style.color=''"">"
			Response.Write "<td height=25>"
			Response.Write "<li>"
			Response.Write "分页面模板<a href=#>(page_admin)</a>&nbsp;&nbsp;</td><td height=""25"" align=""left"">"
			Response.Write "编辑该模块："
			Response.Write "<a href=""?action=Edit_Main&stype=1&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>语言包</a> <a href=""?action=Edit_Main&stype=1&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""><font color=blue>新</font></a> | <a href=""?action=Edit_Main&stype=2&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>图片</a> <a href=""?action=Edit_Main&stype=2&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new"""
			Response.Write "><font color=blue>新</font></a> | <a href=""?action=Edit_Main&stype=3&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>界面风格</a> <a href=""?action=Edit_Main&stype=3&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""><font color=blue>新</font></a>"
			Response.Write "</td>"
			Response.Write "</tr>"
		End If
		If i> 21 Then
			Response.Write "<form method=post action=""?action=pagerename"">"
		End If
		Response.Write "<tr onmouseover=""this.style.backgroundColor='#B3CFEE';this.style.color='red'"" onmouseout=""this.style.backgroundColor='';this.style.color=''"">"
		Response.Write "<td height=25 align=left>"
		Response.Write "<li>"
		Select Case i
			Case 2
				Response.Write "当前模板CSS设置<a href=#>"
			Case 3
				Response.Write "当前模板主模块<a href=#>"
			Case Else 
				Response.Write "分页面模板<a href=#>"
		End Select
		
		If i> 21 Then
			
			Response.Write "</a>&nbsp;&nbsp;page_"
			Response.Write "<input type=text size=15 name=newpagename value="&Replace(Rs(i).Name,"page_","")&">&nbsp;&nbsp;"
			Response.Write "<input type=hidden size=15 name=oldpagename value="&Replace(Rs(i).Name,"page_","")&">"		
			Response.Write "<input type=submit value=""分模板页面改名"" name=""mo"" title=""修改名称后提交"">"
		Else
			Response.Write "("&Rs(i).Name&")</a>&nbsp;&nbsp;"		
		End If		
		Response.Write "</td>"
		Response.Write "<td height=""25"" align=""left"">"
		If i=4 Then
			Response.Write "编辑该模块："
			Response.Write "<a href=""admin_bbsface.asp?Stype=1&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>发贴表情</a> <a href=""?action=edit&stype=1&page="					
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""></a> | <a href=""admin_bbsface.asp?Stype=2&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>发贴心情</a> <a href=""?action=edit&stype=2&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new"""
			Response.Write "></a> | <a href=""admin_bbsface.asp?Stype=3&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>用户头像</a> <a href=""?action=edit&stype=3&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""></a>"		
		ElseIf i>2 Then 
			Response.Write "编辑该模块："
			Response.Write "<a href=""?action=edit&stype=1&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>语言包</a> <a href=""?action=edit&stype=1&page="					
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""><font color=blue>新</font></a> | <a href=""?action=edit&stype=2&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>图片</a> <a href=""?action=edit&stype=2&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new"""
			Response.Write "><font color=blue>新</font></a> | <a href=""?action=edit&stype=3&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>界面风格</a> <a href=""?action=edit&stype=3&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""><font color=blue>新</font></a>"
			If i=3 Then 
				Response.Write " | <a href=""?action=editmain&stype=2&StyleID="&StyleID&""">基本设置</a>"		
			End If
		ElseIf i=2 Then
			Response.Write "编辑该模块："
			Response.Write "<a href=""?action=editcss&StyleID="&StyleID&""">修改CSS样式</a>"	 
		End If
		If i >21 Then 
			Response.Write "&nbsp;&nbsp;<a href=""?action=delpage&StylePageName="&Rs(i).Name&""" title=""注意，删除后不可恢复""> 删除分模板页面 </a>"		
		End If
		Response.Write "</td>"
		Response.Write "</tr>"
		If i>21 Then
			Response.Write "</form>" 
		End If
	Next
	Response.Write "</table><p></p>"
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=2 height=25>论坛模板管理"
	Response.Write "</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<FORM METHOD=POST ACTION=""?action=addstyle"">"
	Response.Write "<td class=""forumRowHighlight"" height=40 align=left width=45% >"
	Response.Write "新建模板：&nbsp;"
	Response.Write "<input type=text size=25 name=""StyleName"">&nbsp;&nbsp;"
	Response.Write "<input type=submit value=""添 加"" name=""mostyle"">&nbsp;&nbsp;填写模板名"
	Response.Write "<br><br><b>说明：</b>新建模板将把当前论坛默认模版数据复制到新的模板中"
	Response.Write "</td>"
	Response.Write "</FORM>"
	Response.Write "<FORM METHOD=POST ACTION=""?action=ghost"">"
	Response.Write "<td class=""forumRowHighlight"" height=25 align=left>源模板："
	Response.Write "<select name=""OlDStyleID"" size=1>"
	For i=0 to UBound(Templateslist)
		Response.Write "<option value="""&Templateslist(i)(0)&""""
		If CLng(Templateslist(i)(0)) = CLng(StyleID) Then 
			Response.Write " selected"
		End If 
		Response.Write ">"&Templateslist(i)(1)&"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;目标模板："
	Response.Write "<select name=""newStyleID"" size=1>"
	For i=0 to UBound(Templateslist)
		Response.Write "<option value="""&Templateslist(i)(0)&""""
		If CLng(Templateslist(i)(0)) = CLng(StyleID) Then 
			Response.Write " selected"
		End If 
		Response.Write ">"&Templateslist(i)(1)&"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;&nbsp;<input type=submit name=submit value=""模板克隆"">"
	Response.Write "<br><br><b>说明：</b>模板克隆将用源模版数据覆盖目标模板中的所有数据"
	Response.Write "</td>"
	Response.Write "</FORM>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<FORM METHOD=POST ACTION=""?action=rename"">"
	Response.Write "<td class=""forumRowHighlight"" height=25 align=left>"
	Response.Write "<select name=""StyleID"" size=1>"
	For i=0 to UBound(Templateslist)
		Response.Write "<option value="""&Templateslist(i)(0)&""""
		If CLng(Templateslist(i)(0)) = CLng(StyleID) Then 
			Response.Write " selected"
		End If 
		Response.Write ">"&Templateslist(i)(1)&"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;&nbsp;"
	Response.Write "改名为：<input type=text size=20 name=""StyleName"" value="""
	Response.Write """>&nbsp;&nbsp;"
	Response.Write "<input type=submit name=submit value=""修改"">"
	Response.Write "</td>"
	Response.Write "</FORM>"
	Response.Write "<td class=""forumRowHighlight"" height=25 align=left>"
	Response.Write "&nbsp;"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table><BR>"
	Rs.Close
	Set Rs=Nothing
End Sub
Sub Edit()
	Dim Page,mystr
	Dim TempStr,TemplateStr,stype
	Dim TempStyleHelp,StyleHelpValue
	stype=Dvbbs.checkStr(request("stype"))
	page=Dvbbs.checkStr(request("page"))
	If Not IsNumeric(stype) Then 
		Errmsg=ErrMsg + "<BR><li>错误的样式参数"
		Dvbbs_error()
	End If
	If Not IsTruePage(page) Then
		Errmsg=ErrMsg + "<BR><li>要编辑的页面模板字段尚未建立。"
		Dvbbs_error()
	End If
	Set Rs=Dvbbs.Execute("Select ID,StyleName,"&page&" From [Dv_StyleHelp] where ID=1")
	TempStr=Split(Rs(2),"@@@")
	Select Case stype
		Case 1
			TempStyleHelp=Split(TempStr(1),"|||")
		Case 2
			TempStyleHelp=Split(TempStr(2),"|||")
		Case 3
			TempStyleHelp=Split(TempStr(0),"|||")
	End Select
	Set Rs=Dvbbs.Execute("Select ID,StyleName,"&page&" From [Dv_Style] Where ID="&StyleID)
	TempStr=Split(Rs(2),"@@@")
	Select Case stype
		Case 1
			TemplateStr=Split(TempStr(1),"|||")
		Case 2
			TemplateStr=Split(TempStr(2),"|||")
		Case 3
			TemplateStr=Split(TempStr(0),"|||")
	End Select
	Response.Write "<form name=""template"" action=""?action=saveedit&page="&page&"&stype="&stype&"&StyleID="&StyleID&""" method=post>"
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=3 height=25>"
	Response.Write Rs(1)
	Response.Write "分页面模板("
	Response.Write page
	Response.Write ")"
	Response.Write "<input Type=""hidden"" name=""dvbbs"" value=""OK!"">"
	Select Case stype
		Case 1
			Response.Write "语言包"
			mystr="template.Strings"
			If page="main_style" Then mystr="Dvbbs.lanstr"
		Case 2
			Response.Write "图片资源(当前默认路径{$PicUrl}为："&Dvbbs.Forum_PicUrl&")"
			mystr="template.pic"
			If page="main_style" Then mystr="Dvbbs.mainpic"
		Case 3
			Response.Write "界面风格"
			mystr="template.html"
			If page="main_style" Then mystr="Dvbbs.mainhtml"
	End Select
	
	Response.Write "管理</th></tr>"
	If TemplateStr(Ubound(TemplateStr))="" Then TemplateStr(Ubound(TemplateStr))="del"
	For i=0 To Ubound(TemplateStr)
		If i<ubound(TempStyleHelp) Then
			StyleHelpValue=TempStyleHelp(i)
		Else
			StyleHelpValue="//"
		End IF
		Response.Write "<tr><td class=""forumRowHighlight"" width=20% height=40 align=left>"
		Response.Write mystr&"("&i&")"
		Response.Write "<br><a href=""#"" onclick=""rundvscript(t"&i&",'page="&page&"&TemplatesCount="&i&"&stype="&stype&"');"" title=""点这里获取这部分模板的官方数据"">获取官方数据</a>"
		Response.Write "</td>"		
		Response.Write "<td class=""forumRowHighlight"" width=80% height=25 align=left>"
		Select Case stype
			Case 1
				If LenB(TemplateStr(i))>70 Then
				Response.Write "<textarea name=""TemplateStr"" id=""t"&i&"""  cols=""100"" rows=""3"">"
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write "</textarea>"
				Else
				Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t"&i&""" value="""
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write """ size=50>"
				End If
				Response.Write "<INPUT TYPE=""hidden"" NAME=""ReadME"" id=""r"&i&""" value="""&StyleHelpValue&""">"
				Response.Write "<a href=# onclick=""helpscript(r"&i&");return false;"" class=""helplink""><img src=""images/manage/help.gif"" border=0 title=""点击查阅管理帮助！""></a>"
			Case 2
				Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t"&i&""" value="""
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write """ size=20> "
				If server.htmlencode(TemplateStr(i))<>"" And (Instr(server.htmlencode(TemplateStr(i)),".gif") or Instr(server.htmlencode(TemplateStr(i)),".jpg")) Then Response.Write "<img src="&server.htmlencode(Replace(TemplateStr(i),"{$PicUrl}",Dvbbs.Forum_PicUrl))&"  border=0>"	
			Case 3
				If page="main_style"  And i=0 Then 
					Response.Write "<input type=hidden name=""TemplateStr"" value="""
					Response.Write server.htmlencode(TemplateStr(i))
					Response.Write """>"
					Response.Write "此字段属于基本设置，  <a href=""?action=editmain&stype=2&StyleID="&StyleID&""">点这里修改基本设置</a>"
					Response.Write "</td><td class=""forumRowHighlight"">"
					Response.Write "<a href=# onclick=""helpscript(r"&i&");return false;"" class=""helplink""><img src=""images/manage/help.gif"" border=0 title=""点击查阅管理帮助！""></a>"
				Else
					
					Response.Write "<textarea name=""TemplateStr"" id=""t"&i&""" cols=""100"" rows=""5"">"
					Response.Write server.htmlencode(TemplateStr(i))
					Response.Write "</textarea>"
					Response.Write "</td><td class=""forumRowHighlight""><a href=""javascript:admin_Size(-5,'t"&i&"')""><img src=""images/manage/minus.gif"" unselectable=""on"" border='0'></a> <a href=""javascript:admin_Size(5,'t"&i&"')""><img src=""images/manage/plus.gif"" unselectable=""on"" border='0'></a>"
					Response.Write "<img src=images/manage/viewpic.gif onclick=runscript(t"&i&")>"
					Response.Write "<a href=# onclick=""helpscript(r"&i&");return false;"" class=""helplink""><img src=""images/manage/help.gif"" border=0 title=""点击查阅管理帮助！""></a> "		
				End If
				Response.Write "<INPUT TYPE=""hidden"" NAME=""ReadME"" id=""r"&i&""" value="""&StyleHelpValue&""">"
			End Select
			
		Response.Write "</td></tr>"
	Next
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"" colspan=""3"">"
	Response.Write "</td></tr>"
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"">"
	Response.Write "<input type=""reset"" name=""Submit"" value=""重 填"">"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" colspan=2 align=""center"">"
	Response.Write "<input type=""submit"" name=""B1"" value=""修 改"">"
	Response.Write "</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td colspan=3 Class=""forumRowHighlight"">"
	Response.Write "<br><li>模板编辑规则：如果想清除该字段，请在对应的文本框中输入""del""，那么模板数据的序号就会前移。"
	Response.Write "<br><li>如果不想改变模板数据的序号,仅把该项目的数据清空,则只需要把内容清空。"
	Response.Write "</td></tr>"
	Response.Write "</table><p></p>"
	Response.Write "</form>"
	Rs.Close
	Set Rs=Nothing
End Sub
Sub SaveEdit()
	If Request("dvbbs")<>"OK!" Then
		Errmsg=ErrMsg + "<BR><li>您提交了非法数据"
		Dvbbs_error()
		Exit Sub
	End If
	Dim Page
	Dim TempStr,TemplateStr,stype
	Dim TempStyleHelp,StyleHelpValue
	stype=Dvbbs.checkStr(request("stype"))
	page=Dvbbs.checkStr(request("page"))
	If Not IsNumeric(stype) Then 
		Errmsg=ErrMsg + "<BR><li>错误的样式参数"
		Dvbbs_error()
	End If
	If Not IsTruePage(page) Then
		Errmsg=ErrMsg + "<BR><li>要编辑的页面模板字段尚未建立。"
		Dvbbs_error()
	End If
	'提取表单中的数据
	TemplateStr=""
	For Each TempStr in Request.form("TemplateStr")
		If LCase(TempStr)<>"del" Then 
			TemplateStr=TemplateStr&Replace(TempStr,"|||","")&"|||"
		End If
	Next
	TemplateStr=Dvbbs.checkStr(Replace(TemplateStr,"@@@",""))
	If Trim(TemplateStr)="" Then 
		TemplateStr="|||"
	End If

	'提取表单中的数据
	StyleHelpValue=""
	For Each TempStyleHelp in Request.form("ReadME")
		If TempStyleHelp<>"" Then 
			StyleHelpValue=StyleHelpValue&TempStyleHelp&"|||"
		End If
	Next
	If Trim(StyleHelpValue)="" Then 
		StyleHelpValue="|||"
	Else
		StyleHelpValue=Dvbbs.checkStr(StyleHelpValue)
	End If

	Set Rs=Dvbbs.Execute("Select ID,StyleName,"&page&" From [Dv_Style] Where ID="&StyleID)
	TempStr=Split(Dvbbs.checkStr(Rs(2)),"@@@")
	Select Case stype
		Case 1
			TemplateStr=TempStr(0)&"@@@"&TemplateStr&"@@@"&TempStr(2)
		Case 2
			TemplateStr=TempStr(0)&"@@@"&TempStr(1)&"@@@"&TemplateStr
		Case 3
			TemplateStr=TemplateStr&"@@@"&TempStr(1)&"@@@"&TempStr(2)
	End Select

	Set Rs=Dvbbs.Execute("Select ID,StyleName,"&page&" From [Dv_StyleHelp] Where ID=1")
	TempStr=Split(Dvbbs.checkStr(Rs(2)),"@@@")
	Select Case stype
		Case 1
			StyleHelpValue=TempStr(0)&"@@@"&StyleHelpValue&"@@@"&TempStr(2)
		Case 2
			StyleHelpValue=TempStr(0)&"@@@"&TempStr(1)&"@@@"&StyleHelpValue
		Case 3
			StyleHelpValue=StyleHelpValue&"@@@"&TempStr(1)&"@@@"&TempStr(2)
	End Select
	Rs.close:Set Rs=Nothing
	Dvbbs.Execute("update [Dv_Style] set "&page&"='"&TemplateStr&"' Where ID="&StyleID)
	Dvbbs.Execute("update [Dv_StyleHelp] set "&page&"='"&StyleHelpValue&"' Where ID=1")
	Select Case stype
		Case 1
			Dv_suc(page&"语言包修改成功!")
		Case 2
			Dv_suc(page&"图片资源修改成功!")
		Case 3
			Dv_suc(page&"界面风格修改成功!")
	End Select
	'更新缓存。此处是在模板数据变化的时候需要更新的代码。如有漏掉，可以在这添加。
	Dvbbs.DelCahe(page&StyleID)
	If InStr(Lcase(page),"main_style")>0 Then
		Dvbbs.DelCahe("head_"&StyleID)
		Dvbbs.DelCahe("GroupTitle")
		Dvbbs.DelCahe("ForumPlusMenu"&StyleID)
	End If
End Sub
Sub Edit_Main()
	Dim Page,mystr
	Dim TempStr,TemplateStr,stype
	stype=Dvbbs.checkStr(request("stype"))
	page=Dvbbs.checkStr(request("page"))
	If Not IsNumeric(stype) Then 
		Errmsg=ErrMsg + "<BR><li>错误的样式参数"
		Dvbbs_error()
	End If
	If page<>"page_admin" Then
		Errmsg=ErrMsg + "<BR><li>要编辑的页面模板字段尚未建立。"
		Dvbbs_error()
	End If
	Set Rs=Dvbbs.Execute("Select H_ID,H_Title,H_Content From [Dv_help] Where H_ID=1")
	TempStr=Split(Rs(2),"@@@")
	Select Case stype
		Case 1
			TemplateStr=Split(TempStr(1),"|||")
		Case 2
			TemplateStr=Split(TempStr(2),"|||")
		Case 3
			TemplateStr=Split(TempStr(0),"|||")
	End Select
	Response.Write "<form action=""?action=Save_Main&page="&page&"&stype="&stype&"&StyleID="&StyleID&""" method=post>"
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=3 height=25>"
	Response.Write Rs(1)
	Response.Write "分页面模板("
	Response.Write page
	Response.Write ")"
	Select Case stype
		Case 1
			Response.Write "语言包"
			mystr="template.Strings"
			If page="main_style" Then mystr="Dvbbs.lanstr"
		Case 2
			Response.Write "图片资源"
			mystr="template.pic"
			If page="main_style" Then mystr="Dvbbs.mainpic"
		Case 3
			Response.Write "界面风格"
			mystr="template.html"
			If page="main_style" Then mystr="Dvbbs.mainhtml"
	End Select
	
	Response.Write "管理</th></tr>"
	For i=0 To Ubound(TemplateStr)
		Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=left>"
		Response.Write mystr&"("&i&")"
		Response.Write "</td>"		
		Response.Write "<td class=""forumRowHighlight"" height=25 align=left>"
		Select Case stype
			Case 1
				If LenB(TemplateStr(i))>70 Then
				Response.Write "<textarea name=""TemplateStr"" cols=""100"" rows=""3"">"
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write "</textarea>"
				Else
				Response.Write "<input Type=""text"" name=""TemplateStr"" value="""
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write """ size=50>"
				End If
			Case 2
				Response.Write "<input Type=""text"" name=""TemplateStr"" value="""
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write ""
				Response.Write """ size=20>"
			Case 3
				If page="main_style"  And i=0 Then 
					Response.Write "<input type=hidden name=""TemplateStr"" value="""
					Response.Write server.htmlencode(TemplateStr(i))
					Response.Write """>"
					Response.Write "此字段属于基本设置，  <a href=""?action=editmain&stype=2&StyleID="&StyleID&""">点这里修改基本设置</a>"		
				Else
					
					Response.Write "<textarea name=""TemplateStr"" id=""t"&i&""" cols=""100"" rows=""5"">"
					Response.Write server.htmlencode(TemplateStr(i))
					Response.Write "</textarea>"
					Response.Write "</td><td class=""forumRowHighlight""><a href=""javascript:admin_Size(-5,'t"&i&"')""><img src=""images/manage/minus.gif"" unselectable=""on"" border='0'></a> <a href=""javascript:admin_Size(5,'t"&i&"')""><img src=""images/manage/plus.gif"" unselectable=""on"" border='0'></a><img src=images/manage/viewpic.gif onclick=runscript(t"&i&")>"
				End If 
			End Select
		Response.Write "</td></tr>"
		
	Next
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"" colspan=""3"">"
	Response.Write "</td></tr>"
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"">"
	Response.Write "<input type=""reset"" name=""Submit"" value=""重 填"">"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""center"" colspan=""2"">"
	Response.Write "<input type=""submit"" name=""B1"" value=""修 改"">"
	Response.Write "</td></tr>"
	Response.Write "</table><p></p>"
	Response.Write "</form>"
	
	Rs.Close
	Set Rs=Nothing
End Sub
Sub Save_Main()
	Dim Page
	Dim TempStr,TemplateStr,stype
	stype=Dvbbs.checkStr(request("stype"))
	page=Dvbbs.checkStr(request("page"))
	If Not IsNumeric(stype) Then 
		Errmsg=ErrMsg + "<BR><li>错误的样式参数"
		Dvbbs_error()
	End If
	If page<>"page_admin" Then
		Errmsg=ErrMsg + "<BR><li>要编辑的页面模板字段尚未建立。"
		Dvbbs_error()
	End If
	'提取表单中的数据
	TemplateStr=""
	For Each TempStr in Request.form("TemplateStr")
		If TempStr<>"" Then 
			TemplateStr=TemplateStr&Replace(TempStr,"|||","")&"|||"
		End If
	Next 
	TemplateStr=Dvbbs.checkStr(Replace(TemplateStr,"@@@",""))
	If Trim(TemplateStr)="" Then 
		TemplateStr="|||"
	End If
	Set Rs=Dvbbs.Execute("Select H_ID,H_title,H_content From [Dv_help] Where H_ID=1")
	TempStr=Split(Dvbbs.checkStr(Rs(2)),"@@@")
	Select Case stype
		Case 1
			TemplateStr=TempStr(0)&"@@@"&TemplateStr&"@@@"&TempStr(2)
		Case 2
			TemplateStr=TempStr(0)&"@@@"&TempStr(1)&"@@@"&TemplateStr
		Case 3
			TemplateStr=TemplateStr&"@@@"&TempStr(1)&"@@@"&TempStr(2)
	End Select
	Set Rs=Nothing 
	Dvbbs.Execute("update [Dv_help] set H_content='"&TemplateStr&"' Where H_ID=1")
	Select Case stype
		Case 1
			Dv_suc(page&"语言包修改成功!")
		Case 2
			Dv_suc(page&"图片资源修改成功!")
		Case 3
			Dv_suc(page&"界面风格修改成功!")
	End Select
End Sub
Function IsTruePage(page)
	IsTruePage=False
	If page<>"" Then 
		page=LCase(Trim(page))
		Dim myrs
		Set Myrs=Dvbbs.Execute("Select top 1 * From [Dv_Style]")
		For i= 2 to MyRs.Fields.Count-1
			If LCase( myrs(i).name)=page Then
				 IsTruePage=True
				 Exit For
			End If
		Next
	End If
End Function

Sub DelStyle()
	'检查是否有版面使用本模版
	If StyleID=SkinID Then 
		Errmsg=ErrMsg + "<BR><li>本模板是默认模版，不允许删除。"
		Dvbbs_error()
	End If
	Set Rs=Dvbbs.Execute("select Count(*) From [Dv_Board] Where sid="&StyleID&"")
	If Rs(0)>0 Then
		Set Rs=Nothing 
		Errmsg=ErrMsg + "<BR><li>本模板尚有分论坛在使用，不能删除。"
		Dvbbs_error()
	End If
	Set rs=Nothing 
	Dvbbs.Execute("Delete From [Dv_Style] Where ID="&StyleID&"")
	Dv_suc("成功删除了一个模板。")
	Dvbbs.DelCahe("Templateslist")'删除模板列表缓存
End Sub
Sub delpage()
	Dim StylePageName
	StylePageName=Dvbbs.checkStr(request("StylePageName"))
	If StylePageName="" Then 
		Errmsg=ErrMsg + "<BR><li>请填写字段名称"
		Dvbbs_error()
	End If
	If Not IsTruePage(StylePageName) Then 
		Errmsg=ErrMsg + "<BR><li>要删除的字段不存在。"
		Dvbbs_error()
	End If
	If IsSqlDataBase = 1 Then
		Dim i,Fieldslist,Rs
		Set Rs=Dvbbs.Execute("select * from Dv_Style")
		Fieldslist="id"
		For i= 1 to Rs.Fields.Count-1
			If LCase(Rs(i).name)<> LCase (StylePageName) Then 
				Fieldslist=Fieldslist&","&Rs(i).name
			End If
		Next
		Set Rs=Nothing 
		'复制有用数据到临时表	
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_tempatble From Dv_Style")
		'删除原有表
		Dvbbs.Execute("Drop table Dv_Style")
		'再把临时表中的数据复制过来.
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_Style From Dv_tempatble ")
		'删除临时表
		Dvbbs.Execute("Drop table Dv_tempatble")
		'复制有用数据到临时表	
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_tempatble From Dv_Stylehelp")
		'删除原有表
		Dvbbs.Execute("Drop table Dv_Stylehelp")
		'再把临时表中的数据复制过来.
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_Stylehelp From Dv_tempatble ")
		'删除临时表
		Dvbbs.Execute("Drop table Dv_tempatble") 
	Else
		Dvbbs.Execute("Alter Table [Dv_Style] Drop ["&StylePageName&"]")
		Dvbbs.Execute("Alter Table [Dv_Stylehelp] Drop ["&StylePageName&"]")
	End If 
	Dv_suc("页面模板删除成功！")
End Sub 
Sub addpage()
	Dim StylePageName
	StylePageName=Dvbbs.checkStr(request("StylePageName"))
	If StylePageName="" Then 
		Errmsg=ErrMsg + "<BR><li>请填写字段名称"
		Dvbbs_error()
	End If
	StylePageName="Page_"&StylePageName
	If IsTruePage(StylePageName) Then 
		Errmsg=ErrMsg + "<BR><li>您要创建的模板字段已经存在。"
		Dvbbs_error()
	End If
	Dvbbs.Execute("alter table [Dv_Stylehelp] add "&StylePageName&" text not Null default'|||@@@|||@@@|||@@@|||'")
	Dvbbs.Execute("Update [Dv_Stylehelp] Set "&StylePageName&"='|||@@@|||@@@|||@@@|||'")
	Dvbbs.Execute("alter table [Dv_Style] add "&StylePageName&" text not Null default'|||@@@|||@@@|||@@@|||'")
	Dvbbs.Execute("Update [Dv_Style] Set "&StylePageName&"='|||@@@|||@@@|||@@@|||'")
	Dv_suc("新页面模板创建成功!")
End Sub

Sub addstyle()
	Dim stylename,sql
	stylename=Dvbbs.checkStr(Request("stylename"))
	If Trim(stylename)=""  Then 
		 Errmsg=ErrMsg + "<BR><li>请输入模板名称。"
		Dvbbs_error()
	End If
	Set Rs=Dvbbs.Execute("select * From [Dv_Style] where ID="&StyleID&"")
	Dim styleFields,stylevalues
	styleFields="StyleName"
	stylevalues="'"&stylename&"'"
	For i= 2 to Rs.Fields.Count-1
		styleFields=styleFields&","&Rs(i).Name
		stylevalues=stylevalues&",'"&Dvbbs.checkStr(Rs(i))&"'"
	Next
	Set rs=Nothing
	sql="insert into [Dv_Style]("&styleFields&")values("&stylevalues&")"
	Dvbbs.Execute(SQL)
	Dv_suc("新模板创建成功!")
	Dvbbs.DelCahe("Templateslist")'删除模板列表缓存
End Sub
Sub pagerename()
	Dim oldpagename,newpagename
	oldpagename=Dvbbs.checkStr(request("oldpagename"))
	newpagename=Dvbbs.checkStr(request("newpagename"))
	If LCase(newpagename)=LCase(oldpagename) Then
		Errmsg=ErrMsg + "<BR><li>你没有更改名称"
		Dvbbs_error()
	End If
	If newpagename="" Then
		Errmsg=ErrMsg + "<BR><li>新名称不能为空"
		Dvbbs_error()
	End If
	If OLDpagename="" Then
		Errmsg=ErrMsg + "<BR><li>您提交的数据是错误的."
		Dvbbs_error()
	End If
	oldpagename="page_"&oldpagename
	newpagename="page_"&newpagename
	If Not IsTruePage(oldpagename) Then 
		Errmsg=ErrMsg + "<BR><li>要改名的字段不存在。"
		Dvbbs_error()
	End If
	If IsTruePage(newpagename) Then 
		Errmsg=ErrMsg + "<BR><li>字段名称"&newpagename&"已经被占用."
		Dvbbs_error()
	End If
		Dim i,Fieldslist,Rs
		Set Rs=Dvbbs.Execute("select * from Dv_Style")
		Fieldslist="id"
		For i= 1 to Rs.Fields.Count-1
			If LCase(Rs(i).name)<> LCase (oldpagename) Then 
				Fieldslist=Fieldslist&","&Rs(i).name
			Else
				Fieldslist=Fieldslist&","&Rs(i).name&" as "& newpagename
			End If
		Next
		Set Rs=Nothing
		'复制有用数据到临时表	
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_tempatble From Dv_Style")
		'删除原有表
		Dvbbs.Execute("Drop table Dv_Style")
		'再把临时表中的数据复制过来.
		Dvbbs.Execute("Select * into Dv_Style From Dv_tempatble ")
		'删除临时表
		Dvbbs.Execute("Drop table Dv_tempatble")
		'复制有用数据到临时表	
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_tempatble From Dv_Stylehelp")
		'删除原有表
		Dvbbs.Execute("Drop table Dv_Stylehelp")
		'再把临时表中的数据复制过来.
		Dvbbs.Execute("Select * into Dv_Stylehelp From Dv_tempatble ")
		'删除临时表
		Dvbbs.Execute("Drop table Dv_tempatble") 
	Dv_suc("成功把 "&oldpagename&" 字段改名为 "&newpagename&"")
End Sub
Sub rename()
	Dim stylename
	stylename=Dvbbs.checkStr(Request("stylename"))
	If Trim(stylename)=""  Then 
		Errmsg=ErrMsg + "<BR><li>修改名称请输入新的模板名称。"
		Dvbbs_error()
	End If
	Dvbbs.Execute("update [Dv_Style] set StyleName='"&StyleName&"' where id="&StyleID&"")
	Dv_suc("模板名修改成功!")
	Dvbbs.DelCahe("Templateslist")'删除模板列表缓存
End Sub

'//编辑修改CSS模板
'CSS风格名称@@@CSS风格内容@@@CSS风格图片包路径	（用@@@分类，用|||分隔内容.）
Sub EditCss()
Dim CssName,CssContent,CssPicUrl,MainStyleName,k,CssCount
Dim TempCssStyle,LockID
Set Rs=Dvbbs.Execute("Select Forum_CSS,StyleName From [Dv_Style] Where ID="&StyleID)
	TempCssStyle=Split(Server.htmlencode(Rs(0)),"@@@")
	MainStyleName=Rs(1)
Rs.close

CssName=Split(TempCssStyle(0),"|||")	'CSS风格名称
CssCount=Ubound(CssName)
If Ubound(TempCssStyle)<1 Then
Redim CssContent(CssCount)
Else
CssContent=Split(TempCssStyle(1),"|||")	'CSS风格内容
End If
If Ubound(TempCssStyle)<2 Then
Redim CssPicUrl(CssCount)
Else
CssPicUrl=Split(TempCssStyle(2),"|||")	'CSS风格图片包路径
End IF

Response.Write "<br><table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
Response.Write "<tr>"
Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=2 height=25>（"
Response.Write MainStyleName
Response.Write "）-CSS样式管理"
Response.Write "</th></tr></table>"
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function EditCss(n)
{
	var trid=document.getElementsByName('CssTR');
	 for (var i=0;i<trid.length;i++)    {
		if (i!=n){
		trid[i].style.display="none";
		}
	   }
	EditTextarea.style.display = '';
	document.cssform.CssContent.value = document.cssform.CssBody[n].value;
	document.cssform.TempID.value = n;
	document.cssform.CssEdit[n].disabled=true;
}

function DllData(n){
	if (document.cssform.DelCss[n].checked==true){
	document.cssform.CssName[n].value = '计划将删除...';
	document.cssform.CssBody[n].disabled = true;
	document.cssform.CssPicUrl[n].disabled = true;
	document.cssform.CssEdit[n].disabled=true;
	}else{
	document.cssform.CssName[n].value = '请重新输入新的风格名称!!';
	document.cssform.CssBody[n].disabled = false;
	document.cssform.CssPicUrl[n].disabled = false;
	document.cssform.CssEdit[n].disabled=false;
	}
}

function SubmitData(){
	var NewData,UpObject
	var e = document.cssform;
	NewData=e.CssContent.value;
	UpObject=e.TempID.value;
	if (NewData!=''){
		e.CssBody[UpObject].value=NewData;
	}
	for (var i=0;i<e.CssName.length;i++){
		if (e.CssName[i].value == '计划将删除...' || e.CssName[i].value == ''){
		e.CssName[i].value = '';
		e.CssBody[i].value = '';
		e.CssPicUrl[i].value = '';
		}
	}
}
//-->
</SCRIPT>
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder" >
<form action="?action=savecss&StyleID=<%=StyleID%>" method="post" name="cssform" onsubmit="SubmitData();">
<tr>
<td width="10%" class="bodytitle" align=center>ID</td>
<td width="30%" class="bodytitle" align=center>名称</td>
<td width="30%" class="bodytitle" align=center>图片包路径(请以“/”结束)</td>
<td width="20%" class="bodytitle" align=center>操作</td>
<td width="10%" class="bodytitle" align=center>删除</td>
</tr>
<tr><td height=1 class="tableBorder1" colspan=5></td></tr>
<%
For k=0 To Ubound(CssName)
%>
<tr Name="CssTR" id="CssTR">
<td class="forumRowHighlight" align=center><%=k%></td>
<td class="forumRow" align=center><input type=text value="<%=CssName(k)%>" name="CssName"></td>
<td class="forumRowHighlight" align=center>
<%
Response.Write "<input name=""CssPicUrl"" size=30 type=text value="""
if CssPicUrl(k)="" and k<>Ubound(CssName) Then Response.Write "Skins/Default/" Else Response.Write CssPicUrl(k)
Response.Write """>"
%>
</td>
<td class="forumRow" align=center>
<input type="button" value="修改Css内容" name="CssEdit" onclick="EditCss(<%=k%>)">
<div style="display:none">
<textarea name="CssBody" id="CssBody" style="width:0;height=0" rows="0" ><%=CssContent(k)%></textarea>
</div>
</td>
<td class="forumRowHighlight" align=center><INPUT TYPE="checkbox" NAME="DelCss" onclick="DllData(<%=k%>)" ></td>
</tr>
<%
Next
%>
<tr id="EditTextarea" style="display:none">
<INPUT TYPE="hidden" NAME="TempID">
<td height=400 class="forumRowHighlight" colspan=5>
<textarea id="CssContent" style="width:100%" rows="30" ></textarea>
</td>
</tr>
<%
Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"" colspan=5>"
Response.Write "<input type=""submit"" name=""B1"" value=""提 交"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.Write "<input type=""button"" value=""返 回"" onclick=""location.reload('admin_template.asp?action=editcss&StyleID="&StyleID&"')"">"
Response.Write "</td></tr>"
%>
</form></table>
<%
End Sub

'//保存修改CSS模板
Sub savecss()
	If StyleID="" or not isnumeric(StyleID) Then
		Errmsg=ErrMsg + "<BR><li>请选择您要修改的CSS样式!"
		Dvbbs_error()
		Exit Sub 
	End If
	Dim TempStr,LockID
	Dim TempStr1,TempStr2,TempStr3
	Dim CssName,CssContent,CssPicUrl,TempCssStyle
	LockID=1
	'提取CSS风格名的数据
	For Each TempStr in Request.Form("CssName")
		If TempStr<>"" Then 
			CssName=CssName&Replace(TempStr,"|||","")&"|||"
		End If
	Next
	'提取图片包的数据
	For Each TempStr in Request.Form("CssPicUrl")
		If TempStr<>"" Then 
			CssPicUrl=CssPicUrl&Replace(TempStr,"|||","")&"|||"
		End If
	Next
	'提取风格内容的数据
	For Each TempStr in Request.Form("CssBody")
		If TempStr<>"" Then 
			CssContent=CssContent&Replace(TempStr,"|||","")&"|||"
		End If
	Next
	CssName=Replace(CssName,"@@@","")
	CssContent=Replace(CssContent,"@@@","")
	CssPicUrl=Replace(CssPicUrl,"@@@","")
	TempStr1=ubound(Split(CssName,"|||"))
	TempStr2=ubound(Split(CssContent,"|||"))
	TempStr3=ubound(Split(CssPicUrl,"|||"))

	If TempStr1<>TempStr2 or TempStr1<>TempStr3 or TempStr2<>TempStr3 Then
		Errmsg=ErrMsg + "<BR><li>图片路径或Css内容不能为空,若要删除Css风格请选取删除选项 !"
		Dvbbs_error()
		Exit Sub 
	End If
	'取出正被使用的CSS风格ID下标界
	Set Rs=Conn.Execute("Select max(Cid) From Dv_Board Where sid="&StyleID)
	If Not Rs.eof Then
		LockID=Rs(0)
	End If
	Rs.close:Set Rs=Nothing
	If TempStr1-1<LockID Then
		Errmsg=ErrMsg + "<BR><li>部分Css风格风格正被使用，删除被中止，请在版块取消选择再进行删除 !"
		Dvbbs_error()
		Exit Sub
	End If

	TempCssStyle = CssName &"@@@"& CssContent &"@@@"& CssPicUrl
	TempCssStyle = Dvbbs.checkStr(TempCssStyle)
	Dvbbs.Execute("Update [Dv_Style] set Forum_CSS='"&TempCssStyle&"' Where ID="&StyleID&"")
	Dv_suc("Css样式修改成功！")
	Dvbbs.DelCahe("Forum_CSS"&StyleID)'删除模板缓存
End Sub

Sub editmain()
	Dim stype,NowEditinfo
	Dim mystr
	stype=Request("stype")
	
	Select Case stype
		Case "1"
			NowEditinfo="语言包"
			mystr="Dvbbs.lanstr"
		Case "2"
			NowEditinfo="基本设置"
			mystr="mainsetting"
		Case "3"
			NowEditinfo="HTTP头部分"
			mystr="mainhtml(0)"
		Case "4"
			NowEditinfo="页面开始部分"
			mystr="mainhtml(1)"
		Case "5"
			NowEditinfo="顶部通栏"
			mystr="mainhtml(2)"
		Case "6"
			NowEditinfo="顶部表格"
			mystr="mainhtml(3)"
		Case "7"
			NowEditinfo="导航栏"
			mystr="mainhtml(5)"
		Case "8"
			NowEditinfo="论坛菜单"
			mystr="mainhtml(6)"
		Case "9"
			mystr="mainhtml(4)"
			NowEditinfo="结束部分"
		Case "10"
			mystr="mainpic"
			NowEditinfo="图片设置"
		Case Else
			Errmsg=ErrMsg + "<BR><li>您提交了错误的参数."
			Dvbbs_error()	
	End Select
	Set Rs=Dvbbs.Execute("Select Main_Style ,StyleName From [Dv_Style] Where ID="&StyleID&"")
	Dim TemplateStr
	TemplateStr=Split(server.htmlencode(Rs(0)),"@@@")
	Response.Write "<form action=""?action=savemain&stype="&stype&"&StyleID="&StyleID&""" method=post>"
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=2 height=25>"
	Response.Write Rs(1)
	Response.Write NowEditinfo
	Response.Write "("&mystr&")设置</th></tr>"
	Select Case stype
		Case "1"
			TemplateStr(0)=split(TemplateStr(0),"|||")
			For i=0 to UBound(TemplateStr(0))
				Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""left"" colspan=2>"
				Response.Write mystr&"("&i&")&nbsp;&nbsp;&nbsp;"		
				Response.Write "<input type=text name=""TemplateStr"" value="""&TemplateStr(0)(i)&""" Size=100>"
				Response.Write "</td></tr>"
			Next	
		Case "2"		
			TemplateStr(0)=split(TemplateStr(0),"|||")
			TemplateStr(0)(0)=split(TemplateStr(0)(0),"||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<table cellspacing=""1"" cellpadding=""0"" border=""0"" align=""left"" width=""100%"" >"
			Response.Write "<tr>"
			Response.Write "<td width=""300"" align=""right"" >表格宽度：</td>"
			Response.Write "<td align=""left"" class=""forumRowHighlight"" colspan=""3"" >"
			Response.Write "<input type=""text"" size=""5"" name=""TemplateStr"" value="""&TemplateStr(0)(0)(0)&""">&nbsp;"&mystr&"(0)"
			Response.Write "</td>"
			Response.Write "</tr>"
			Dim j,vtitle
			vtitle="aa|警告提醒语句的颜色：|显示帖子的时候，相关帖子，转发帖子，回复等的颜色：|首页连接颜色：|一般用户名称字体颜色：|一般用户名称上的光晕颜色：|版主名称字体颜色：|版主名称上的光晕颜色：|管理员名称字体颜色：|管理员名称上的光晕颜色：|贵宾名称字体颜色：|贵宾名称上的光晕颜色：|表格边框颜色："
			vtitle=split(vtitle,"|")
			For j=1 to 12
				Response.Write "<tr>"
				Response.Write "<td height=""1"" colspan=""4"" bgcolor=""#6595D6""><td>"
				Response.Write "</tr>"
				Response.Write "<tr>"
				Response.Write "<td height=""25"" width=""300"" align=""right"" >"&vtitle(j)&"</td>"
				Response.Write "<td width=""20"" bgcolor="""&TemplateStr(0)(0)(j)&"""></td>"
				Response.Write "<td width=""180"" align=""left"">"
				Response.Write "<input type=""text"" size=""10"" value="""&TemplateStr(0)(0)(j)&""" name=""TemplateStr"">"&mystr&"("&j&")"
				Response.Write "</td>"
				Response.Write "<td></td>"
				Response.Write "</tr>"
			Next
			Response.Write "</table>"		
			Response.Write "</td></tr>"
			
		Case "3"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(0)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "4"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(1)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "5"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(2)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "6"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(3)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "7"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(5)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "8"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			TemplateStr(2)(6)=split(TemplateStr(2)(6),"###")
			For i=0 to UBound(TemplateStr(2)(6))
				Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
				Select Case i
					Case 0
						Response.Write "已登录用户菜单代码<br>"		
						Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
						Response.Write TemplateStr(2)(6)(i)
						Response.Write "</textarea>"
						Response.Write "</td></tr>"
					Case 1
						Response.Write "未登录用户菜单代码<br>"		
						Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
						Response.Write TemplateStr(2)(6)(i)
						Response.Write "</textarea>"
						Response.Write "</td></tr>"
					Case 2 
						Response.Write "扩展菜单代码<br>"		
						Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
						Response.Write TemplateStr(2)(6)(i)
						Response.Write "</textarea>"
						Response.Write "</td></tr>"
					Case 3
						Response.Write "菜单分隔图片设置：&nbsp;&nbsp;"		
						Response.Write "<input type=""text"" size=""20"" value="""&TemplateStr(2)(6)(i)&""" name=""TemplateStr"">&nbsp;"
						Response.Write "</td></tr>"
				End Select 
			Next		
		Case "9"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(4)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "10"
			TemplateStr(3)=split(TemplateStr(3),"|||")
			For i=0 to UBound(TemplateStr(3))
				Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""left"" colspan=2>"
				Response.Write mystr&"("&i&")&nbsp;&nbsp;&nbsp;"		
				Response.Write "<input type=text name=""TemplateStr"" value="""&TemplateStr(3)(i)&""" Size=100>"
				Response.Write "</td></tr>"
			Next	
		Case Else
				
	End Select
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"">"
	Response.Write "<input type=""reset"" name=""Submit"" value=""重 填"">"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""center"">"
	Response.Write "<input type=""submit"" name=""B1"" value=""修 改"">"
	Response.Write "</table>"
	Response.Write "</form>"
	Set Rs=Nothing 
End Sub 
Sub savemain()
	Dim stype,NowEditinfo,TemplateStr,tempstr,Main_Style
	stype=Request("stype")
	TemplateStr=""
	Set Rs=Dvbbs.Execute("Select Main_Style From [Dv_Style] Where ID="&StyleID&"")
	Main_Style=Rs(0)
	Set rs=Nothing 
	Main_Style=split(Main_Style,"@@@")
	Select Case stype
		Case "1"
			NowEditinfo="语言包"
			For Each TempStr in Request.form("TemplateStr")
				If TempStr<>"" Then 
					TemplateStr=TemplateStr&TempStr&"|||"
				End If
			Next 
			TemplateStr=TemplateStr&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)&"@@@"&Main_Style(3)
			Exit Sub 
		Case "2"
			NowEditinfo="基本设置"
			For Each TempStr in Request.form("TemplateStr")
				If TempStr<>"" Then
					TemplateStr=TemplateStr&TempStr&"||"
				Else
					TemplateStr=TemplateStr&Chr(1)&"||"
				End If
			Next	
			TemplateStr=Left(TemplateStr,Len(TemplateStr)-2)
			Main_Style(0)=split(Main_Style(0),"|||")
			Dim i
			For i=1 to UBound(Main_Style(0))
				TemplateStr=TemplateStr&"|||"&Main_Style(0)(i)
			Next		
			TemplateStr=TemplateStr&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)'&"@@@"&Main_Style(3)
		Case "3"
			NowEditinfo="HTTP头部分"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&TemplateStr&"|||"&Main_Style(2)(1)&"|||"&Main_Style(2)(2)&"|||"&Main_Style(2)(3)&"|||"&Main_Style(2)(4)&"|||"&Main_Style(2)(5)&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
			Exit Sub 
		Case "4"
			NowEditinfo="页面开始部分"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&TemplateStr&"|||"&Main_Style(2)(2)&"|||"&Main_Style(2)(3)&"|||"&Main_Style(2)(4)&"|||"&Main_Style(2)(5)&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
			Exit Sub 
		Case "5"
			NowEditinfo="顶部通栏"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&Main_Style(2)(1)&"|||"&TemplateStr&"|||"&Main_Style(2)(3)&"|||"&Main_Style(2)(4)&"|||"&Main_Style(2)(5)&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
			Exit Sub 
		Case "6"
			NowEditinfo="顶部表格"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&Main_Style(2)(1)&"|||"&Main_Style(2)(2)&"|||"&TemplateStr&"|||"&Main_Style(2)(4)&"|||"&Main_Style(2)(5)&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
		Case "7"
			NowEditinfo="导航栏"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&Main_Style(2)(1)&"|||"&Main_Style(2)(2)&"|||"&Main_Style(2)(3)&"|||"&Main_Style(2)(4)&"|||"&TemplateStr&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
		Case "8"
			NowEditinfo="论坛菜单"
			For Each TempStr in Request.form("TemplateStr")
				TemplateStr=TemplateStr&TempStr&"###"
			Next 
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&Main_Style(2)(1)&"|||"&Main_Style(2)(2)&"|||"&Main_Style(2)(3)&"|||"&Main_Style(2)(4)&"|||"&Main_Style(2)(5)&"|||"&TemplateStr&"@@@"&Main_Style(3)
			Exit Sub 
		Case "9"
			NowEditinfo="结束部分"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&Main_Style(2)(1)&"|||"&Main_Style(2)(2)&"|||"&Main_Style(2)(3)&"|||"&TemplateStr&"|||"&Main_Style(2)(5)&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
		Case "10"
			NowEditinfo="图片设置"
			For Each TempStr in Request.form("TemplateStr")
				If TempStr<>"" Then 
					TemplateStr=TemplateStr&TempStr&"|||"
				End If
			Next 
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)&"@@@"&TemplateStr
			Exit Sub 
		Case Else
			Errmsg=ErrMsg + "<BR><li>您提交了错误的参数."
			Dvbbs_error()	
	End Select
	TemplateStr=Dvbbs.checkStr(TemplateStr)
	'Response.Write TemplateStr
	Dvbbs.Execute("Update [Dv_Style] set Main_Style='"&TemplateStr&"' Where ID="&StyleID&"")
	Dv_suc("主模板"&NowEditinfo&"修改成功!")
	Dvbbs.DelCahe "Main_Style"&StyleID 	
End Sub  
Sub ghost()
	Dim oldStyleID,newStyleID
	oldStyleID=Request("oldStyleID")
	newStyleID=Request("newStyleID")
	If Not IsNumeric(newStyleID) or Not IsNumeric(oldStyleID) Then
		Errmsg=ErrMsg + "<BR><li>参数错误。"
		Dvbbs_error()
		Exit Sub
	End If
	If newStyleID="" Or oldStyleID="" Then
		Errmsg=ErrMsg + "<BR><li>参数错误。"
		Dvbbs_error()
		Exit Sub
	End If
	oldStyleID=CLng(oldStyleID)
	newStyleID=CLng(newStyleID)
	If newStyleID =	oldStyleID Then 
		Errmsg=ErrMsg + "<BR><li>目标模板和源模板不能相同。"
		Dvbbs_error()
		Exit Sub
	End If
	Set Rs=Dvbbs.Execute("select * From [Dv_Style] where ID="&oldStyleID&"")
	If Rs.EOF Or Rs.BOF Then
		Errmsg=ErrMsg + "<BR><li>无法取出源模版数据"
		Dvbbs_error()
		Exit Sub
	End If
	Dim SQLSTR
	For i= 2 to Rs.Fields.Count-1
		SQLSTR=SQLSTR&Rs(i).Name&"='"&Dvbbs.checkStr(Rs(i))&"'"
		If i<> Rs.Fields.Count-1 Then 
			SQLSTR=SQLSTR&","
		End If
		Dvbbs.DelCahe(Rs(i).Name&newStyleID)
		Dvbbs.DelCahe("head_"&newStyleID)
		Dvbbs.DelCahe("ForumPlusMenu"&newStyleID)
		Dvbbs.DelCahe("BbsListTop"&newStyleID)
	Next 
	Dvbbs.DelCahe("BbsListTop"&newStyleID)
	Dvbbs.Execute("Update [Dv_Style] set "&SQLSTR&" where ID="&newStyleID&" ")
	Dv_suc("模板克隆成功!")
	
End Sub 
%>