<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<!--#include file="inc/ubblist.asp"-->
<%
Dvbbs.LoadTemplates("postjob")
Dim abgcolor
Dim canpostann,caneditann
Dim canmodifyusername
Dim username
Dim rs,sql
canpostann=false
caneditann=false
canmodifyusername=false
Dvbbs.stats=Template.Strings(10)
Dvbbs.ShowErr()
Dim dv_ubb
Dim EmotPath
EmotPath=Split(Dvbbs.Forum_emot,"|||")(0)		'em心情路径
Set dv_ubb=new Dvbbs_UbbCode
If Request("action")<>"showone" Then
	Dvbbs.Nav()
	If Dvbbs.Boardid=0 then
		dvbbs.Head_var 0,0,"",""
		'把版主权限去掉，这样版主就在首页就发不了公告了。
		dvbbs.boardmaster=False
	Else
		Dvbbs.head_var 1,Dvbbs.Board_Data(4,0),"",""
	End if
Else
	Dvbbs.head()
	Response.Write "<br>"
End If
If (dvbbs.master or dvbbs.superboardmaster or dvbbs.boardmaster) and Cint(Dvbbs.GroupSetting(25))=1 Then
	canpostann=True
Else
	canpostann=False
End If
If Dvbbs.FoundUserPer and Cint(Dvbbs.GroupSetting(25))=1 Then
	canpostann=True
ElseIf dvbbs.FoundUserPer and Cint(Dvbbs.GroupSetting(25))=0 Then
	canpostann=False 
End If
If (dvbbs.master or dvbbs.superboardmaster or dvbbs.boardmaster) and Cint(Dvbbs.GroupSetting(26))=1 Then
	caneditann=True 
Else
	caneditann=False
End If
If dvbbs.FoundUserPer and Cint(Dvbbs.GroupSetting(26))=1 Then
	caneditann=true
ElseIf dvbbs.FoundUserPer and Cint(Dvbbs.GroupSetting(26))=0 Then
	caneditann=false
End If
If canpostann or caneditann Then
	response.write Replace(Template.Strings(11),"{$boardid}",Dvbbs.boardid)
End If
If request("action")="AddAnn" Then
	Call addann()
ElseIf request("action")="SaveAnn" Then
	Call saveann()
ElseIf request("action")="EditAnn" Then 
	Call editann()
ElseIf request("action")="EditAnnInfo" Then
	Call EditAnnInfo()
ElseIf request("action")="SaveEdit" Then
	Call SaveEdit()
ElseIf request("action")="delann" Then
	Call delann()
Else
	Call main()
end if
Dvbbs.ShowErr()
Dvbbs.activeonline()
Set dv_ubb=Nothing
Dvbbs.Footer()

Sub main()
	Dim Tempwrite,i
	Response.Write Replace(Template.Strings(12),"{$boardid}",Dvbbs.boardid)
	If Request("action")="showone" Then
		sql="select top 1 title,content,username,addtime,bgs from Dv_bbsnews where boardid="&Dvbbs.BoardID&" order by id desc"
	Else
		sql="select title,content,username,addtime,bgs from Dv_bbsnews where boardid="&Dvbbs.BoardID&" order by id desc"
	End If
	
	Set rs=Dvbbs.execute(sql)
	If rs.eof and rs.bof then
		Tempwrite=Template.html(8)
		Tempwrite=Replace(Tempwrite,"{$title}",Template.Strings(13))
		Tempwrite=Replace(Tempwrite,"{$content}",Template.Strings(14))
		Tempwrite=Replace(Tempwrite,"{$username}",Template.Strings(15))
		Tempwrite=Replace(Tempwrite,"{$addtime}",Now())
		Tempwrite=Replace(Tempwrite,"{$bgs}","No")
		Response.Write Tempwrite
	Else
		Sql=Rs.GetRows(-1)
		For i=0 to Ubound(sql,2)
			Tempwrite=Tempwrite&Template.html(8)
			Tempwrite=Replace(Tempwrite,"{$title}",Dv_FilterJS(Sql(0,i)))
			ubblists=ubblist(Sql(1,i))&"39,"
			Tempwrite=Replace(Tempwrite,"{$content}",dv_ubb.Dv_UbbCode(Sql(1,i),Dvbbs.UserGroupID,2,1))
			Tempwrite=Replace(Tempwrite,"{$username}",Dvbbs.HtmlEnCode(Sql(2,i)))
			REM 修正显示公告时间为NULL值时出错 2004-6-1 Dv.Yz
			If Isdate(Sql(3,i)) Then
				Tempwrite=Replace(Tempwrite,"{$addtime}",Sql(3,i))
			Else
				Tempwrite=Replace(Tempwrite,"{$addtime}",Now())
			End If
			If Sql(4,i)="" or Isnull(Sql(4,i)) then
				Tempwrite=Replace(Tempwrite,"{$bgs}","No")
			Else
				If Request("action")="showone" Then
					Tempwrite=Replace(Tempwrite,"{$bgs}","<img src=Skins/Default/filetype/mid.gif border=0><bgsound src="&Sql(4,i)&" border=0>")
				Else
					Tempwrite=Replace(Tempwrite,"{$bgs}","Yes")
				End if
			End if
		Next
		Response.Write Tempwrite
	End if
	Rs.close:set rs=nothing
End Sub

Sub AddAnn()
	Dim Tempwrite,Boardlist,Readme,i
	If not canpostann then
		Dvbbs.AddErrCode(28)
		Exit sub
	End if
	Sql="select boardid,boardtype from Dv_board"
	Set rs=Dvbbs.execute(sql)
	If Rs.eof and Rs.bof Then
		Boardlist=""
	Else
		Sql=Rs.GetRows(-1)
		For i=0 to Ubound(Sql,2)
			Boardlist=Boardlist&"<option value='"&Sql(0,i)&"'"
			If Dvbbs.BoardID=Clng(Sql(0,i)) then Boardlist=Boardlist&" selected "
			Boardlist=Boardlist&">"&Sql(1,i)&"</option>"
		Next
	End if
	Rs.close:Set rs=nothing
	If Dvbbs.boardmaster Then
		Readme=""
	Else
		Readme=Template.Strings(16)
	End if
	Tempwrite=Template.html(9)
	Tempwrite=Replace(Tempwrite,"{$username}",Dvbbs.membername)
	Tempwrite=Replace(Tempwrite,"{$boardlist}",Boardlist)
	Tempwrite=Replace(Tempwrite,"{$readme}",Readme)
	Tempwrite=Replace(Tempwrite,"{$title}","")
	Tempwrite=Replace(Tempwrite,"{$content}","")
	Tempwrite=Replace(Tempwrite,"{$action}","?action=SaveAnn")
	Tempwrite=Replace(Tempwrite,"{$dowhat}",Template.Strings(23))
	Tempwrite=Replace(Tempwrite,"{$bgs}","")
	Response.Write Tempwrite
End sub

Sub SaveAnn()
	If not Canpostann then
		Dvbbs.AddErrCode(28)
		Exit sub
	End if
	Dim username,title,content,bgs
	If request("username")="" then
		Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(17)&"&action=OtherErr"
	Else
		username=Dvbbs.CheckStr(request("username"))
	End if
	If request("title")="" then
		Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(18)&"&action=OtherErr"
	Else
		title=Dvbbs.CheckStr(request("title"))
	End if
	If request("content")="" then
		Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(19)&"&action=OtherErr"
	Else
		content=Dvbbs.CheckStr(request("content"))
	End If
	bgs=Dv_FilterJS(request("bgs"))

	Set rs=server.createobject("adodb.recordset")
	Sql="select * from Dv_bbsnews"
	If Not IsObject(Conn) Then ConnectionDatabase
	rs.open sql,conn,1,3
	rs.addnew
	rs("username")=fixjs(username)
	rs("title")=fixjs(title)
	rs("content")=content
	rs("addtime")=Now()
	rs("boardid")=Dvbbs.BoardID
	rs("bgs")=bgs
	rs.update
	rs.close:Set rs=Nothing
	If Dvbbs.BoardID>0 Then 
		Dvbbs.Name = "BoardInfo_" & Dvbbs.BoardID
		Dvbbs.LoadBoardNews_Paper(Dvbbs.BoardID)
	Else
		Dvbbs.DelCahe "news"&Dvbbs.boardid
	End If
	Dvbbs.Dvbbs_suc("<li>"&Template.Strings(20))
End sub

Sub EditAnn()
	Dim Tempwrite,Newslist,i
	If not caneditann then
		Dvbbs.AddErrCode(28)
		Exit sub
	End if
	If Dvbbs.BoardID=0 then
		Set rs=Dvbbs.execute("select id,boardid,title,username,addtime,bgs from Dv_bbsnews order by addtime desc")
	Else
		Set rs=Dvbbs.execute("select id,boardid,title,username,addtime,bgs from Dv_bbsnews where boardid="&Dvbbs.BoardID&" order by addtime desc")
	End if
	If Rs.eof and Rs.bof Then
		Newslist=Template.Strings(21)
	Else
		Sql=Rs.GetRows(-1)
		For i=0 To Ubound(Sql,2)
			'修复以往公告的错误。
			If isnull(Sql(1,i)) Then Dvbbs.execute("update Dv_bbsnews set boardid=0 where boardid is null")
			Newslist=Newslist&Template.html(11)
			Newslist=Replace(Newslist,"{$boardid}",Sql(1,i)&"")
			Newslist=Replace(Newslist,"{$id}",Sql(0,i))
			Newslist=Replace(Newslist,"{$title}",Dv_FilterJS(Sql(2,i)))
			Newslist=Replace(Newslist,"{$username}",Dvbbs.HtmlEnCode(Sql(3,i)))
			REM 修正显示公告时间为NULL值时出错 2004-6-1 Dv.Yz
			If Isdate(Sql(4,i)) Then
				Newslist=Replace(Newslist,"{$addtime}",Sql(4,i))
			Else
				Newslist=Replace(Newslist,"{$addtime}",Now())
			End If
			Newslist=Replace(Newslist,"{$bgs}",Dv_FilterJS(Sql(5,i)))
		Next
	End if
	Rs.close:set rs=nothing
	Tempwrite=Template.html(10)
	Tempwrite=Replace(Tempwrite,"{$boardid}",Dvbbs.Boardid)
	Tempwrite=Replace(Tempwrite,"{$newslist}",Newslist)
	Response.Write Tempwrite
End sub

Sub EditAnnInfo()
	Dim Tempwrite,Boardlist,Readme,i
	dim trs,newsid,title,content,bgs
	If not caneditann then
		Dvbbs.AddErrCode(28)
		Exit sub
	End if
	If not isnumeric(request("id")) then
		Dvbbs.AddErrCode(42)
		Exit sub
	Else
		newsid=Clng(request("id"))
	End if
	Sql="select boardid,boardtype from Dv_board"
	Set rs=Dvbbs.execute(sql)
	If Rs.eof and Rs.bof Then
		Boardlist=""
	Else
		Sql=Rs.GetRows(-1)
		For i=0 to Ubound(Sql,2)
			Boardlist=Boardlist&"<option value='"&Sql(0,i)&"'"
			If Dvbbs.BoardID=Clng(Sql(0,i)) then Boardlist=Boardlist&" selected "
			Boardlist=Boardlist&">"&Sql(1,i)&"</option>"
		Next
	End if
	Rs.close:Set rs=nothing
	If Dvbbs.boardmaster Then
		Readme=""
	Else
		Readme=Template.Strings(16)
	End if
	Set rs=Dvbbs.execute("select title,content,bgs from Dv_bbsnews where id="&newsid)
	If rs.eof and rs.bof then
		title=""
		content=""
		bgs=""
	Else
		title=rs(0)
		content=rs(1)
		bgs=rs(2)
	End if
	Tempwrite=Template.html(9)
	Tempwrite=Replace(Tempwrite,"{$username}",Dvbbs.membername)
	Tempwrite=Replace(Tempwrite,"{$boardlist}",Boardlist)
	Tempwrite=Replace(Tempwrite,"{$readme}",Readme)
	Tempwrite=Replace(Tempwrite,"{$title}",Dv_FilterJS(title))
	Tempwrite=Replace(Tempwrite,"{$content}",Dv_FilterJS(content))
	Tempwrite=Replace(Tempwrite,"{$action}","?action=SaveEdit&id="&newsid)
	Tempwrite=Replace(Tempwrite,"{$dowhat}",Template.Strings(24))
	Tempwrite=Replace(Tempwrite,"{$bgs}",Dv_FilterJS(bgs))
	Response.Write Tempwrite
End sub

Sub SaveEdit()
	If not caneditann then
		Dvbbs.AddErrCode(28)
		Exit sub
	End if
	Dim username,title,content,bgs
	If not isnumeric(request("id")) then
		Dvbbs.AddErrCode(42)
		Exit sub
	End if
	If request("username")="" then
		Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(17)&"&action=OtherErr"
	Else
		username=Dvbbs.CheckStr(request("username"))
	End if
	If request("title")="" then
		Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(18)&"&action=OtherErr"
	Else
		title=Dvbbs.CheckStr(request("title"))
	End if
	If request("content")="" then
		Response.redirect "showerr.asp?ErrCodes=<li>"&template.Strings(19)&"&action=OtherErr"
	Else
		content=Dvbbs.CheckStr(request("content"))
	End if
	bgs=Dv_FilterJS(request("bgs"))
	Set rs=server.createobject("adodb.recordset")
	Sql="select * from Dv_bbsnews where id="&cstr(request("id"))
	If Not IsObject(Conn) Then ConnectionDatabase
	rs.open sql,conn,1,3
	rs("username")=fixjs(username)
	rs("title")=fixjs(title)
	rs("content")=content
	rs("addtime")=Now()
	rs("boardid")=Dvbbs.BoardID
	rs("bgs")=bgs
	rs.update
	rs.close
	Set Rs=Nothing
	If Dvbbs.BoardID>0 Then 
		Dvbbs.Name = "BoardInfo_" & Dvbbs.BoardID
		Dvbbs.LoadBoardNews_Paper(Dvbbs.BoardID)
	Else
		Dvbbs.DelCahe "news"&Dvbbs.boardid
	End If
	
	Dvbbs.Dvbbs_suc("<li>"&Template.Strings(25))
End sub

Sub delann()
	If not caneditann then
		Dvbbs.AddErrCode(28)
		Exit sub
	End if
	Dim delid,fixid
	delid=replace(request.form("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	fixid=replace(delid," ","")
	fixid=replace(fixid,",","")
	If Not IsNumeric(fixid) Then
		Dvbbs.AddErrCode(42)
		Exit Sub
	End If
	Dvbbs.execute("delete from Dv_bbsnews where id in ("&delid&")")
	Dvbbs.Dvbbs_suc("<li>"&Template.Strings(22))
	ReloadAnnouncement_All
End sub

Function ReloadAnnouncement_All
	Dim trs,i
	Set tRs=Dvbbs.execute("select boardid from [Dv_board]")
	If Not (tRs.eof and tRs.bof) then
		Sql=tRs.GetRows(-1)
		For i=0 To Ubound(Sql,2)
			Dvbbs.Name = "BoardInfo_" & Sql(0,i)
			Dvbbs.LoadBoardNews_Paper(Sql(0,i))
		Next
	End if
	tRs.close:Set trs=nothing
End Function

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
Function Dv_FilterJS(v)
	If Not Isnull(V) Then
		Dim t
		Dim re
		Dim reContent
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(&#)"
		t=re.Replace(v,"<I>&#</I>")
		re.Pattern="(script)"
		t=re.Replace(t,"<I>script</I>")
		re.Pattern="(js:)"
		t=re.Replace(t,"<I>js:</I>")
		re.Pattern="(value)"
		t=re.Replace(t,"<I>value</I>")
		re.Pattern="(about:)"
		t=re.Replace(t,"<I>about:</I>")
		re.Pattern="(file:)"
		t=re.Replace(t,"<I>file:</I>")
		re.Pattern="(Document.cookie)"
		t=re.Replace(t,"<I>Documents.cookie</I>")
		re.Pattern="(vbs:)"
		t=re.Replace(t,"<I>vbs:</I>")
		re.Pattern="(on(mouse|Exit|error|click|key))"
		t=re.Replace(t,"<I>on$2</I>")
		Dv_FilterJS=t
		Set Re=Nothing
	End If 
End Function
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