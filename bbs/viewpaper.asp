<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/Dv_ubbcode.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<%
Dvbbs.LoadTemplates("paper_even_toplist")
Dvbbs.stats=template.Strings(3)
Dvbbs.Head()
Dim paperid
Dim username
If request("id")="" Then
	Dvbbs.AddErrCode(35)
ElseIf Not IsNumeric(request("id")) Then
	Dvbbs.AddErrCode(35)
Else
	paperID=clng(request("id"))
End If
Dvbbs.ShowErr()
Dim dv_ubb,abgcolor
Set dv_ubb=new Dvbbs_UbbCode
Dim Rs,SQl
set rs=server.createobject("adodb.recordset")
sql="select * from dv_smallpaper where s_id="&paperid
set rs=dvbbs.execute(sql)
if rs.eof and rs.bof Then
	Dvbbs.AddErrCode(32)
	rs.close
	Set rs=nothing	
	Dvbbs.ShowErr()
Else
	dvbbs.execute("update dv_smallpaper set s_hits=s_hits+1 where s_id="&paperid)
	Dim TempStr
	TempStr = template.html(4)
	TempStr = Replace(TempStr,"{$title}",Dvbbs.Htmlencode(rs("s_title")))
	TempStr = Replace(TempStr,"{$username}",Dvbbs.Htmlencode(rs("s_username")))
	TempStr = Replace(TempStr,"{$hits}",rs("s_hits"))
	TempStr = Replace(TempStr,"{$content}",dv_ubb.Dv_UbbCode(Rs("s_content"),4,2,1))
	TempStr = Replace(TempStr,"{$addtime}",Dvbbs.Htmlencode(rs("s_addtime")))
	Response.Write TempStr
	rs.close
	Set rs=nothing
End If

Dvbbs.activeonline()
Dvbbs.footer()
%>