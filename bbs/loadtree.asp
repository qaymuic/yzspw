<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<!--#include file="inc/dv_ubbcode.asp"-->
<%
If Dvbbs.BoardID = 0 Then
	Response.Write "参数错误"
	Response.End 
End If
If Dvbbs.GroupSetting(2)="0"  Then
	Response.Write "<script language=""javascript"">alert('您没有权限查看贴子!')</script>"
	Response.End	
End If
Dvbbs.LoadTemplates("dispbbs")
Showtitle()
Dvbbs.Head()
Showtree()
Dvbbs.Footer()
Sub Showtitle()
	Dim treeData
	treeData=template.html(11)
	treeData=Replace(treeData,"{$treeloop}",template.html(14))
	Response.Write "<script language=""javascript"">"
	Response.Write Chr(10)
	Response.Write "var treedata='"
	Response.Write Replace(Replace(Replace(Replace(treeData,"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")
	Response.Write "';"
	Response.Write vbNewLine
	template.html(13)=split(template.html(13),"||")
	Response.Write template.html(13)(0)
	Response.Write Chr(10)
	Response.Write "</script>"
	Response.flush
End Sub 
Sub Showtree()
	'接收参数
	Dim AnnounceID,ReplyID,Star,TotalUseTable,openid
	openid=Request("openid")
	If openid="" Or Not IsNumeric(openid) Then openid=0
	openid=CLng(openid)
	AnnounceID=Request("ID")
	If AnnounceID="" Or Not IsNumeric(AnnounceID) Then Exit Sub 
	AnnounceID=CLng(AnnounceID)
	ReplyID=Request("ReplyID")
	If ReplyID="" Or Not IsNumeric(ReplyID) Then ReplyID=AnnounceID
	ReplyID=CLng(ReplyID)
	Star=Request("Star")
	If Star="" Or Not IsNumeric(Star) Then Star=1
	Star=Clng(Star)
	Dim SQl,Rs
	sql="Select PostTable from Dv_topic where topicID="&Announceid
	Set rs=Dvbbs.Execute(sql)
	If Not Rs.eof Then
		 TotalUseTable=Rs(0)
	End If
	Dim tmpstr,treedata,blank,i,j,RecordCount,PageCount
	PageCount=0
	If TotalUseTable <>"" Then 
		Set Rs=server.createobject("adodb.recordset")
		sql="select t.AnnounceID,t.parentID,t.BoardID,t.UserName,t.PostUserid,t.Topic,t.DateAndTime,t.length,t.RootID,t.layer,t.orders,t.Expression,t.body,t.locktopic,u.LockUser from "&TotalUseTable&" t Inner Join [dv_user] U On T.postuserid=u.userid where BoardID="&Dvbbs.BoardID&" and t.RootID="&Announceid&" and t.BoardID<>777 and t.BoardID<>444 order by t.RootID desc,t.orders"
		rs.open sql,conn,1,1
		j=0
		RecordCount=Rs.RecordCount
		If Not(Rs.EOF And Rs.BOF) Then
			Rs.PageSize=Cint(Dvbbs.Board_Setting(27))
			Rs.AbsolutePage=Star
			PageCount=Rs.PageCount
			Do while Not Rs.EOF
				treedata=template.html(12)
				For i=1 to Rs(9)
					blank=blank&"&nbsp;"
				Next
				If Rs("locktopic")=2 Then 
					treedata=Replace(treedata,"{$topic}","==此发言已被管理员屏蔽==")
				ElseIf Rs("LockUser")=1 Then
					treedata=Replace(treedata,"{$topic}","==此人已被管理员锁定==")
				ElseIf Rs("LockUser")=2 Then
					treedata=Replace(treedata,"{$topic}","==此人所有发言已被管理员屏蔽==")
				Else
					If Rs("topic")="" or isnull(rs("topic")) Then 
						treedata=Replace(treedata,"{$topic}",cutStr(Reubbcode(Rs("body")),35))
					Else
						treedata=Replace(treedata,"{$topic}",cutStr(Rs("Topic"),35))
					End If
				End If			
				If Rs(0)=openid Then
					treedata=Replace(treedata,"{$alertcolor}",Dvbbs.mainsetting(1))
				Else
					treedata=Replace(treedata,"{$alertcolor}","")
				End If 
				treedata=Replace(treedata,"{$announceid}",Rs(0))
				treedata=Replace(treedata,"{$boardid}",Rs(2))
				treedata=Replace(treedata,"{$username}",Rs(3))
				treedata=Replace(treedata,"{$DateAndTime}",Rs(6))
				If Rs(7)=0 Then 
					treedata=Replace(treedata,"{$length}",template.Strings(14))
				Else
					treedata=Replace(treedata,"{$length}",Rs(7)&template.Strings(13))
				End If
				treedata=Replace(treedata,"{$rootid}",Rs(8))
				treedata=Replace(treedata,"{$Expression}",Rs(11))
				treedata=Replace(treedata,"{$blank}",blank)
				blank=""
				tmpstr=tmpstr&treedata
				Rs.MoveNext
				j=j+1
				If j=Cint(Dvbbs.Board_Setting(27)) Then Exit Do
			Loop
		End If
		Rs.close
		Set rs=Nothing 
	End If
	template.html(11)=Replace(template.html(11),"{$treeloop}",tmpstr)
	Response.Write "<script language=""javascript"">"
	Response.Write Chr(10)
	Response.Write "var treedata='"
	Response.Write Replace(Replace(Replace(Replace(template.html(11),"\","\\"),"'","\'"),VbCrLf,"\n"),chr(13),"")
	Response.Write "';"
	Response.Write vbNewLine
	Response.Write template.html(13)(0)
	tmpstr=template.html(13)(1)
	tmpstr=Replace(tmpstr,"{$alertcolor}",Dvbbs.mainsetting(1))
	tmpstr=Replace(tmpstr,"{$boardid}",Dvbbs.BoardID)
	tmpstr=Replace(tmpstr,"{$replyid}",ReplyID)
	tmpstr=Replace(tmpstr,"{$announceid}",AnnounceID)
	tmpstr=Replace(tmpstr,"{$openid}",openid)
	Response.Write tmpstr
	'//树型分页代码,参数：页码,记录总数，每页显示数,页数
	'function showpage(page,RecordCount,PageSize,PageCount)
	Response.Write "showpage("&Star&","&RecordCount&","&Cint(Dvbbs.Board_Setting(27))&","&PageCount&");"
	Response.Write Chr(10)
	Response.Write "</script>"
End Sub

Function reUBBCode(strContent)
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	strContent=replace(strContent,"&nbsp;"," ")
	re.Pattern="(\[QUOTE\])(.|\n)*(\[\/QUOTE\])"
	strContent=re.Replace(strContent,"$2")
	re.Pattern="(\[point=*([0-9]*)\])(.|\n)*(\[\/point\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[post=*([0-9]*)\])(.|\n)*(\[\/post\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[power=*([0-9]*)\])(.|\n)*(\[\/power\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[usercp=*([0-9]*)\])(.|\n)*(\[\/usercp\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[money=*([0-9]*)\])(.|\n)*(\[\/money\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[replyview\])(.|\n)*(\[\/replyview\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="(\[usemoney=*([0-9]*)\])(.|\n)*(\[\/usemoney\])"
	strContent=re.Replace(strContent,"&nbsp;")
	re.Pattern="\[username=(.[^\[]*)\](.|\n)*\[\/username\]"
	strContent=re.Replace(strContent,"&nbsp;")
	strContent=replace(strContent,"<I></I>","")
	set re=Nothing
	reUBBCode=strContent
End Function
'截取指定字符
Function cutStr(str,strlen)
	'去掉所有HTML标记
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="<(.[^>]*)>"
	str=re.Replace(str,"")
	set re=Nothing
	str=Replace(str,chr(10),"")
	str = Dvbbs.HTMLEncode(str)
	Dim l,t,c,i
	l=Len(str)
	t=0
	For i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		If c>255 Then
			t=t+2
		Else
			t=t+1
		End If
		If t>=strlen Then
			cutStr=left(str,i)&"..."
			Exit For
		Else
			cutStr=str
		End If
	Next
End Function
%>