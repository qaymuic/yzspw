<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	Dim n,m,h,s
	Dim y,z,show
	Dim stats,Forum_userfaceNum,Forum_userface,Forum_facepath
	Dim i
	stats="论坛头像列表"
	h=6 '每行显示个数
	s=30 '每页显示个数

	Dvbbs.LoadTemplates("usermanager")
	Dvbbs.Stats="用户头像列表"
	Dvbbs.head()
	'Dvbbs.Head_var 0,0,template.Strings(0),"usermanager.asp"
	Forum_userface=split(Dvbbs.Forum_userface,"|||")
	Forum_facepath=Forum_userface(0)
	Forum_userfacenum=ubound(Forum_userface)
	M=Forum_userfaceNum-1 '头象个数
	Call main()
	Dvbbs.Footer()

Sub main()
	Dim TempStr,TempLateStr
	Dim Facebar,p
	TempStr=template.Strings(66)
	If m>s Then
		For Y=1 to (m+s-1)\s
		Facebar=Facebar&Replace(TempStr,"{$page}",y)
		Next
	End If
	TempStr=Split(template.html(19),"||")
	TempStr(0)=Replace(TempStr(0),"{$FaceBar}",Facebar)
	Response.Write TempStr(0)
	i=0
	p= trim(request("show"))
	if p="" or not IsNumeric(p) then
	p=1
	end if
	z=s*p
	if z/m>1 then
		z=m
	else
		z=s*p
	end if
	for n=s*p+1-s to z
	i=i+1
	TempLateStr=TempStr(1)
	TempLateStr=Replace(TempLateStr,"{$Forum_facepath}",Forum_facepath)
	TempLateStr=Replace(TempLateStr,"{$Forum_userface_n}",Forum_userface(n))
	Response.Write TempLateStr
	if i=h then
	response.write TempStr(2)
	i=0
	end if 
	next
	Response.Write TempStr(3)
End Sub
%>