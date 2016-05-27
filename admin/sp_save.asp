<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
Const PurviewLevel=2
%>
<!--#include file="conn.asp"-->
<!--#include file="ChkPurview.asp"-->
<!--#include file="../inc/function.asp"-->
<%
dim splb,spname,spgqlb,spdq,spmj,spjg,spcontact,spaddress,spren,spcontent,spphoto,sptop1,sptop2,spendtime,spaddren,userip,BigClassName,SmallClassName
BigClassName=trim(request("BigClassName"))
SmallClassName=trim(request("SmallClassName"))
splb=trim(request("splb"))
spname=trim(request("spname"))
spgqlb=trim(request("spgqlb"))
spdq=trim(request("spdq"))
spmj=trim(request("spmj"))
spjg=trim(request("spjg"))
spcontact=trim(request("spcontact"))
spaddress=trim(request("spaddress"))
spren=trim(request("spren"))
spcontent=replace(replace(request.form("spcontent")," ","&nbsp;"),chr(13),"<br>")
spphoto=trim(request("Document1"))
sptop1=trim(request("sptop1"))
sptop2=trim(request("sptop2"))
spendtime=trim(request("spendtime"))
userip=request.ServerVariables("REMOTE_ADDR")
spaddren="后台"
	if spname="" or spmj="" or spjg="" or spcontact="" or spaddress="" then
		response.write  "<script>alert('请将商铺的相关信息填写完整！');history.go(-1)</script>"
		response.end
	end if

	dim sqlReg,rsReg
	sqlReg="select * from spw"
	set rsReg=server.createobject("adodb.recordset")
	rsReg.open sqlReg,conn,1,3
		rsReg.addnew
		rsReg("splb")=splb
		rsReg("spname")=spname
		rsReg("spgqlb")=spgqlb
		rsReg("BigClassName")=BigClassName
		rsReg("SmallClassName")=SmallClassName
		rsReg("spmj")=spmj
		rsReg("spjg")=spjg
		rsReg("spcontact")=spcontact
		rsReg("spaddress")=spaddress
		rsReg("spren")=spren
                rsReg("userip")=userip
		rsReg("spcontent")=spcontent
		if spphoto="" then
		rsReg("spphoto")="UploadFiles/noimg.jpg"
		else
		rsReg("spphoto")=spphoto
		end if
		if sptop1="yes" then
		rsReg("sptop1")=True
		end if
		if sptop2="yes" then
		rsReg("sptop2")=True
		end if
		rsReg("spendtime")=spendtime
		rsReg("spaddren")=spaddren
		rsReg.update
	rsReg.close
	set rsReg=nothing
	response.write  "<script>alert('商铺发布成功！');location.href='sp_manage.asp'</script>"
	response.end
%>
