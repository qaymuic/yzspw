<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim SmallClassName,BigClassName,splb,spname,spgqlb,spdq,spmj,spjg,spcontact,spaddress,spren,spcontent,spphoto,sptop1,sptop2,spendtime,spaddren,userip
splb=trim(request("splb"))
BigClassName=trim(request("BigClassName"))
spname=trim(request("spname"))
spgqlb=trim(request("spgqlb"))
SmallClassName=trim(request("SmallClassName"))
spmj=trim(request("spmj"))
spjg=trim(request("spjg"))
spcontact=trim(request("spcontact"))
spaddress=trim(request("spaddress"))
spren=trim(request("spren"))
spcontent=replace(replace(request.form("spcontent")," ","&nbsp;"),chr(13),"<br>")
spphoto=trim(request("Document1"))
spendtime=trim(request("spendtime"))
userip=request.ServerVariables("REMOTE_ADDR")
spaddren="前台"
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
		rsReg("spendtime")=spendtime
		rsReg("spaddren")=spaddren
		rsReg.update
	rsReg.close
	set rsReg=nothing
	response.write  "<script>alert('商铺发布成功！');location.href='splist.asp'</script>"
	response.end
%>
