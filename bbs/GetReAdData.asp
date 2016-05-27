<!--#include file="Conn.asp"-->
<!--#include file="connad.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/base64.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Response.Clear
Server.ScriptTimeout=999999
dim rs,sql,i
'on error resume next
dim rechallengeWord,retokerWord,redata,paycode
dim challengeWord_key,rechallengeWord_key
dim trs,boarduser
dim datanum,maxadid
dim forum_ad1,forum_ad2,forum_ad3
dim adinfo_lengthb


If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(11)=1) Then
	Response.Write "本论坛没有开启同步广告功能。"
	Response.End
End If

'redata=adinfo_text
rechallengeWord=trim(Dvbbs.CheckStr(request("challengeWord")))
retokerWord=trim(request("tokenWord"))

challengeWord_key=session("challengeWord_key")
session("challengeWord_key")=Empty
'Response.Write Dvbbs.CacheData(21,0)
if md5(rechallengeWord & ":" & Dvbbs.CacheData(21,0),32)=retokerWord then
datanum=Clng(request("datanum"))
for i=1 to datanum
	redata=redata & trim(request.form("data" & i))
next
Response.Write "100"
'Response.Write datanum
'Response.Write ","
'Response.Write left(redata,10)
'Response.end
	'返回成功信息
	'假设有20条广告和每条广告有30条资源信息
	'每条广告循环
	dim AdLength,AdLength_for
	dim ii,iii
	dim iaddress,filetype,rate,adcode_length,adcode
	dim adinfo_length
	dim First_length,Getadinfo_length
	dim Adinfo_name_length,Have_length,Adinfo_name,Adinfo_type_length,Adinfo_type,Adinfo_content_length,Adinfo_content
	dim Ad_for_length
	dim TotalID
	dim foundad1,foundad2,foundad3,foundad4
	foundad1=false
	foundad2=false
	foundad3=false
	foundad4=false

	set rs=dvbbs.execute("select top 1 * from Dv_ChallengeInfo")
	Dim MouseID
	MouseID=rs("D_username")
	set rs=dvbbs.execute("select max(A_ID) from Dv_AdCode")
	MaxAdID=rs(0)+1

	for iii=1 to 100

	'Response.Write iii
	if iii=1 then
		AdLength=cCur(left(redata,10))
		'广告条入库（父广告条）
		iaddress=mid(redata,11,4)
		filetype=mid(redata,15,4)
		rate=mid(redata,19,4)
		'广告代码长度（base64编码）
		adcode_length=cCur(mid(redata,23,8))
		'广告代码（base64解码）
		adcode=strAnsi2Unicode(Base64decode(strUnicode2Ansi(mid(redata,31,adcode_length))))
		
		if IsSqlDataBase=1 then
			dvbbs.execute("delete from dv_AdCode where a_address='"&iaddress&"' and A_ID<"&MaxAdID&"")
		else
			dvbbs.execute("delete from dv_AdCode where a_address='"&iaddress&"' and A_ID<"&MaxAdID&"")
		end if
		Select Case "iaddress"
		Case "0001"
			foundad1=true
		Case "0002"
			foundad2=true
		Case "0003"
			foundad3=true
		Case "0004"
			foundad4=true
		End Select
		If Not Trim(iaddress)="9999" Then
		set rs=Server.CreateObject("ADODB.Recordset")
		sql="select * from Dv_AdCode"
		rs.open sql,conn,1,3
		rs.addnew
		rs("A_Address")=iaddress
		rs("A_filetype")=filetype
		rs("A_rate")=rate
		rs("A_Adcode")=RePicUrl(adcode)
		rs.update
		rs.close
		set rs=nothing
		End If

		First_length=30 + adcode_length
		'父广告条中资源循环
		'父广告条中所有资源的总长度
		Getadinfo_length=AdLength - First_length
		if (AdLength + 10)>First_length then
		for ii=1 to 2000
			if ii=1 then
				'资源名称长度，4位
				Adinfo_name_length=cCur(mid(redata,First_length + 1,4))
				'Response.Write ","&mid(redata,First_length + 1,4)&""
				'资源名称
				Have_length=First_length + 4
				Adinfo_name=strAnsi2Unicode(Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_name_length))))
				'Response.Write "," & Adinfo_name
				'资源类型长度，2位
				Have_length=Have_length + Adinfo_name_length
				'Response.Write mid(redata,Have_length + 1,2)
				Adinfo_type_length=cCur(mid(redata,Have_length + 1,2))
				'Response.Write "," & Adinfo_type_length
				'资源类型
				Have_length=Have_length + 2
				Adinfo_type=strAnsi2Unicode(Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_type_length))))
				'Response.Write "," & Adinfo_type
				'资源长度，8位
				Have_length=Have_length + Adinfo_type_length
				Adinfo_content_length=cCur(mid(redata,Have_length + 1,8))
				'Response.Write "," & Adinfo_content_length
				'资源
				Have_length=Have_length + 8
				Adinfo_content=Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_content_length)))
				'Response.Write "," & Adinfo_content
				'本资源总长度
				Adinfo_length=Have_length + Adinfo_content_length
				adinfo_lengthb=lenb(Adinfo_content)
				if adinfo_lengthb mod 2 <> 0 then
					Adinfo_content=Adinfo_content & chrB(13) & chrB(10)
				end if
				'入库
				If Trim(iaddress)="9999" Then
					Connad.Execute("delete from dv_chanad where A_Adname='"&Adinfo_name&"'")
				Else
				set rs=Server.CreateObject("ADODB.Recordset")
				sql="select * from Dv_ChanAd where A_Adname='"&Adinfo_name&"'"
				rs.open sql,connad,1,3
				if rs.eof and rs.bof then
				rs.addnew
				rs("A_Adname")=Adinfo_name
				rs("A_Adtype")=Adinfo_type
				rs("A_data").Appendchunk Adinfo_content
				rs.update
				else
				rs("A_Adname")=Adinfo_name
				rs("A_Adtype")=Adinfo_type
				rs("A_data").Appendchunk Adinfo_content
				rs.update
				end if
				rs.close
				set rs=nothing
				End If
			else
				if Adinfo_length=Getadinfo_length then exit for
				if Adinfo_length<Getadinfo_length then
					'资源名称长度，4位
					Adinfo_name_length=cCur(mid(redata,Adinfo_length + 1,4))
					'资源名称
					Have_length=Adinfo_length + 4
					Adinfo_name=strAnsi2Unicode(Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_name_length))))
					'资源类型长度，2位
					Have_length=Have_length + Adinfo_name_length
					Adinfo_type_length=cCur(mid(redata,Have_length + 1,2))
					'资源类型
					Have_length=Have_length + 2
					Adinfo_type=strAnsi2Unicode(Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_type_length))))
					'资源长度，8位
					Have_length=Have_length + Adinfo_type_length
					Adinfo_content_length=cCur(mid(redata,Have_length + 1,8))
					'资源
					Have_length=Have_length + 8
					Adinfo_content=Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_content_length)))
					'本资源总长度
					Adinfo_length=Have_length + Adinfo_content_length
					adinfo_lengthb=lenb(Adinfo_content)
					if adinfo_lengthb mod 2 <> 0 then
						Adinfo_content=Adinfo_content & chrB(13) & chrB(10)
					end if
					If Trim(iaddress)="9999" Then
						Connad.Execute("delete from dv_chanad where A_Adname='"&Adinfo_name&"'")
					Else
					set rs=Server.CreateObject("ADODB.Recordset")
					sql="select * from Dv_ChanAd where A_Adname='"&Adinfo_name&"'"
					rs.open sql,connad,1,3
					if rs.eof and rs.bof then
					rs.addnew
					rs("A_Adname")=Adinfo_name
					rs("A_Adtype")=Adinfo_type
					rs("A_data").Appendchunk Adinfo_content
					rs.update
					else
					rs("A_Adname")=Adinfo_name
					rs("A_Adtype")=Adinfo_type
					rs("A_data").Appendchunk Adinfo_content
					rs.update
					end if
					rs.close
					set rs=nothing
					End If
				else
					exit for
				end if
			end if
		next
		end if
		AdLength=AdLength + 10
		'Response.end
	else
		if AdLength=len(redata) then exit for
		if AdLength<len(redata) then
			'当前广告条长度，10位
			
			'Ad_for_length=cCur(left(redata,AdLength + 10))
			Ad_for_length=cCur(mid(redata,AdLength + 1,10))
			'Response.Write Ad_for_length
			'Response.end

			'广告条入库（父广告条）
			iaddress=mid(redata,AdLength + 11,4)
			filetype=mid(redata,AdLength + 15,4)
			rate=mid(redata,AdLength + 19,4)
			'广告代码长度（base64编码）
			adcode_length=cCur(mid(redata,AdLength + 23,8))
			'广告代码（base64解码）
			adcode=strAnsi2Unicode(Base64decode(strUnicode2Ansi(mid(redata,AdLength + 31,adcode_length))))
			'Response.Write base64decode(cstr(mid(redata,AdLength + 31,adcode_length)))
			'response.end
	
			if IsSqlDataBase=1 then
				dvbbs.execute("delete from dv_AdCode where a_address='"&iaddress&"' and A_ID<"&MaxAdID&"")
			else
				dvbbs.execute("delete from dv_AdCode where a_address='"&iaddress&"' and A_ID<"&MaxAdID&"")
			end if
			If Not Trim(iaddress)="9999" Then
			set rs=Server.CreateObject("ADODB.Recordset")
			sql="select * from Dv_AdCode"
			rs.open sql,conn,1,3
			rs.addnew
			rs("A_Address")=iaddress
			rs("A_filetype")=filetype
			rs("A_rate")=rate
			rs("A_Adcode")=RePicUrl(adcode)
			rs.update
			rs.close
			set rs=nothing
			End If
			
			First_length=30 + adcode_length
			'父广告条中资源循环
			'父广告条中所有资源的总长度
			Getadinfo_length=AdLength + Ad_for_length - First_length
			
			'Response.Write Ad_for_length + 10
			'Response.Write ","
			'Response.Write First_length
			'Response.Write "."
			if (Ad_for_length + 10)>First_length then
			for ii=1 to 2000
				if ii=1 then
					'资源名称长度，4位
					'Response.Write AdLength & "," & i
					'response.end
					First_length=AdLength + First_length
					Adinfo_name_length=cCur(mid(redata,First_length + 1,4))
					'资源名称
					Have_length=First_length + 4
					Adinfo_name=strAnsi2Unicode(Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_name_length))))
					'资源类型长度，2位
					Have_length=Have_length + Adinfo_name_length
					Adinfo_type_length=cCur(mid(redata,Have_length + 1,2))
					'资源类型
					Have_length=Have_length + 2
					Adinfo_type=strAnsi2Unicode(Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_type_length))))
					'资源长度，8位
					Have_length=Have_length + Adinfo_type_length
					Adinfo_content_length=cCur(mid(redata,Have_length + 1,8))
					'资源
					Have_length=Have_length + 8
					'Response.Write mid(redata,Have_length + 1,Adinfo_content_length)
					'response.end
					Adinfo_content=Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_content_length)))
					'本资源总长度
					Adinfo_length=Have_length + Adinfo_content_length
					adinfo_lengthb=lenb(Adinfo_content)
					if adinfo_lengthb mod 2 <> 0 then
						Adinfo_content=Adinfo_content & chrB(13) & chrB(10)
					end if

					'入库
					If Trim(iaddress)="9999" Then
						Connad.Execute("delete from dv_chanad where A_Adname='"&Adinfo_name&"'")
					Else
					set rs=Server.CreateObject("ADODB.Recordset")
					sql="select * from Dv_ChanAd where A_Adname='"&Adinfo_name&"'"
					rs.open sql,connad,1,3
					if rs.eof and rs.bof then
					rs.addnew
					rs("A_Adname")=Adinfo_name
					rs("A_Adtype")=Adinfo_type
					rs("A_data").Appendchunk Adinfo_content
					rs.update
					else
					rs("A_Adname")=Adinfo_name
					rs("A_Adtype")=Adinfo_type
					rs("A_data").Appendchunk Adinfo_content
					rs.update
					end if
					rs.close
					set rs=nothing
					End If
				else
					if Adinfo_length=Getadinfo_length then exit for
					if Adinfo_length<Getadinfo_length then
						'资源名称长度，4位
						Adinfo_name_length=cCur(mid(redata,Adinfo_length + 1,4))
						'资源名称
						Have_length=Adinfo_length + 4
						Adinfo_name=strAnsi2Unicode(Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_name_length))))
						'Response.Write adinfo_name & ","
						'资源类型长度，2位
						Have_length=Have_length + Adinfo_name_length
						Adinfo_type_length=cCur(mid(redata,Have_length + 1,2))
						'资源类型
						Have_length=Have_length + 2
						Adinfo_type=strAnsi2Unicode(Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_type_length))))
						'资源长度，8位
						Have_length=Have_length + Adinfo_type_length
						Adinfo_content_length=cCur(mid(redata,Have_length + 1,8))
						'资源
						Have_length=Have_length + 8
						Adinfo_content=Base64decode(strUnicode2Ansi(mid(redata,Have_length + 1,Adinfo_content_length)))
						'本资源总长度
						Adinfo_length=Have_length + Adinfo_content_length
						adinfo_lengthb=lenb(Adinfo_content)
						if adinfo_lengthb mod 2 <> 0 then
							Adinfo_content=Adinfo_content & chrB(13) & chrB(10)
						end if
						If Trim(iaddress)="9999" Then
							Connad.Execute("delete from dv_chanad where A_Adname='"&Adinfo_name&"'")
						Else
						set rs=Server.CreateObject("ADODB.Recordset")
						sql="select * from Dv_ChanAd where A_Adname='"&Adinfo_name&"'"
						rs.open sql,connad,1,3
						if rs.eof and rs.bof then
						rs.addnew
						rs("A_Adname")=Adinfo_name
						rs("A_Adtype")=Adinfo_type
						rs("A_data").Appendchunk Adinfo_content
						rs.update
						else
						rs("A_Adname")=Adinfo_name
						rs("A_Adtype")=Adinfo_type
						rs("A_data").Appendchunk Adinfo_content
						rs.update
						end if
						rs.close
						set rs=nothing
						End If
					else
						exit for
					end if
				end if
			next
			end if
			AdLength=AdLength + Ad_for_length + 10
		else
			exit for
		end if
		'if isnull(left(redata,AdLength+10)) or left(redata,AdLength+10)="" then
		'	exit for
		'else
		'	AdLength_for=cCur(left(redata,AdLength+10))

		'	AdLength=AdLength + AdLength_for
		'end if
	end if
	next

	dvbbs.execute("update dv_Setup set forum_ad=''")
	set rs=dvbbs.execute("select * from dv_adcode where a_address='0001'")
	do while not rs.eof
		if forum_ad1="" then
			forum_ad1=rs("a_id")
		else
			forum_ad1=forum_ad1 & "," & rs("a_id")
		end if
	rs.movenext
	loop
	set rs=dvbbs.execute("select * from dv_adcode where a_address='0002'")
	do while not rs.eof
		if forum_ad2="" then
			forum_ad2=rs("a_id")
		else
			forum_ad2=forum_ad2 & "," & rs("a_id")
		end if
	rs.movenext
	loop
	set rs=dvbbs.execute("select * from dv_adcode where a_address='0003'")
	do while not rs.eof
		if forum_ad3="" then
			forum_ad3=rs("a_id")
		else
			forum_ad3=forum_ad3 & "," & rs("a_id")
		end if
	rs.movenext
	loop
	Forum_Ad1 = Forum_Ad1 & "||" & Forum_Ad2 & "||" & Forum_Ad3
	dvbbs.execute("update dv_setup set forum_ad='"&forum_ad1&"'")
	set rs=nothing
	Dvbbs.Name="setup"
	Dvbbs.ReloadSetup
	Dvbbs.DelCahe "ForumAdCode1"
	Dvbbs.DelCahe "ForumAdCode2"
	Dvbbs.DelCahe "ForumAdCode3"
	Dvbbs.DelCahe "TopicAdCode"

else
	Response.Write "101"

	Response.Write "ray chanword:" & rechallengeWord
	Response.Write ","
	Response.Write "local chanword:" & challengeWord_key
	Response.Write ","
	Response.Write "ray tokerword:"&retokerWord
	Response.Write ","
	Response.Write "local tokerword:" &md5(rechallengeWord & ":" & Dvbbs.CacheData(21,0),32)
end if

Function RePicUrl(poststr)
	if poststr="" then
		RePicUrl=poststr
		exit function
	end if
	poststr=replace(poststr,".gif%show%",".gif")
	poststr=replace(poststr,".jpg%show%",".jpg")
	poststr=replace(poststr,".bmp%show%",".bmp")
	poststr=replace(poststr,".jpeg%show%",".jpeg")
	poststr=replace(poststr,".png%show%",".png")
	poststr=replace(poststr,".tif%show%",".tif")
	poststr=replace(poststr,".swf%show%",".swf")
	poststr=replace(poststr,"%show%","show_ad_sc.asp?fn=")
	RePicUrl=poststr
End Function

connad.close
set connad=nothing
%>