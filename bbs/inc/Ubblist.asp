<%
Function Ubblist(Str1)
	Dim Str
	Str=Str1
	Str=LCase(Str)
	Dim Dv_ubb(51),i
	Dv_ubb(0)="javascript":Dv_ubb(1)="jscript:":Dv_ubb(2)="js:":Dv_ubb(3)="value":Dv_ubb(4)="about:"
	Dv_ubb(5)="file:":Dv_ubb(6)="document.cookie":Dv_ubb(7)="vbscript:":Dv_ubb(8)="vbs:":Dv_ubb(9)="mouse"
	Dv_ubb(10)="exit":Dv_ubb(11)="error":Dv_ubb(12)="click":Dv_ubb(13)="key":Dv_ubb(14)="[/img]"
	Dv_ubb(15)="[/upload]":Dv_ubb(16)="[/dir]":Dv_ubb(17)="[/qt]":Dv_ubb(18)="[/mp]":Dv_ubb(19)="[/rm]"
	Dv_ubb(20)="[/sound]":Dv_ubb(21)="[/flash]":Dv_ubb(22)="[/money]":Dv_ubb(23)="[/point]":Dv_ubb(24)="[/usercp]"
	Dv_ubb(25)="[/power]":Dv_ubb(26)="[/post]":Dv_ubb(27)="[/replyview]":Dv_ubb(28)="[/usemoney]":Dv_ubb(29)="[/url]"
	Dv_ubb(30)="[/email]":Dv_ubb(31)="http":Dv_ubb(32)="https":Dv_ubb(33)="ftp":Dv_ubb(34)="rtsp":Dv_ubb(35)="mms"
	Dv_ubb(36)="[/html]":Dv_ubb(37)="[/code]":Dv_ubb(38)="[/color]":Dv_ubb(39)="[/face]":Dv_ubb(40)="[/align]"
	Dv_ubb(41)="[/quote]":Dv_ubb(42)="[/fly]":Dv_ubb(43)="[/move]":Dv_ubb(44)="[/shadow]":Dv_ubb(45)="[/glow]"
	Dv_ubb(46)="[/size]":Dv_ubb(47)="[/i]":Dv_ubb(48)="[/b]":Dv_ubb(49)="[/u]":Dv_ubb(50)="[em":Dv_ubb(51)="www."
	Ubblist=","
	'检查脚本标记0
	For i=0 to 13
		If InStr(Str,Dv_ubb(i))>0 Or InStr(Str,"&#")>0  or InStr(Str,"button") Then
				Ubblist=Ubblist&"0,"
			Exit For
		End If
	Next
	'Response.Write "Ubblist:<br>"
	For i=14 to 51
		If InStr(Str,Dv_ubb(i))>0 Then
			Ubblist=Ubblist & i-13 &","
		End If
		'Response.Write Dv_ubb(i)&"编号:"&i-13&"."		
	Next
	Ubblist=Ubblist & "39,"
	Dim TmpStr
	'去掉字符串中的空格
	If IsNull(Str) Then Exit Function
	TmpStr=Replace(Str," ","")
	TmpStr=Replace(TmpStr,"  ","")
	TmpStr=Replace(TmpStr,Chr(10),"")
	TmpStr=Replace(TmpStr,Chr(13),"")
	'Ubblist=Ubblist&fixtable(TmpStr,"table")'&fixtable(TmpStr,"td")&fixtable(TmpStr,"th")&fixtable(TmpStr,"tr")
	Ubblist=Ubblist&checkTable(TmpStr)
End Function
'检查表格
Function fixtable(Str,flag)
	Dim bodyStr
	bodyStr=Str
	If bodyStr<>"" Then
		Dim instr1,instr2
		Dim Str1,Str2,newstr1,newstr2
		Str2="</"&flag&">"
		Str1="<"&flag
		Instr1=InStr(bodyStr,str1)
		Instr2=InStr(bodyStr,Str2)
		If Instr1 >0 And Instr2>0 Then
			If Instr1 > Instr2 Then
				fixtable=flag&","
				Exit Function
			Else
				newstr1=Mid(bodyStr,InStr2+Len(Str2),Len(bodyStr)-(instr2+Len(Str2)-instr1))
				Newstr2=Mid(bodyStr,InStr1+Len(Str1),instr2-instr1-Len(str2)+2)												
				fixtable=fixtable(newstr1,flag)&fixtable(newstr2,flag)
			End If
		ElseIf Instr1 =0 And Instr2=0 Then
			Exit Function
		ElseIf Instr1 >0 And Instr2=0 Then
			fixtable=flag&","
			Exit Function
		ElseIf Instr1 =0 And Instr2>0 Then
			fixtable=flag&","
			Exit Function
		End If
	Else
		fixtable=""
	End If
End Function
Function UbblistOLD(Str)
	UbblistOLD=Ubblist(Str)
	UbblistOLD=Replace(UbblistOLD,"39,","")
End Function
Function checkTable(tablestr)
	checkTable=""
	Dim s,tmpstr1,tmpstr2
	s=tablestr
	tmpstr1=split(s,"<table")
	tmpstr2=split(s,"</table>")
	If UBound(tmpstr1)<>UBound(tmpstr2) Then
		checkTable=checkTable&"table,"
	End If
	tmpstr1=split(s,"<td")
	tmpstr2=split(s,"</td>")
	If UBound(tmpstr1)<>UBound(tmpstr2) Then
		checkTable=checkTable&"td,"
	End If
	tmpstr1=split(s,"<div")
	tmpstr2=split(s,"</div>")
	If UBound(tmpstr1)<>UBound(tmpstr2) Then
		checkTable=checkTable&"tr,"
	End If
	tmpstr1=split(s,"<th")
	tmpstr2=split(s,"</th>")
	If UBound(tmpstr1)<>UBound(tmpstr2) Then
		checkTable=checkTable&"th,"
	End If
End Function
%>