<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/dv_clsother.asp"-->
<%
Dim ErrString,action,i,showstr
action=Request("action")
Dvbbs.LoadTemplates("showerr")
If Dvbbs.forum_setting(79)="0" Then
	Template.html(1) = Replace(Template.html(1),"{$getcode}","")
Else
	template.html(3)=Replace(template.html(3),"{$codestr}",Dvbbs.GetCode())
	Template.html(1) = Replace(Template.html(1),"{$getcode}",template.html(3))
End If
Select Case action
	Case "stop"'��̳��ͣ
		Dvbbs.Stats=Template.Strings(1)
		Dvbbs.head()
		template.html(2)=Replace(template.html(2),"{$title}",Template.Strings(2)&Template.Strings(1))
		If Dvbbs.BoardID=0 Then
			If Dvbbs.Forum_Setting(69)="0" Or Dvbbs.Forum_Setting(21)="1" Then 
				template.html(2)=Replace(template.html(2),"{$stopreadme}",Stopreadme)
			Else
				Dvbbs.Forum_Setting(70)=Split(Dvbbs.Forum_Setting(70),"|")
				showstr="<br><b>&nbsp;&nbsp;"&Dvbbs.Forum_Info(0)&"</b>�����˶�ʱ���ţ��밴�����ʱ����ʣ�<hr size=""1""><ul>"
				For i=0 to UBound(Dvbbs.Forum_Setting(70))
					If i mod 6=0 Then showstr=showstr&"<li>"
					If i<10 Then showstr=showstr&"&nbsp;"
					showstr=showstr&i&"�㣺"
					If Dvbbs.Forum_Setting(70)(i)="1" Then
						showstr=showstr&"����&nbsp;&nbsp;"
					Else
						showstr=showstr&"<font color=""#FF0000"">�ر�</font>&nbsp;&nbsp;"
					End If
				Next
				showstr=showstr&"</ul>"
				template.html(2)=Replace(template.html(2),"{$stopreadme}",showstr)
			End If
			
		Else
			Dvbbs.Board_Setting(22)=Split(Dvbbs.Board_Setting(22),"|")
			showstr="<br><b>&nbsp;&nbsp;"&Dvbbs.boardtype&"</b>�����˶�ʱ���ţ��밴�����ʱ����ʣ�<hr size=""1""><ul>"
			For i=0 to UBound(Dvbbs.Board_Setting(22))
				If i mod 6=0 Then showstr=showstr&"<li>"
				If i<10 Then showstr=showstr&"&nbsp;"
				showstr=showstr&i&"�㣺"
				If Dvbbs.Board_Setting(22)(i)="1" Then
					showstr=showstr&"����&nbsp;&nbsp;"
				Else
					showstr=showstr&"<font color=""#FF0000"">�ر�</font>&nbsp;&nbsp;"
				End If
			Next
			showstr=showstr&"</ul>"
			template.html(2)=Replace(template.html(2),"{$stopreadme}",showstr)
		End If 
		Response.Write Template.html(2)
	Case "iplock"'IP����
		Dvbbs.Stats=Template.Strings(4)
		Dvbbs.head()
		Session(Dvbbs.CacheName & "UserID")=empty 
		template.html(2)=Replace(template.html(2),"{$title}",Template.Strings(4))
		Template.Strings(5)=replace(Template.Strings(5),"{$ip}",Dvbbs.usertrueip)
		Template.Strings(5)=replace(Template.Strings(5),"{$email}",Dvbbs.Forum_Info(5))
		template.html(2)=Replace(template.html(2),"{$stopreadme}",Template.Strings(5))
		Response.Write Template.html(2)
	Case "limitedonline"'���߱���
		Dvbbs.Stats=Template.Strings(4)
		Dvbbs.head()
		template.html(2)=Replace(template.html(2),"{$title}",Template.Strings(23))
		template.html(2)=Replace(template.html(2),"{$stopreadme}",Replace(Template.Strings(22),"{$onlinelimited}",Request("lnum")))
		Response.Write Template.html(2)
	Case "OtherErr"
		Dvbbs.Stats=action&"-"&Template.Strings(0)
		Dvbbs.head()
		Dvbbs.showtoptable()
		Dvbbs.Head_var 0,"",Template.Strings(0),""
		template.html(0)=Replace(template.html(0),"{$color}",Dvbbs.mainsetting(1))
		template.html(0)=Replace(template.html(0),"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		template.html(0)=Replace(template.html(0),"{$action}","������̳")
		template.html(0)=Replace(template.html(0),"{$ErrCount}",1)
		template.html(0)=Replace(template.html(0),"{$ErrString}",Request("ErrCodes"))
		If Request("autoreload")=1 Then	
			Response.Write "<meta http-equiv=refresh content=""2;URL="&Request.ServerVariables("HTTP_REFERER")&""">"		
		End If
		Response.Write Template.html(0)
		If dvbbs.userid=0 Then 
			Response.Write Template.html(1)
		End If
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
	Case "readonly"
		Dvbbs.Stats="��ǰ��̳Ϊֻ��"
		Dvbbs.head()
		Dvbbs.showtoptable()
		Dvbbs.Head_var 1,Dvbbs.boardtype,"",""
		template.html(2)=Replace(template.html(2),"{$title}",Template.Strings(2)&"��ǰʱ����̳Ϊֻ��")
		If Dvbbs.Board_Setting(21)="2" Then
			Dvbbs.Board_Setting(22)=Split(Dvbbs.Board_Setting(22),"|")
			showstr="<br><b>&nbsp;&nbsp;"&Dvbbs.boardtype&"</b>�����˶�ʱ���ŷ��������ڹ涨��ʱ���ڷ�����<hr size=""1""><ul>"
			
			For i=0 to UBound(Dvbbs.Board_Setting(22))
				If i mod 6=0 Then showstr=showstr&"<li>"
				If i<10 Then showstr=showstr&"&nbsp;"
				showstr=showstr&i&"�㣺"
				If Dvbbs.Board_Setting(22)(i)="1" Then
					showstr=showstr&"����&nbsp;&nbsp;"
				Else
					showstr=showstr&"<font color=""#FF0000"">�ر�</font>&nbsp;&nbsp;"
				End If
			Next
			showstr=showstr&"</ul>"
		End If
		If Dvbbs.Forum_Setting(69) ="2" Then 
			Dvbbs.Forum_Setting(70)=Split(Dvbbs.Forum_Setting(70),"|")
				showstr="<br><b>&nbsp;&nbsp;"&Dvbbs.Forum_Info(0)&"</b>�����˵�ǰʱ��Ϊֻ��״̬�����ڹ涨��ʱ���ڷ�����<hr size=""1""><ul>"
				For i=0 to UBound(Dvbbs.Forum_Setting(70))
					If i mod 6=0 Then showstr=showstr&"<li>"
					If i<10 Then showstr=showstr&"&nbsp;"
					showstr=showstr&i&"�㣺"
					If Dvbbs.Forum_Setting(70)(i)="1" Then
						showstr=showstr&"����&nbsp;&nbsp;"
					Else
						showstr=showstr&"<font color=""#FF0000"">�ر�</font>&nbsp;&nbsp;"
					End If
				Next
				showstr=showstr&"</ul>"
				template.html(2)=Replace(template.html(2),"{$stopreadme}",showstr)
			End If
			template.html(2)=Replace(template.html(2),"{$stopreadme}",showstr)
		Response.Write Template.html(2)
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
	Case "lock"
		Dvbbs.Stats="��̳������"
		Dvbbs.head()
		Dvbbs.showtoptable()
		Dvbbs.Head_var 0,"",Dvbbs.boardtype,""
		template.html(2)=Replace(template.html(2),"{$title}",Template.Strings(2)&"��̳������")
		template.html(2)=Replace(template.html(2),"{$stopreadme}","����̳�Ѿ�����������������������")
		Response.Write Template.html(2)
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
	Case "plus"
		Dvbbs.Stats=action&"-"&Template.Strings(0)
		Dvbbs.head()
		Dvbbs.showtoptable()
		Dvbbs.Head_var 0,"",Template.Strings(0),""
		template.html(0)=Replace(template.html(0),"{$color}",Dvbbs.mainsetting(1))
		template.html(0)=Replace(template.html(0),"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		template.html(0)=Replace(template.html(0),"{$action}","ʹ����̳���")
		template.html(0)=Replace(template.html(0),"{$ErrCount}",1)
		template.html(0)=Replace(template.html(0),"{$ErrString}",Request("ErrCodes"))
		Response.Write Template.html(0)
		If dvbbs.userid=0 Then 
			Response.Write Template.html(1)
		End If
		Dvbbs.ActiveOnline()
		Dvbbs.footer()	
	Case Else
		Dvbbs.Stats = Action & Template.Strings(0)
		Dvbbs.head()
		Dvbbs.showtoptable()
		Dvbbs.Head_var 0,"",Template.Strings(0),""
		template.html(0)=Replace(template.html(0),"{$color}",Dvbbs.mainsetting(1))
		template.html(0)=Replace(template.html(0),"{$errtitle}",Dvbbs.Forum_Info(0)&"-"&Dvbbs.Stats)
		template.html(0)=Replace(template.html(0),"{$action}",action)
		template.html(0)=Replace(template.html(0),"{$ErrCount}",ErrCount)
		template.html(0)=Replace(template.html(0),"{$ErrString}",ErrString)
		Response.Write Template.html(0)
		If dvbbs.userid=0 Then 
			Response.Write Template.html(1)
		End If
		Dvbbs.ActiveOnline()
		Dvbbs.footer()
End Select
Function Stopreadme()
	Dim Setting
	Setting=Dvbbs.CacheData(1,0)
	Setting=split(Setting,"|||")
	Stopreadme=Setting(5)
End Function 
Function  ErrCount()
	Dim ErrCodes,i
	ErrCount=0
	ErrCodes=Request("ErrCodes")
	If ErrCodes<>"" Then
		ErrCodes=Split(ErrCodes,",")
		For i=0 to UBound(ErrCodes)
			If IsNumeric(ErrCodes(i)) Then 
				If i=0 Then
				ErrString=ErrString&"<li>"&Template.Strings(ErrCodes(i))
				Else
				ErrString=ErrString&"<br><li>"&Template.Strings(ErrCodes(i))
				End If
				ErrCount=ErrCount+1
			End If 
		Next 
	End If 
End Function 
%>