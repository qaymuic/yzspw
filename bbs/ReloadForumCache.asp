<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Option Explicit
Response.Buffer = True
%>
<!--#Include File="inc/Dv_ClsMain.asp"-->
<%
dim admin_flag
admin_flag=","
Dim CacheName,Dvbbs
Set Dvbbs= New Cls_Forum
CacheName=Dvbbs.CacheName
If InStr(session("flag"),admin_flag) >0  Then 
	Call delallcache()
End If
Function  GetallCache()
	Dim Cacheobj
	For Each Cacheobj in Application.Contents
	If CStr(Left(Cacheobj,Len(CacheName)+1))=CStr(CacheName&"_") Then	
		GetallCache=GetallCache&Cacheobj&","
	End If
	Next
End Function
Sub delallcache()
	Dim cachelist,i
	Cachelist=split(GetallCache(),",")
	If UBound(cachelist)>1 Then
		For i=0 to UBound(cachelist)-1
			DelCahe Cachelist(i)
			Response.Write "���� <b>"&Replace(cachelist(i),CacheName&"_","")&"</b> ���<br>"		
		Next
		Response.Write "������"
		Response.Write UBound(cachelist)-1
		Response.Write "���������<br>"	
	Else
		Response.Write "���ж����Ѿ����¡�"
	End If
End Sub 
Sub DelCahe(MyCaheName)
	Application.Lock
	Application.Contents.Remove(MyCaheName)
	Application.unLock
End Sub
%>