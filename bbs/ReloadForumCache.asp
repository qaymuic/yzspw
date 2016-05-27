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
			Response.Write "更新 <b>"&Replace(cachelist(i),CacheName&"_","")&"</b> 完成<br>"		
		Next
		Response.Write "更新了"
		Response.Write UBound(cachelist)-1
		Response.Write "个缓存对象<br>"	
	Else
		Response.Write "所有对象已经更新。"
	End If
End Sub 
Sub DelCahe(MyCaheName)
	Application.Lock
	Application.Contents.Remove(MyCaheName)
	Application.unLock
End Sub
%>