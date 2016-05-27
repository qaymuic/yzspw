<%
'��̳����������
Dim Dv_Cache
Set Dv_Cache=New Cls_Cache
'Set Dv_Cache=Server.CreateObject("webserver.Cls_cache")
Dv_Cache.Reloadtime=14400
Dv_Cache.CacheName="Dvbbs"
'ҳ����ʱ����
Dim Page_Cache
Set Page_Cache=New Cls_Cache
Page_Cache.Reloadtime=0.5
Page_Cache.CacheName="pages"
Dim templates_Cache
Set templates_Cache=New Cls_Cache
templates_Cache.CacheName="templates"
Class Cls_Cache
	Rem ==================ʹ��˵��=================================================================================
	Rem = ����ģ���Ƕ����ȷ�ԭ�������ߣ��Գ����ӡ�����ñ���ģ�飬�벻Ҫȥ�����˵�������ע�Ͳ���Ӱ��ִ�е��ٶȡ�=
	Rem = ���ã�����ͻ��������                                                                                  =
	Rem = ���б�����Reloadtime ����ʱ�䣨��λΪ���ӣ�ȱʡֵΪ14400,                                               =
	Rem = MaxCount �����������ֵ���������Զ�ɾ��ʹ�ô����ٵĶ���ȱʡֵΪ300                                  =
	Rem = CacheName ������������ƣ�ȱʡֵΪ"Dvbbs",���һ��վ�����г���һ�������飬����Ҫ�ⲿ�ı����ֵ��        =
	Rem = ����:Name ���建��������ƣ�ֻд���ԡ�                                                                  =
	Rem = ����:value ��ȡ��д�뻺�����ݡ�                                                                         = 
	Rem = ������ObjIsEmpty()�жϵ�ǰ�����Ƿ���ڡ�                                                                =
	Rem = ������DelCahe(MyCaheName)�ֹ�ɾ��һ��������󣬲����ǻ����������ơ�                                   =
	Rem ===========================================================================================================
	Public Reloadtime,MaxCount,CacheName
	Private LocalCacheName,CacheData,DelCount
	Private Sub Class_Initialize()
		Reloadtime=14400
		CacheName="Dvbbs"
	End Sub
	Private Sub SetCache(SetName,NewValue)
		Application.Lock
		Application(SetName) = NewValue
		Application.unLock
	End Sub 
	Private Sub makeEmpty(SetName)
		Application.Lock
		Application(SetName) = Empty
		Application.unLock
	End Sub 
	Public  Property Let Name(ByVal vNewValue)
		LocalCacheName=LCase(vNewValue)
	End Property
	Public  Property Let Value(ByVal vNewValue)
		If LocalCacheName<>"" Then 
			CacheData=Application(CacheName&"_"&LocalCacheName)
			If IsArray(CacheData)  Then
				CacheData(0)=vNewValue
				CacheData(1)=Now()
			Else
				ReDim CacheData(2)
				CacheData(0)=vNewValue
				CacheData(1)=Now()
			End If
			SetCache CacheName&"_"&LocalCacheName,CacheData
		Else
			Err.Raise vbObjectError + 1, "DvbbsCacheServer", " please change the CacheName."
		End If		
	End Property
	Public Property Get Value()
		If LocalCacheName<>"" Then 
			CacheData=Application(CacheName&"_"&LocalCacheName)	
			If IsArray(CacheData) Then
				Value=CacheData(0)
			Else
				Err.Raise vbObjectError + 1, "DvbbsCacheServer", " The CacheData Is Empty."
			End If
		Else
			Err.Raise vbObjectError + 1, "DvbbsCacheServer", " please change the CacheName."
		End If
	End Property
	Public Function ObjIsEmpty()
		ObjIsEmpty=True
		CacheData=Application(CacheName&"_"&LocalCacheName)
		If Not IsArray(CacheData) Then Exit Function
		If Not IsDate(CacheData(1)) Then Exit Function
		If DateDiff("s",CDate(CacheData(1)),Now()) < 60*Reloadtime  Then
			ObjIsEmpty=False
		End If
	End Function
	Public Sub DelCahe(MyCaheName)
		makeEmpty(CacheName&"_"&MyCaheName)
	End Sub
End Class
%>