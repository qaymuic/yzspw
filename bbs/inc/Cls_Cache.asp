<%
'论坛主缓存数据
Dim Dv_Cache
Set Dv_Cache=New Cls_Cache
'Set Dv_Cache=Server.CreateObject("webserver.Cls_cache")
Dv_Cache.Reloadtime=14400
Dv_Cache.CacheName="Dvbbs"
'页面临时数据
Dim Page_Cache
Set Page_Cache=New Cls_Cache
Page_Cache.Reloadtime=0.5
Page_Cache.CacheName="pages"
Dim templates_Cache
Set templates_Cache=New Cls_Cache
templates_Cache.CacheName="templates"
Class Cls_Cache
	Rem ==================使用说明=================================================================================
	Rem = 本类模块是动网先锋原创，作者：迷城浪子。如采用本类模块，请不要去掉这个说明。这段注释不会影响执行的速度。=
	Rem = 作用：缓存和缓存管理类                                                                                  =
	Rem = 公有变量：Reloadtime 过期时间（单位为分钟）缺省值为14400,                                               =
	Rem = MaxCount 缓存对象的最大值，超过则自动删除使用次数少的对象。缺省值为300                                  =
	Rem = CacheName 缓存组的总名称，缺省值为"Dvbbs",如果一个站点中有超过一个缓存组，则需要外部改变这个值。        =
	Rem = 属性:Name 定义缓存对象名称，只写属性。                                                                  =
	Rem = 属性:value 读取和写入缓存数据。                                                                         = 
	Rem = 函数：ObjIsEmpty()判断当前缓存是否过期。                                                                =
	Rem = 方法：DelCahe(MyCaheName)手工删除一个缓存对象，参数是缓存对象的名称。                                   =
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