<%
'设置网站缓存函数
'CacheFlag和CacheTime请在config.asp里面设置

' 清除缓存
Function ClearCache()
	Dim objCache
	Application.lock
	For each objCache in Application.contents
		If Cstr(left(objCache, len(CacheFlag))) = Cstr(CacheFlag) Then Application.contents.Remove (objCache)
	Next
	Application.unlock
End Function


' 设置缓存
Function SetCache(Byval cacheName, Byval cacheValue)
	Dim cacheData
	cacheName = LCase(FilterStr(cacheName))
	cacheData = Application(Cacheflag & cacheName)
	If IsArray(cacheData) Then
		cacheData(0) = cacheValue
		cacheData(1) = Now()
	Else
		Redim cacheData(2)
		cacheData(0) = cacheValue
		cacheData(1) = Now()
	End If
	Application.lock
	Application(CacheFlag & cacheName) = cacheData
	Application.unlock
End Function


' 获取缓存
Function GetCache(Byval cacheName)
	dim cacheData
	cacheName = LCase(FilterStr(cacheName))
	cacheData = Application(Cacheflag & cacheName)
	If IsArray(cacheData) Then GetCache = cacheData(0) Else GetCache = ""
End Function


' 检测缓存
Function ChkCache(Byval cacheName)
	dim cacheData
	ChkCache = false
	cacheName = LCase(FilterStr(cacheName))
	cacheData = Application(Cacheflag & cacheName)
	If Not IsArray(cacheData) Then Exit Function
	If Not IsDate(cacheData(1)) Then Exit Function
	If DateDIff("s", CDate(cacheData(1)), Now()) < 60 * CacheTime Then chkcache = true
End Function

%>