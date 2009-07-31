<%
'������վ���溯��
'CacheFlag��CacheTime����config.asp��������

' �������
Function ClearCache()
	Dim objCache
	Application.lock
	For each objCache in Application.contents
		If Cstr(left(objCache, len(CacheFlag))) = Cstr(CacheFlag) Then Application.contents.Remove (objCache)
	Next
	Application.unlock
End Function


' ���û���
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


' ��ȡ����
Function GetCache(Byval cacheName)
	dim cacheData
	cacheName = LCase(FilterStr(cacheName))
	cacheData = Application(Cacheflag & cacheName)
	If IsArray(cacheData) Then GetCache = cacheData(0) Else GetCache = ""
End Function


' ��⻺��
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