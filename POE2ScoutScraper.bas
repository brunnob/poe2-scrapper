'*******************************************************************************
' POE2 Scout Price Scraper for Excel
' Version: 3.0.1
' Description: Fetch real-time POE2 item prices with auto-refresh support
' Author: POE2 Community
' License: MIT
' Data Source: https://poe2scout.com
'
' NEW IN v3.0.1:
' - Fixed base URL to api.poe2scout.com (canonical, avoids legacy proxy)
' - Added retry logic with backoff for HTTP 429 (rate limit)
' - Added small delay between page requests to respect rate limits
'
' NEW IN v3.0:
' - Migrated to new API structure: /poe2/Leagues/{league}/...
' - League is now a URL path segment (was a query param)
' - API response is a pagination envelope: {Items:[...], Pages:N, CurrentPage:N}
' - Field names are PascalCase: Text, CurrentPrice, CategoryApiId, PriceLogs
' - Two separate endpoints: Currencies/ByCategory and Uniques/ByCategory
' - Full automatic pagination support
' - Two independent caches: currency items and unique items
' - All public function signatures preserved
'*******************************************************************************

Option Explicit

Private cachedCurrencyData As Collection
Private cachedUniqueData As Collection
Private cacheCurrencyLeague As String
Private cacheUniqueLeague As String
Private cacheCurrencyTime As Date
Private cacheUniqueTime As Date
Private Const CACHE_DURATION_MINUTES As Integer = 5
Private Const PER_PAGE As Integer = 100
Private Const API_BASE As String = "https://api.poe2scout.com/poe2/Leagues/"

'*******************************************************************************
' PUBLIC FUNCTIONS
'*******************************************************************************

Public Function getPOE2Price(ByVal itemName As String, _
                             Optional ByVal category As String = "", _
                             Optional ByVal league As String = "Fate of the Vaal") As Variant
    On Error GoTo ErrorHandler

    If Trim(itemName) = "" Then
        getPOE2Price = "Error: Item name required"
        Exit Function
    End If

    Dim items As Collection
    Set items = FetchAndFilter(league, category)

    Dim foundItem As Object
    Set foundItem = FindItem(items, itemName)

    If foundItem Is Nothing Then
        getPOE2Price = "Error: Item '" & itemName & "' not found"
    Else
        On Error Resume Next
        getPOE2Price = ParseNumberFromJSON(GetDictValue(foundItem, "CurrentPrice", "0"))
        If Err.Number <> 0 Then getPOE2Price = "Error: Invalid price data"
        On Error GoTo ErrorHandler
    End If
    Exit Function
ErrorHandler:
    getPOE2Price = "Error: " & Err.Description & " (Code: " & Err.Number & ")"
End Function

Public Function getPOE2ItemDetails(ByVal itemName As String, _
                                   Optional ByVal category As String = "", _
                                   Optional ByVal league As String = "Fate of the Vaal") As Variant
    On Error GoTo ErrorHandler

    If Trim(itemName) = "" Then
        getPOE2ItemDetails = Array("Error", "Item name required", "", "")
        Exit Function
    End If

    Dim items As Collection
    Set items = FetchAndFilter(league, category)

    Dim foundItem As Object
    Set foundItem = FindItem(items, itemName)

    If foundItem Is Nothing Then
        getPOE2ItemDetails = Array("Error", "Item not found", "", "")
    Else
        Dim result(0 To 3) As Variant
        result(0) = GetDictValue(foundItem, "Text", "Unknown")
        result(1) = ParseNumberFromJSON(GetDictValue(foundItem, "CurrentPrice", "0"))
        result(2) = GetQuantity(foundItem)
        result(3) = GetDictValue(foundItem, "CategoryApiId", "Unknown")
        getPOE2ItemDetails = result
    End If
    Exit Function
ErrorHandler:
    getPOE2ItemDetails = Array("Error", Err.Description, "", "")
End Function

Public Function getPOE2Items(Optional ByVal category As String = "", _
                             Optional ByVal league As String = "Fate of the Vaal") As Variant
    On Error GoTo ErrorHandler

    Dim items As Collection
    Set items = FetchAndFilter(league, category)

    If items Is Nothing Or items.Count = 0 Then
        getPOE2Items = Array(Array("Error", "No items found", "", ""))
        Exit Function
    End If

    Dim sortedItems As Collection
    Set sortedItems = SortByPrice(items)

    Dim result() As Variant
    ReDim result(0 To sortedItems.Count, 0 To 3)

    result(0, 0) = "Item Name"
    result(0, 1) = "Current Price"
    result(0, 2) = "Quantity"
    result(0, 3) = "Category"

    Dim i As Long
    Dim item As Object
    For i = 1 To sortedItems.Count
        Set item = sortedItems(i)
        result(i, 0) = GetDictValue(item, "Text", "Unknown")
        result(i, 1) = ParseNumberFromJSON(GetDictValue(item, "CurrentPrice", "0"))
        result(i, 2) = GetQuantity(item)
        result(i, 3) = GetDictValue(item, "CategoryApiId", "Unknown")
    Next i

    getPOE2Items = result
    Exit Function
ErrorHandler:
    getPOE2Items = Array(Array("Error", Err.Description, "", ""))
End Function

Public Function getPOE2Categories(Optional ByVal league As String = "Fate of the Vaal") As Variant
    On Error GoTo ErrorHandler

    Dim allItems As Collection
    Set allItems = FetchAllItems(league)

    Dim categoryCounts As Object
    Set categoryCounts = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim item As Object
    Dim cat As String

    For i = 1 To allItems.Count
        Set item = allItems(i)
        cat = GetDictValue(item, "CategoryApiId", "unknown")
        If categoryCounts.Exists(cat) Then
            categoryCounts(cat) = categoryCounts(cat) + 1
        Else
            categoryCounts.Add cat, 1
        End If
    Next i

    Dim result() As Variant
    ReDim result(0 To categoryCounts.Count, 0 To 1)
    result(0, 0) = "Category"
    result(0, 1) = "Item Count"

    Dim keys As Variant
    keys = categoryCounts.keys
    For i = 0 To UBound(keys)
        result(i + 1, 0) = keys(i)
        result(i + 1, 1) = categoryCounts(keys(i))
    Next i

    getPOE2Categories = result
    Exit Function
ErrorHandler:
    getPOE2Categories = Array(Array("Error", Err.Description))
End Function

'*******************************************************************************
' QUICK CATEGORY FUNCTIONS
'*******************************************************************************

Public Function getPOE2Currency(Optional ByVal league As String = "Fate of the Vaal") As Variant
    getPOE2Currency = getPOE2Items("currency", league)
End Function

Public Function getPOE2Fragments(Optional ByVal league As String = "Fate of the Vaal") As Variant
    getPOE2Fragments = getPOE2Items("fragments", league)
End Function

Public Function getPOE2Runes(Optional ByVal league As String = "Fate of the Vaal") As Variant
    getPOE2Runes = getPOE2Items("runes", league)
End Function

Public Function getPOE2Talismans(Optional ByVal league As String = "Fate of the Vaal") As Variant
    getPOE2Talismans = getPOE2Items("talismans", league)
End Function

Public Function getPOE2Essences(Optional ByVal league As String = "Fate of the Vaal") As Variant
    getPOE2Essences = getPOE2Items("essences", league)
End Function

Public Function getPOE2Accessories(Optional ByVal league As String = "Fate of the Vaal") As Variant
    getPOE2Accessories = getPOE2Items("accessory", league)
End Function

Public Function getPOE2Armour(Optional ByVal league As String = "Fate of the Vaal") As Variant
    getPOE2Armour = getPOE2Items("armour", league)
End Function

Public Function getPOE2Weapons(Optional ByVal league As String = "Fate of the Vaal") As Variant
    getPOE2Weapons = getPOE2Items("weapon", league)
End Function

'*******************************************************************************
' REFRESH MACROS (USER-CALLABLE)
'*******************************************************************************

Public Sub AtualizarPrecos()
    On Error Resume Next
    ClearCache
    Application.CalculateFull
    MsgBox "Precos atualizados com sucesso!" & vbCrLf & _
           "Ultima atualizacao: " & Format(Now, "dd/mm/yyyy hh:nn:ss"), _
           vbInformation, "POE2 Scout - Atualizacao"
End Sub

Public Sub AtualizarPlanilhaAtiva()
    On Error Resume Next
    ClearCache
    ActiveSheet.Calculate
    MsgBox "Planilha atualizada!" & vbCrLf & _
           "Ultima atualizacao: " & Format(Now, "dd/mm/yyyy hh:nn:ss"), _
           vbInformation, "POE2 Scout - Atualizacao"
End Sub

Public Sub AtualizarSelecao()
    On Error Resume Next
    ClearCache
    Selection.Calculate
    MsgBox "Selecao atualizada!" & vbCrLf & _
           "Ultima atualizacao: " & Format(Now, "dd/mm/yyyy hh:nn:ss"), _
           vbInformation, "POE2 Scout - Atualizacao"
End Sub

Public Sub ForcarAtualizacao()
    On Error Resume Next
    Application.ScreenUpdating = False
    ClearCache
    Application.CalculateFullRebuild
    Application.ScreenUpdating = True
    MsgBox "Atualizacao forcada concluida!" & vbCrLf & _
           "Cache limpo e todas as formulas recalculadas." & vbCrLf & _
           "Ultima atualizacao: " & Format(Now, "dd/mm/yyyy hh:nn:ss"), _
           vbInformation, "POE2 Scout - Atualizacao Forcada"
End Sub

Public Sub ConfigurarAtualizacaoAutomatica()
    Dim minutos As String
    Dim intervalo As Integer

    minutos = InputBox("Digite o intervalo de atualizacao automatica (em minutos):" & vbCrLf & _
                       "Recomendado: 5-15 minutos" & vbCrLf & vbCrLf & _
                       "Digite 0 para desabilitar.", _
                       "POE2 Scout - Atualizacao Automatica", "5")

    If minutos = "" Then Exit Sub
    intervalo = Val(minutos)

    If intervalo > 0 Then
        Application.OnTime Now + TimeValue("00:" & Format(intervalo, "00") & ":00"), "AtualizarPrecos"
        MsgBox "Atualizacao automatica configurada!" & vbCrLf & _
               "Intervalo: " & intervalo & " minutos" & vbCrLf & vbCrLf & _
               "Proxima atualizacao: " & Format(Now + TimeValue("00:" & Format(intervalo, "00") & ":00"), "hh:nn:ss"), _
               vbInformation, "POE2 Scout"
    Else
        MsgBox "Atualizacao automatica desabilitada.", vbInformation, "POE2 Scout"
    End If
End Sub

'*******************************************************************************
' CORE DATA FETCHING (PRIVATE)
'*******************************************************************************

Private Function FetchAndFilter(ByVal league As String, ByVal category As String) As Collection
    Dim all As Collection

    If Trim(category) = "" Then
        Set all = FetchAllItems(league)
    ElseIf IsCurrencyCategory(category) Then
        Set all = FetchCurrencyItems(league)
    Else
        Set all = FetchUniqueItems(league)
    End If

    Set FetchAndFilter = FilterByCategory(all, category)
End Function

Private Function FetchAllItems(ByVal league As String) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim curr As Collection
    Set curr = FetchCurrencyItems(league)

    Dim uniq As Collection
    Set uniq = FetchUniqueItems(league)

    Dim i As Long
    For i = 1 To curr.Count
        result.Add curr(i)
    Next i
    For i = 1 To uniq.Count
        result.Add uniq(i)
    Next i

    Set FetchAllItems = result
End Function

Private Function FetchCurrencyItems(ByVal league As String) As Collection
    If Not cachedCurrencyData Is Nothing Then
        If cacheCurrencyLeague = league Then
            If DateDiff("n", cacheCurrencyTime, Now) < CACHE_DURATION_MINUTES Then
                Set FetchCurrencyItems = cachedCurrencyData
                Exit Function
            End If
        End If
    End If

    Dim url As String
    url = API_BASE & URLEncode(league) & "/Currencies/ByCategory"

    Dim result As Collection
    Set result = FetchAllPages(url)
    If result Is Nothing Then Set result = New Collection

    Set cachedCurrencyData = result
    cacheCurrencyLeague = league
    cacheCurrencyTime = Now

    Set FetchCurrencyItems = result
End Function

Private Function FetchUniqueItems(ByVal league As String) As Collection
    If Not cachedUniqueData Is Nothing Then
        If cacheUniqueLeague = league Then
            If DateDiff("n", cacheUniqueTime, Now) < CACHE_DURATION_MINUTES Then
                Set FetchUniqueItems = cachedUniqueData
                Exit Function
            End If
        End If
    End If

    Dim url As String
    url = API_BASE & URLEncode(league) & "/Uniques/ByCategory"

    Dim result As Collection
    Set result = FetchAllPages(url)
    If result Is Nothing Then Set result = New Collection

    Set cachedUniqueData = result
    cacheUniqueLeague = league
    cacheUniqueTime = Now

    Set FetchUniqueItems = result
End Function

' Fetches all pages with retry on 429 and a small delay between pages.
Private Function FetchAllPages(ByVal baseUrl As String) As Collection
    On Error GoTo ErrorHandler

    Dim result As Collection
    Set result = New Collection

    Dim page As Long
    Dim totalPages As Long
    page = 1
    totalPages = 1

    Do While page <= totalPages
        Dim pageUrl As String
        pageUrl = baseUrl & "?Page=" & page & "&PerPage=" & PER_PAGE

        Dim responseText As String
        responseText = HttpGetWithRetry(pageUrl)
        If Len(responseText) = 0 Then Exit Do

        Dim parsedPages As Long
        Dim pageItems As Collection
        Set pageItems = ParseEnvelopeResponse(responseText, parsedPages)

        If page = 1 Then
            totalPages = parsedPages
            If totalPages < 1 Then totalPages = 1
        End If

        If Not pageItems Is Nothing Then
            Dim i As Long
            For i = 1 To pageItems.Count
                result.Add pageItems(i)
            Next i
        End If

        page = page + 1
        If page > 50 Then Exit Do  ' safety cap

        ' Brief pause between pages to respect rate limits
        If page <= totalPages Then
            Dim waitUntil As Date
            waitUntil = Now + TimeSerial(0, 0, 1)
            Do While Now < waitUntil
                DoEvents
            Loop
        End If
    Loop

    Set FetchAllPages = result
    Exit Function
ErrorHandler:
    Debug.Print "FetchAllPages Error: " & Err.Description
    Set FetchAllPages = result
End Function

' Makes an HTTP GET request with up to 3 retries on HTTP 429 (rate limit).
Private Function HttpGetWithRetry(ByVal url As String) As String
    Dim attempt As Integer
    Dim waitSeconds As Integer

    For attempt = 1 To 3
        Dim response As String
        Dim httpStatus As Long
        httpStatus = HttpGetRaw(url, response)

        Select Case httpStatus
            Case 200
                HttpGetWithRetry = response
                Exit Function
            Case 429
                ' Rate limited — wait and retry
                waitSeconds = attempt * 3  ' 3s, 6s, 9s
                Debug.Print "HTTP 429 on attempt " & attempt & ", waiting " & waitSeconds & "s before retry..."
                Dim waitUntil As Date
                waitUntil = Now + TimeSerial(0, 0, waitSeconds)
                Do While Now < waitUntil
                    DoEvents
                Loop
            Case 0
                ' Network error — stop retrying
                HttpGetWithRetry = ""
                Exit Function
            Case Else
                Debug.Print "HTTP " & httpStatus & " for " & url
                HttpGetWithRetry = ""
                Exit Function
        End Select
    Next attempt

    Debug.Print "All retries exhausted for " & url
    HttpGetWithRetry = ""
End Function

' Returns HTTP status code and sets responseBody. Returns 0 on network error.
Private Function HttpGetRaw(ByVal url As String, ByRef responseBody As String) As Long
    On Error GoTo ErrorHandler

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "POE2-Excel-Scraper/3.0"
    http.setRequestHeader "Accept", "application/json"
    http.Send

    HttpGetRaw = http.Status
    If http.Status = 200 Then responseBody = http.responseText
    Exit Function
ErrorHandler:
    Debug.Print "HttpGetRaw Error: " & Err.Description
    HttpGetRaw = 0
    responseBody = ""
End Function

' Parses {Items:[...], Pages:N} envelope. outPages receives the total page count.
Private Function ParseEnvelopeResponse(ByVal jsonText As String, ByRef outPages As Long) As Collection
    On Error GoTo TryFallback

    outPages = 1

    Dim sc As Object
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"

    Dim jsObj As Object
    Set jsObj = sc.Eval("(" & jsonText & ")")

    On Error Resume Next
    outPages = CLng(jsObj.Pages)
    If Err.Number <> 0 Or outPages < 1 Then outPages = 1
    Err.Clear
    On Error GoTo TryFallback

    Dim jsItems As Object
    Set jsItems = jsObj.Items

    Dim result As Collection
    Set result = New Collection

    Dim i As Long
    For i = 0 To jsItems.length - 1
        result.Add ConvertJSObject(jsItems(i))
    Next i

    Set ParseEnvelopeResponse = result
    Exit Function

TryFallback:
    Debug.Print "ScriptControl failed, using fallback parser"
    Set ParseEnvelopeResponse = ParseEnvelopeFallback(jsonText, outPages)
End Function

Private Function ConvertJSObject(ByVal jsObj As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    dict.Add "Text", CStr(jsObj.Text)
    dict.Add "CurrentPrice", CStr(jsObj.CurrentPrice)
    dict.Add "CategoryApiId", CStr(jsObj.CategoryApiId)
    dict.Add "ApiId", CStr(jsObj.ApiId)
    dict.Add "IconUrl", CStr(jsObj.IconUrl)
    dict.Add "CurrentQuantity", CStr(jsObj.CurrentQuantity)
    dict.Add "PriceLogs", jsObj.PriceLogs
    On Error GoTo 0

    Set ConvertJSObject = dict
End Function

'*******************************************************************************
' FALLBACK JSON PARSER (PRIVATE)
'*******************************************************************************

Private Function ParseEnvelopeFallback(ByVal jsonText As String, ByRef outPages As Long) As Collection
    On Error GoTo ErrorHandler

    outPages = 1

    Dim pagesStr As String
    pagesStr = ExtractJSONValue(jsonText, "Pages")
    If IsNumeric(pagesStr) And pagesStr <> "" Then outPages = CLng(pagesStr)

    Dim itemsKeyPos As Long
    itemsKeyPos = InStr(jsonText, """Items"":")
    If itemsKeyPos = 0 Then itemsKeyPos = InStr(jsonText, """items"":")

    If itemsKeyPos = 0 Then
        If Left(Trim(jsonText), 1) = "[" Then
            Set ParseEnvelopeFallback = ParseBareArray(jsonText)
        Else
            Set ParseEnvelopeFallback = New Collection
        End If
        Exit Function
    End If

    Dim arrStart As Long
    arrStart = InStr(itemsKeyPos, jsonText, "[")
    If arrStart = 0 Then
        Set ParseEnvelopeFallback = New Collection
        Exit Function
    End If

    Dim depth As Long
    Dim arrEnd As Long
    Dim i As Long
    Dim insideStr As Boolean
    Dim prevCh As String
    depth = 0
    insideStr = False
    prevCh = ""

    For i = arrStart To Len(jsonText)
        Dim ch As String
        ch = Mid(jsonText, i, 1)
        If ch = """" And prevCh <> "\" Then insideStr = Not insideStr
        If Not insideStr Then
            If ch = "[" Then depth = depth + 1
            If ch = "]" Then
                depth = depth - 1
                If depth = 0 Then arrEnd = i: Exit For
            End If
        End If
        prevCh = ch
    Next i

    If arrEnd = 0 Then
        Set ParseEnvelopeFallback = New Collection
        Exit Function
    End If

    Set ParseEnvelopeFallback = ParseBareArray(Mid(jsonText, arrStart, arrEnd - arrStart + 1))
    Exit Function
ErrorHandler:
    Debug.Print "ParseEnvelopeFallback Error: " & Err.Description
    Set ParseEnvelopeFallback = New Collection
End Function

Private Function ParseBareArray(ByVal jsonText As String) As Collection
    Dim result As Collection
    Set result = New Collection

    jsonText = Trim(jsonText)
    If Left(jsonText, 1) = "[" Then jsonText = Mid(jsonText, 2)
    If Right(jsonText, 1) = "]" Then jsonText = Left(jsonText, Len(jsonText) - 1)

    Dim objects() As String
    objects = SplitJSONObjects(jsonText)

    Dim i As Long
    For i = 0 To UBound(objects)
        Dim item As Object
        Set item = ParseJSONObject(objects(i))
        If Not item Is Nothing Then result.Add item
    Next i

    Set ParseBareArray = result
End Function

Private Function SplitJSONObjects(ByVal jsonText As String) As String()
    Dim result() As String
    Dim objCount As Long
    Dim depth As Long
    Dim currentObj As String
    Dim i As Long
    Dim ch As String
    Dim insideStr As Boolean
    Dim prevCh As String

    objCount = 0
    depth = 0
    currentObj = ""
    insideStr = False
    prevCh = ""

    For i = 1 To Len(jsonText)
        ch = Mid(jsonText, i, 1)
        If ch = """" And prevCh <> "\" Then insideStr = Not insideStr
        If Not insideStr Then
            If ch = "{" Then depth = depth + 1
            If ch = "}" Then depth = depth - 1
        End If
        currentObj = currentObj & ch
        If depth = 0 And Len(currentObj) > 2 Then
            ReDim Preserve result(objCount)
            result(objCount) = currentObj
            objCount = objCount + 1
            currentObj = ""
        End If
        prevCh = ch
    Next i

    SplitJSONObjects = result
End Function

Private Function ParseJSONObject(ByVal jsonText As String) As Object
    On Error Resume Next

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    dict.Add "Text", ExtractJSONValue(jsonText, "Text")
    dict.Add "CurrentPrice", ExtractJSONValue(jsonText, "CurrentPrice")
    dict.Add "CategoryApiId", ExtractJSONValue(jsonText, "CategoryApiId")
    dict.Add "ApiId", ExtractJSONValue(jsonText, "ApiId")
    dict.Add "CurrentQuantity", ExtractJSONValue(jsonText, "CurrentQuantity")

    If Err.Number <> 0 Then
        Set ParseJSONObject = Nothing
    Else
        Set ParseJSONObject = dict
    End If
End Function

Private Function ExtractJSONValue(ByVal jsonText As String, ByVal key As String) As String
    On Error Resume Next

    Dim pattern As String
    pattern = """" & key & """:"

    Dim startPos As Long
    startPos = InStr(jsonText, pattern)
    If startPos = 0 Then
        ExtractJSONValue = ""
        Exit Function
    End If

    startPos = startPos + Len(pattern)
    Do While Mid(jsonText, startPos, 1) = " "
        startPos = startPos + 1
    Loop

    Dim ch As String
    ch = Mid(jsonText, startPos, 1)

    Dim endPos As Long
    If ch = """" Then
        startPos = startPos + 1
        endPos = InStr(startPos, jsonText, """")
        ExtractJSONValue = Mid(jsonText, startPos, endPos - startPos)
    Else
        endPos = startPos
        Do While endPos <= Len(jsonText)
            ch = Mid(jsonText, endPos, 1)
            If ch = "," Or ch = "}" Or ch = "]" Then Exit Do
            endPos = endPos + 1
        Loop
        ExtractJSONValue = Trim(Mid(jsonText, startPos, endPos - startPos))
    End If
End Function

'*******************************************************************************
' UTILITY FUNCTIONS (PRIVATE)
'*******************************************************************************

Private Function IsCurrencyCategory(ByVal category As String) As Boolean
    Select Case LCase(Trim(category))
        Case "currency", "fragments", "runes", "talismans", "essences", _
             "catalysts", "oils", "incubators", "scarabs", "delirium", _
             "breach", "expedition", "ritual", "ultimatum"
            IsCurrencyCategory = True
        Case Else
            IsCurrencyCategory = False
    End Select
End Function

Private Function URLEncode(ByVal text As String) As String
    URLEncode = Replace(text, " ", "%20")
    URLEncode = Replace(URLEncode, "&", "%26")
    URLEncode = Replace(URLEncode, "+", "%2B")
End Function

Private Function FilterByCategory(ByVal data As Collection, ByVal category As String) As Collection
    Dim filtered As Collection
    Set filtered = New Collection

    Dim i As Long
    If Trim(category) = "" Then
        For i = 1 To data.Count
            filtered.Add data(i)
        Next i
    Else
        Dim item As Object
        For i = 1 To data.Count
            Set item = data(i)
            If LCase(GetDictValue(item, "CategoryApiId", "")) = LCase(Trim(category)) Then
                filtered.Add item
            End If
        Next i
    End If

    Set FilterByCategory = filtered
End Function

Private Function FindItem(ByVal items As Collection, ByVal itemName As String) As Object
    Dim itemNameLower As String
    itemNameLower = LCase(Trim(itemName))

    Dim i As Long
    Dim item As Object
    Dim itemText As String

    For i = 1 To items.Count
        Set item = items(i)
        itemText = LCase(GetDictValue(item, "Text", ""))
        If itemText = itemNameLower Then
            Set FindItem = item
            Exit Function
        End If
    Next i

    For i = 1 To items.Count
        Set item = items(i)
        itemText = LCase(GetDictValue(item, "Text", ""))
        If InStr(itemText, itemNameLower) > 0 Then
            Set FindItem = item
            Exit Function
        End If
    Next i

    Set FindItem = Nothing
End Function

Private Function SortByPrice(ByVal items As Collection) As Collection
    Dim sorted As Collection
    Set sorted = New Collection

    If items.Count = 0 Then
        Set SortByPrice = sorted
        Exit Function
    End If

    Dim arr() As Variant
    ReDim arr(1 To items.Count)

    Dim i As Long
    For i = 1 To items.Count
        Set arr(i) = items(i)
    Next i

    Dim j As Long
    Dim temp As Object
    Dim swapped As Boolean
    Dim price1 As Double, price2 As Double

    For i = 1 To UBound(arr) - 1
        swapped = False
        For j = 1 To UBound(arr) - i
            price1 = ParseNumberFromJSON(GetDictValue(arr(j), "CurrentPrice", "0"))
            price2 = ParseNumberFromJSON(GetDictValue(arr(j + 1), "CurrentPrice", "0"))
            If price1 < price2 Then
                Set temp = arr(j)
                Set arr(j) = arr(j + 1)
                Set arr(j + 1) = temp
                swapped = True
            End If
        Next j
        If Not swapped Then Exit For
    Next i

    For i = 1 To UBound(arr)
        sorted.Add arr(i)
    Next i

    Set SortByPrice = sorted
End Function

Private Function GetDictValue(ByVal dict As Object, ByVal key As String, ByVal defaultValue As Variant) As Variant
    On Error Resume Next
    If dict.Exists(key) Then
        If IsObject(dict(key)) Then
            Set GetDictValue = dict(key)
        Else
            GetDictValue = dict(key)
        End If
    Else
        GetDictValue = defaultValue
    End If
    On Error GoTo 0
End Function

Private Function GetQuantity(ByVal item As Object) As Variant
    On Error Resume Next

    Dim qty As Variant
    qty = GetDictValue(item, "CurrentQuantity", "")
    If qty <> "" And IsNumeric(qty) And CDbl(qty) > 0 Then
        GetQuantity = CDbl(qty)
        Exit Function
    End If

    Dim priceLogs As Object
    Set priceLogs = item("PriceLogs")
    If Not priceLogs Is Nothing Then
        If priceLogs.length > 0 Then
            Dim latestLog As Object
            Set latestLog = priceLogs(0)
            If Not latestLog Is Nothing Then
                GetQuantity = latestLog("Quantity")
                Exit Function
            End If
        End If
    End If

    GetQuantity = "N/A"
    On Error GoTo 0
End Function

Private Function ParseNumberFromJSON(ByVal value As Variant) As Double
    On Error GoTo ErrorHandler

    If IsNumeric(value) And VarType(value) <> vbString Then
        ParseNumberFromJSON = CDbl(value)
        Exit Function
    End If

    Dim strValue As String
    strValue = Trim(CStr(value))
    strValue = Replace(strValue, " ", "")

    Dim dotPos As Long
    Dim integerPart As String
    Dim decimalPart As String

    dotPos = InStr(strValue, ".")
    If dotPos > 0 Then
        integerPart = Left(strValue, dotPos - 1)
        decimalPart = Mid(strValue, dotPos + 1)
    Else
        integerPart = strValue
        decimalPart = "0"
    End If

    Dim isNegative As Boolean
    If Left(integerPart, 1) = "-" Then
        isNegative = True
        integerPart = Mid(integerPart, 2)
    End If

    Dim intVal As Double
    Dim i As Long
    Dim digit As String
    intVal = 0
    For i = 1 To Len(integerPart)
        digit = Mid(integerPart, i, 1)
        If digit >= "0" And digit <= "9" Then
            intVal = intVal * 10 + (Asc(digit) - Asc("0"))
        End If
    Next i

    Dim decVal As Double
    Dim divisor As Double
    decVal = 0
    divisor = 10
    For i = 1 To Len(decimalPart)
        digit = Mid(decimalPart, i, 1)
        If digit >= "0" And digit <= "9" Then
            decVal = decVal + (Asc(digit) - Asc("0")) / divisor
            divisor = divisor * 10
        End If
    Next i

    Dim result As Double
    result = intVal + decVal
    If isNegative Then result = -result

    ParseNumberFromJSON = result
    Exit Function
ErrorHandler:
    ParseNumberFromJSON = 0
End Function

'*******************************************************************************
' PUBLIC UTILITIES
'*******************************************************************************

Public Sub ClearCache()
    Set cachedCurrencyData = Nothing
    Set cachedUniqueData = Nothing
    cacheCurrencyLeague = ""
    cacheUniqueLeague = ""
    cacheCurrencyTime = 0
    cacheUniqueTime = 0
    Debug.Print "Cache cleared at " & Now
End Sub

Public Sub TestAPIConnection()
    Debug.Print "Testing API connection (v3.0.1)..."
    ClearCache

    Dim currencies As Collection
    Set currencies = FetchCurrencyItems("Fate of the Vaal")
    If currencies Is Nothing Or currencies.Count = 0 Then
        Debug.Print "FAILED: Could not fetch currency data"
    Else
        Debug.Print "SUCCESS: Fetched " & currencies.Count & " currency items"
    End If

    Dim uniques As Collection
    Set uniques = FetchUniqueItems("Fate of the Vaal")
    If uniques Is Nothing Or uniques.Count = 0 Then
        Debug.Print "WARNING: No unique items fetched"
    Else
        Debug.Print "SUCCESS: Fetched " & uniques.Count & " unique items"
    End If

    Debug.Print "Testing getPOE2Price for Exalted Orb..."
    Dim testPrice As Variant
    testPrice = getPOE2Price("Exalted Orb", "currency")
    Debug.Print "Exalted Orb price: " & testPrice
End Sub
