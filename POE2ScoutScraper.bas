'*******************************************************************************
' POE2 Scout Price Scraper for Excel
' Version: 3.1.0
' Description: Fetch real-time POE2 item prices with auto-refresh support
' Author: POE2 Community
' License: MIT
' Data Source: https://poe2scout.com
'
' NEW IN v3.1 (correct API mapping):
' - Correct endpoint: /api/poe2/Leagues/{league}/Items (single aggregated list)
' - No pagination, no Page/PerPage params (these caused HTTP 422)
' - Response is a plain JSON array (not a paginated envelope)
' - Field names are snake_case: text, current_price, category_api_id,
'   name, type, api_id, icon_url, item_id
' - Categories endpoint: /api/poe2/Leagues/{league}/Items/Categories
' - Retry with backoff on HTTP 429 (rate limit)
' - All public function signatures preserved
'*******************************************************************************

Option Explicit

Private cachedData As Collection
Private cacheLeague As String
Private cacheTimestamp As Date
Private Const CACHE_DURATION_MINUTES As Integer = 5
Private Const API_BASE As String = "https://poe2scout.com/api/poe2/Leagues/"

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

    Dim data As Collection
    Set data = FetchPOE2Data(league)

    If data Is Nothing Then
        getPOE2Price = "Error: Failed to fetch data from API"
        Exit Function
    End If

    Dim items As Collection
    Set items = FilterByCategory(data, category)

    Dim foundItem As Object
    Set foundItem = FindItem(items, itemName)

    If foundItem Is Nothing Then
        getPOE2Price = "Error: Item '" & itemName & "' not found"
    Else
        On Error Resume Next
        getPOE2Price = ParseNumberFromJSON(GetDictValue(foundItem, "current_price", "0"))
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

    Dim data As Collection
    Set data = FetchPOE2Data(league)

    If data Is Nothing Then
        getPOE2ItemDetails = Array("Error", "Failed to fetch data", "", "")
        Exit Function
    End If

    Dim items As Collection
    Set items = FilterByCategory(data, category)

    Dim foundItem As Object
    Set foundItem = FindItem(items, itemName)

    If foundItem Is Nothing Then
        getPOE2ItemDetails = Array("Error", "Item not found", "", "")
    Else
        Dim result(0 To 3) As Variant
        result(0) = GetDictValue(foundItem, "text", "Unknown")
        result(1) = ParseNumberFromJSON(GetDictValue(foundItem, "current_price", "0"))
        result(2) = GetQuantity(foundItem)
        result(3) = GetDictValue(foundItem, "category_api_id", "Unknown")
        getPOE2ItemDetails = result
    End If
    Exit Function
ErrorHandler:
    getPOE2ItemDetails = Array("Error", Err.Description, "", "")
End Function

Public Function getPOE2Items(Optional ByVal category As String = "", _
                             Optional ByVal league As String = "Fate of the Vaal") As Variant
    On Error GoTo ErrorHandler

    Dim data As Collection
    Set data = FetchPOE2Data(league)

    If data Is Nothing Then
        getPOE2Items = Array(Array("Error", "Failed to fetch data", "", ""))
        Exit Function
    End If

    Dim items As Collection
    Set items = FilterByCategory(data, category)

    If items.Count = 0 Then
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
        result(i, 0) = GetDictValue(item, "text", "Unknown")
        result(i, 1) = ParseNumberFromJSON(GetDictValue(item, "current_price", "0"))
        result(i, 2) = GetQuantity(item)
        result(i, 3) = GetDictValue(item, "category_api_id", "Unknown")
    Next i

    getPOE2Items = result
    Exit Function
ErrorHandler:
    getPOE2Items = Array(Array("Error", Err.Description, "", ""))
End Function

Public Function getPOE2Categories(Optional ByVal league As String = "Fate of the Vaal") As Variant
    On Error GoTo ErrorHandler

    Dim data As Collection
    Set data = FetchPOE2Data(league)

    If data Is Nothing Then
        getPOE2Categories = Array(Array("Error", "Failed to fetch data"))
        Exit Function
    End If

    Dim categoryCounts As Object
    Set categoryCounts = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim item As Object
    Dim cat As String

    For i = 1 To data.Count
        Set item = data(i)
        cat = GetDictValue(item, "category_api_id", "unknown")
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

' Fetches the full aggregated item list (currency + uniques) for a league.
Private Function FetchPOE2Data(ByVal league As String) As Collection
    On Error GoTo ErrorHandler

    ' Check cache
    If Not cachedData Is Nothing Then
        If cacheLeague = league Then
            If DateDiff("n", cacheTimestamp, Now) < CACHE_DURATION_MINUTES Then
                Set FetchPOE2Data = cachedData
                Exit Function
            End If
        End If
    End If

    Dim url As String
    url = API_BASE & URLEncode(league) & "/Items"

    Dim responseText As String
    responseText = HttpGetWithRetry(url)

    If Len(responseText) = 0 Then
        Set FetchPOE2Data = Nothing
        Exit Function
    End If

    Dim parsed As Collection
    Set parsed = ParseJSONResponse(responseText)

    If parsed Is Nothing Then
        Set FetchPOE2Data = Nothing
        Exit Function
    End If

    Set cachedData = parsed
    cacheLeague = league
    cacheTimestamp = Now

    Set FetchPOE2Data = parsed
    Exit Function
ErrorHandler:
    Debug.Print "FetchPOE2Data Error: " & Err.Description & " (Code: " & Err.Number & ")"
    Set FetchPOE2Data = Nothing
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
                waitSeconds = attempt * 3  ' 3s, 6s, 9s
                Debug.Print "HTTP 429 on attempt " & attempt & ", waiting " & waitSeconds & "s..."
                WaitSecs waitSeconds
            Case 0
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
    http.setRequestHeader "User-Agent", "POE2-Excel-Scraper/3.1 (contact: excel-addin)"
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

Private Sub WaitSecs(ByVal secs As Integer)
    Dim waitUntil As Date
    waitUntil = Now + TimeSerial(0, 0, secs)
    Do While Now < waitUntil
        DoEvents
    Loop
End Sub

' Parses a plain JSON array of item objects.
Private Function ParseJSONResponse(ByVal jsonText As String) As Collection
    On Error GoTo TryFallback

    Dim sc As Object
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"

    Dim jsArray As Object
    Set jsArray = sc.Eval("(" & jsonText & ")")

    Dim result As Collection
    Set result = New Collection

    Dim i As Long
    For i = 0 To jsArray.length - 1
        result.Add ConvertJSObject(jsArray(i))
    Next i

    Set ParseJSONResponse = result
    Exit Function

TryFallback:
    Debug.Print "ScriptControl failed, using fallback parser"
    Set ParseJSONResponse = ParseJSONAlternative(jsonText)
End Function

Private Function ConvertJSObject(ByVal jsObj As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    dict.Add "item_id", CStr(jsObj.item_id)
    dict.Add "api_id", CStr(jsObj.api_id)
    dict.Add "text", CStr(jsObj.text)
    dict.Add "name", CStr(jsObj.name)
    dict.Add "type", CStr(jsObj.type)
    dict.Add "category_api_id", CStr(jsObj.category_api_id)
    dict.Add "current_price", CStr(jsObj.current_price)
    dict.Add "icon_url", CStr(jsObj.icon_url)
    On Error GoTo 0

    Set ConvertJSObject = dict
End Function

'*******************************************************************************
' FALLBACK JSON PARSER (PRIVATE) - used if MSScriptControl is unavailable
'*******************************************************************************

Private Function ParseJSONAlternative(ByVal jsonText As String) As Collection
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

    Set ParseJSONAlternative = result
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

    dict.Add "text", ExtractJSONValue(jsonText, "text")
    dict.Add "current_price", ExtractJSONValue(jsonText, "current_price")
    dict.Add "category_api_id", ExtractJSONValue(jsonText, "category_api_id")
    dict.Add "name", ExtractJSONValue(jsonText, "name")
    dict.Add "type", ExtractJSONValue(jsonText, "type")
    dict.Add "api_id", ExtractJSONValue(jsonText, "api_id")
    dict.Add "item_id", ExtractJSONValue(jsonText, "item_id")

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
            If LCase(GetDictValue(item, "category_api_id", "")) = LCase(Trim(category)) Then
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

    ' Exact match first
    For i = 1 To items.Count
        Set item = items(i)
        itemText = LCase(GetDictValue(item, "text", ""))
        If itemText = itemNameLower Then
            Set FindItem = item
            Exit Function
        End If
    Next i

    ' Partial match
    For i = 1 To items.Count
        Set item = items(i)
        itemText = LCase(GetDictValue(item, "text", ""))
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
            price1 = ParseNumberFromJSON(GetDictValue(arr(j), "current_price", "0"))
            price2 = ParseNumberFromJSON(GetDictValue(arr(j + 1), "current_price", "0"))
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

' The aggregated /Items endpoint does not expose quantity, so this returns N/A.
Private Function GetQuantity(ByVal item As Object) As Variant
    On Error Resume Next
    Dim qty As Variant
    qty = GetDictValue(item, "current_quantity", "")
    If qty <> "" And IsNumeric(qty) And CDbl(qty) > 0 Then
        GetQuantity = CDbl(qty)
        Exit Function
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
    Set cachedData = Nothing
    cacheLeague = ""
    cacheTimestamp = 0
    Debug.Print "Cache cleared at " & Now
End Sub

Public Sub TestAPIConnection()
    Debug.Print "Testing API connection (v3.1.0)..."
    ClearCache

    Dim data As Collection
    Set data = FetchPOE2Data("Fate of the Vaal")

    If data Is Nothing Then
        Debug.Print "FAILED: Could not fetch data"
    Else
        Debug.Print "SUCCESS: Fetched " & data.Count & " items"

        Debug.Print "Testing number parsing..."
        Debug.Print "ParseNumberFromJSON(""1.0"") = " & ParseNumberFromJSON("1.0")
        Debug.Print "ParseNumberFromJSON(""26.73"") = " & ParseNumberFromJSON("26.73")
        Debug.Print "ParseNumberFromJSON(""3293.27"") = " & ParseNumberFromJSON("3293.27")

        Debug.Print "Testing price for Exalted Orb..."
        Dim testPrice As Variant
        testPrice = getPOE2Price("Exalted Orb")
        Debug.Print "Exalted Orb price: " & testPrice
    End If
End Sub
