'*******************************************************************************
' POE2 Scout Price Scraper for Excel - OPTIMIZED VERSION
' Version: 2.5.0
' Description: Fetch real-time POE2 item prices with auto-refresh support
' Author: POE2 Community
' License: MIT
' Data Source: https://poe2scout.com
'
' NEW IN v2.5:
' - Removed Application.Volatile to prevent constant recalculation
' - Excel now only recalculates when you manually update (Ctrl+Shift+A or F9)
' - Much better performance with many formulas
' - Manual refresh macros still available
'
' PREVIOUS (v2.4):
' - Updated default league to "Fate of the Vaal" (new league)
' - Verified API compatibility with new league structure
' - All functions now use current league by default
'*******************************************************************************

Option Explicit

' Module-level variables for caching
Private cachedData As Object
Private cacheTimestamp As Date
Private Const CACHE_DURATION_MINUTES As Integer = 5

' Auto-refresh timer
Private refreshTimer As Date

'*******************************************************************************
' CORE FUNCTIONS (NON-VOLATILE FOR BETTER PERFORMANCE)
'*******************************************************************************

Public Function getPOE2Price(ByVal itemName As String, _
                            Optional ByVal category As String = "", _
                            Optional ByVal league As String = "Fate of the Vaal") As Variant
    
    On Error GoTo ErrorHandler
    
    If Trim(itemName) = "" Then
        getPOE2Price = "Error: Item name required"
        Exit Function
    End If
    
    Dim data As Object
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
        getPOE2Price = ParseNumberFromJSON(GetDictValue(foundItem, "currentPrice", "0"))
        If Err.Number <> 0 Then
            getPOE2Price = "Error: Invalid price data"
        End If
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
    
    Dim data As Object
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
        result(1) = ParseNumberFromJSON(GetDictValue(foundItem, "currentPrice", "0"))
        result(2) = GetQuantity(foundItem)
        result(3) = GetDictValue(foundItem, "categoryApiId", "Unknown")
        getPOE2ItemDetails = result
    End If
    
    Exit Function
    
ErrorHandler:
    getPOE2ItemDetails = Array("Error", Err.Description, "", "")
End Function

Public Function getPOE2Items(Optional ByVal category As String = "", _
                            Optional ByVal league As String = "Fate of the Vaal") As Variant
    
    On Error GoTo ErrorHandler
    
    Dim data As Object
    Set data = FetchPOE2Data(league)
    
    If data Is Nothing Then
        getPOE2Items = Array(Array("Error", "Failed to fetch data", "", ""))
        Exit Function
    End If
    
    Dim items As Collection
    Set items = FilterByCategory(data, category)
    
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
        result(i, 1) = ParseNumberFromJSON(GetDictValue(item, "currentPrice", "0"))
        result(i, 2) = GetQuantity(item)
        result(i, 3) = GetDictValue(item, "categoryApiId", "Unknown")
    Next i
    
    getPOE2Items = result
    
    Exit Function
    
ErrorHandler:
    getPOE2Items = Array(Array("Error", Err.Description, "", ""))
End Function

Public Function getPOE2Categories(Optional ByVal league As String = "Fate of the Vaal") As Variant
    
    On Error GoTo ErrorHandler
    
    Dim data As Object
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
        cat = GetDictValue(item, "categoryApiId", "unknown")
        
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
' HELPER FUNCTIONS (PRIVATE)
'*******************************************************************************

Private Function ParseNumberFromJSON(ByVal value As Variant) As Double
    On Error GoTo ErrorHandler
    
    If IsNumeric(value) And VarType(value) <> vbString Then
        ParseNumberFromJSON = CDbl(value)
        Exit Function
    End If
    
    Dim strValue As String
    strValue = Trim(CStr(value))
    strValue = Replace(strValue, " ", "")
    
    Dim integerPart As String
    Dim decimalPart As String
    Dim dotPos As Long
    
    dotPos = InStr(strValue, ".")
    
    If dotPos > 0 Then
        integerPart = Left(strValue, dotPos - 1)
        decimalPart = Mid(strValue, dotPos + 1)
    Else
        integerPart = strValue
        decimalPart = "0"
    End If
    
    Dim intVal As Double
    Dim isNegative As Boolean
    
    If Left(integerPart, 1) = "-" Then
        isNegative = True
        integerPart = Mid(integerPart, 2)
    End If
    
    intVal = 0
    Dim i As Long
    For i = 1 To Len(integerPart)
        Dim digit As String
        digit = Mid(integerPart, i, 1)
        If digit >= "0" And digit <= "9" Then
            intVal = intVal * 10 + (Asc(digit) - Asc("0"))
        End If
    Next i
    
    Dim decVal As Double
    decVal = 0
    Dim divisor As Double
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
    
    If isNegative Then
        result = -result
    End If
    
    ParseNumberFromJSON = result
    Exit Function
    
ErrorHandler:
    Debug.Print "ParseNumberFromJSON Error: " & Err.Description & " for value: " & value
    ParseNumberFromJSON = 0
End Function

Private Function FetchPOE2Data(ByVal league As String) As Object
    On Error GoTo ErrorHandler
    
    If Not cachedData Is Nothing Then
        If DateDiff("n", cacheTimestamp, Now) < CACHE_DURATION_MINUTES Then
            Set FetchPOE2Data = cachedData
            Exit Function
        End If
    End If
    
    Dim url As String
    url = "https://poe2scout.com/api/items?league=" & Replace(league, " ", "+")
    
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "POE2-Excel-Scraper/2.4"
    http.setRequestHeader "Accept", "application/json"
    
    http.Send
    
    If http.Status <> 200 Then
        Debug.Print "HTTP Error: " & http.Status & " - " & http.statusText
        Set FetchPOE2Data = Nothing
        Exit Function
    End If
    
    Dim responseText As String
    responseText = http.responseText
    
    If Len(responseText) = 0 Then
        Debug.Print "Empty response from API"
        Set FetchPOE2Data = Nothing
        Exit Function
    End If
    
    Dim json As Object
    Set json = ParseJSONResponse(responseText)
    
    If json Is Nothing Then
        Debug.Print "Failed to parse JSON response"
        Set FetchPOE2Data = Nothing
        Exit Function
    End If
    
    Set cachedData = json
    cacheTimestamp = Now
    
    Set FetchPOE2Data = json
    Exit Function
    
ErrorHandler:
    Debug.Print "FetchPOE2Data Error: " & Err.Description & " (Code: " & Err.Number & ")"
    Set FetchPOE2Data = Nothing
End Function

Private Function ParseJSONResponse(ByVal jsonText As String) As Object
    On Error GoTo ErrorHandler
    
    Dim sc As Object
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    
    Dim jsArray As Object
    Set jsArray = sc.Eval("(" & jsonText & ")")
    
    Dim result As Collection
    Set result = New Collection
    
    Dim i As Long
    For i = 0 To jsArray.length - 1
        Dim item As Object
        Set item = ConvertJSObject(jsArray(i))
        result.Add item
    Next i
    
    Set ParseJSONResponse = result
    Exit Function
    
ErrorHandler:
    Debug.Print "ScriptControl failed, trying alternative parser"
    Set ParseJSONResponse = ParseJSONAlternative(jsonText)
End Function

Private Function ConvertJSObject(ByVal jsObj As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    dict.Add "itemId", jsObj.itemId
    dict.Add "apiId", jsObj.apiId
    dict.Add "text", jsObj.text
    dict.Add "categoryApiId", jsObj.categoryApiId
    dict.Add "currentPrice", CStr(jsObj.currentPrice)
    dict.Add "priceLogs", jsObj.priceLogs
    dict.Add "iconUrl", jsObj.iconUrl
    On Error GoTo 0
    
    Set ConvertJSObject = dict
End Function

Private Function ParseJSONAlternative(ByVal jsonText As String) As Object
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
        If Not item Is Nothing Then
            result.Add item
        End If
    Next i
    
    Set ParseJSONAlternative = result
End Function

Private Function SplitJSONObjects(ByVal jsonText As String) As String()
    Dim result() As String
    Dim objCount As Long
    Dim braceLevel As Long
    Dim currentObj As String
    Dim i As Long
    Dim char As String
    Dim inString As Boolean
    Dim prevChar As String
    
    objCount = 0
    braceLevel = 0
    currentObj = ""
    inString = False
    prevChar = ""
    
    For i = 1 To Len(jsonText)
        char = Mid(jsonText, i, 1)
        
        If char = """" And prevChar <> "\" Then
            inString = Not inString
        End If
        
        If Not inString Then
            If char = "{" Then
                braceLevel = braceLevel + 1
            ElseIf char = "}" Then
                braceLevel = braceLevel - 1
            End If
        End If
        
        currentObj = currentObj & char
        
        If braceLevel = 0 And Len(currentObj) > 2 Then
            ReDim Preserve result(objCount)
            result(objCount) = currentObj
            objCount = objCount + 1
            currentObj = ""
        End If
        
        prevChar = char
    Next i
    
    SplitJSONObjects = result
End Function

Private Function ParseJSONObject(ByVal jsonText As String) As Object
    On Error Resume Next
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.Add "text", ExtractJSONValue(jsonText, "text")
    dict.Add "currentPrice", ExtractJSONValue(jsonText, "currentPrice")
    dict.Add "categoryApiId", ExtractJSONValue(jsonText, "categoryApiId")
    dict.Add "itemId", ExtractJSONValue(jsonText, "itemId")
    
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
    
    Dim endPos As Long
    Dim char As String
    char = Mid(jsonText, startPos, 1)
    
    If char = """" Then
        startPos = startPos + 1
        endPos = InStr(startPos, jsonText, """")
        ExtractJSONValue = Mid(jsonText, startPos, endPos - startPos)
    Else
        endPos = startPos
        Do While endPos <= Len(jsonText)
            char = Mid(jsonText, endPos, 1)
            If char = "," Or char = "}" Or char = "]" Then Exit Do
            endPos = endPos + 1
        Loop
        ExtractJSONValue = Trim(Mid(jsonText, startPos, endPos - startPos))
    End If
End Function

Private Function FilterByCategory(ByVal data As Object, ByVal category As String) As Collection
    Dim filtered As Collection
    Set filtered = New Collection
    
    If Trim(category) = "" Then
        Dim i As Long
        For i = 1 To data.Count
            filtered.Add data(i)
        Next i
    Else
        Dim item As Object
        For i = 1 To data.Count
            Set item = data(i)
            If LCase(GetDictValue(item, "categoryApiId", "")) = LCase(Trim(category)) Then
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
        itemText = LCase(GetDictValue(item, "text", ""))
        
        If itemText = itemNameLower Then
            Set FindItem = item
            Exit Function
        End If
    Next i
    
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
            price1 = ParseNumberFromJSON(GetDictValue(arr(j), "currentPrice", "0"))
            price2 = ParseNumberFromJSON(GetDictValue(arr(j + 1), "currentPrice", "0"))
            
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
    
    Dim priceLogs As Object
    Set priceLogs = item("priceLogs")
    
    If Not priceLogs Is Nothing Then
        If priceLogs.length > 0 Then
            Dim latestLog As Object
            Set latestLog = priceLogs(0)
            If Not latestLog Is Nothing Then
                GetQuantity = latestLog("quantity")
                Exit Function
            End If
        End If
    End If
    
    GetQuantity = "N/A"
    On Error GoTo 0
End Function

Public Sub ClearCache()
    Set cachedData = Nothing
    cacheTimestamp = 0
    Debug.Print "Cache cleared at " & Now
End Sub

Public Sub TestAPIConnection()
    Debug.Print "Testing API connection..."
    Dim data As Object
    Set data = FetchPOE2Data("Fate of the Vaal")
    
    If data Is Nothing Then
        Debug.Print "FAILED: Could not fetch data"
    Else
        Debug.Print "SUCCESS: Fetched " & data.Count & " items"
        
        Debug.Print "Testing number parsing..."
        Debug.Print "ParseNumberFromJSON(""1.0"") = " & ParseNumberFromJSON("1.0")
        Debug.Print "ParseNumberFromJSON(""26.73"") = " & ParseNumberFromJSON("26.73")
        Debug.Print "ParseNumberFromJSON(""3293.27"") = " & ParseNumberFromJSON("3293.27")
        
        Debug.Print "Testing price conversion for Exalted Orb..."
        Dim testPrice As Variant
        testPrice = getPOE2Price("Exalted Orb")
        Debug.Print "Exalted Orb price: " & testPrice
    End If
End Sub
