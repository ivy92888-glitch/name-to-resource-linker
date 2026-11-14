Attribute VB_Name = "NameToResource"
' ============================================================
' name-to-resource-linker
' ------------------------------------------------------------
' Generic, domain-agnostic VBA macro that, for each free-text
' item name in a source sheet, finds the best-matching resource
' records from one or more configurable reference tables and
' writes up to N ranked reference links (with a clickable
' hyperlink) back into the sheet.
'
' This is an INDEPENDENT, clean-room reimplementation of a
' generic "name -> ranked reference resources" pattern. It
' contains no business-specific logic, no proprietary data,
' and no hard-coded domain codes.
'
' The engine performs seven reusable ideas:
'   1. Name-keyed resource registry: each reference name maps
'      to a list of resource records (attributes + a URL).
'   2. Two-tier name matching: exact match is preferred; a
'      boundary-aware fuzzy match (so "leg" does not match
'      "legacy") is used as a fallback.
'   3. Link filtering: records whose link contains any excluded
'      keyword are dropped.
'   4. Cross-field normalization: numeric rate values are
'      coerced to a percentage string; invalid metric values
'      are pushed to the end of the sort.
'   5. Metric-based ranking: each name's records are sorted by
'      a numeric metric (e.g. lowest unit price first).
'   6. Top-N priority blend: exact matches fill first, fuzzy
'      matches fill the rest, capped at MAX_LINKS.
'   7. Composite output + hyperlink: a readable attribute
'      string plus a clickable hyperlink is written per link.
'
' All sheet names, columns, keywords, weights and the output
' range are configurable constants at the top of this module.
' ============================================================

Option Explicit
Option Compare Text

' ---- Source sheet -----------------------------------------
Private Const SRC_SHEET        As String = ""          ' "" = active sheet
Private Const NAME_COL         As String = "C"         ' item name column
Private Const HEADER_ROW       As Long = 3             ' row holding the name header
Private Const DATA_START_ROW   As Long = 4             ' first data row
Private Const NAME_HEADER_TEXT As String = "Item name" ' header to expect; "" = skip

' ---- Reference tables (in the macro workbook) -------------
' Comma-separated sheet names; each holds the resource registry.
Private Const RESOURCE_SHEETS  As String = "Resources,ResourcesExtra"

' ---- Reference table column layout (10 columns A..J) ------
Private Const RES_NAME_COL     As String = "A"         ' key name (may use "/" aliases, "-suffix")
Private Const RES_DISPLAY_COL  As String = "B"         ' display / alias name
Private Const RES_CODE_COL     As String = "C"         ' optional code / identifier
Private Const RES_RATE_COL     As String = "D"         ' rate (numeric -> percentage)
Private Const RES_PRICE_COL    As String = "E"         ' price (generic)
Private Const RES_WEIGHT_COL   As String = "F"         ' weight / measure (generic)
Private Const RES_LINK_COL     As String = "G"         ' URL
Private Const RES_CATEGORY_COL As String = "H"         ' category (output only)
Private Const RES_GROUP_COL    As String = "I"         ' group (output only)
Private Const RES_METRIC_COL   As String = "J"         ' numeric sort metric

' ---- Output -----------------------------------------------
Private Const OUTPUT_LINK_COL_START As String = "AG"    ' first output column
Private Const MAX_LINKS       As Long = 3             ' max links per item
Private Const HYPERLINK_ON    As Boolean = True       ' attach clickable links
Private Const FORMAT_ON       As Boolean = True       ' style output columns

' ---- Exclusion / matching rules ---------------------------
' Links containing any of these (case-insensitive) are skipped.
Private Const EXCLUDE_LINK_KEYWORDS As String = "out of stock,discontinued,sold out"
' Database names ending in these generic words (and the user input
' does not) are NOT matched, to avoid over-broad hits.
Private Const EXCLUDE_SUFFIX_WORDS As String = "accessory,kit,cover,base,cap,decor,part"

' ---- Misc -------------------------------------------------
Private Const COMPOSITE_SEP   As String = "="         ' field separator in output string
Private Const NO_LINK_VALUE   As String = ""          ' cleared cell value

' ============================================================
' Entry point
' ============================================================
Public Sub LinkResources()
    Dim t0 As Double
    t0 = Timer

    Dim srcWB As Workbook, srcWs As Worksheet
    Dim indexes As Object
    Dim lastRow As Long, processed As Long, found As Long

    On Error GoTo ErrHandler
    Application.StatusBar = "Initializing resource linker..."
    DoEvents

    Set srcWB = ActiveWorkbook
    If SRC_SHEET = "" Then
        Set srcWs = srcWB.ActiveSheet
    Else
        On Error Resume Next
        Set srcWs = srcWB.Sheets(SRC_SHEET)
        On Error GoTo 0
        If srcWs Is Nothing Then
            MsgBox "Source sheet '" & SRC_SHEET & "' not found.", vbCritical, "Config error"
            Exit Sub
        End If
    End If

    If NAME_HEADER_TEXT <> "" Then
        If Trim(srcWs.Cells(HEADER_ROW, NAME_COL).Value) = "" Then
            MsgBox "Header cell " & NAME_COL & HEADER_ROW & " is empty." & vbCrLf & _
                   "Expected header: " & NAME_HEADER_TEXT, vbExclamation, "Format check"
            Exit Sub
        End If
    End If

    lastRow = srcWs.Cells(srcWs.Rows.Count, NAME_COL).End(xlUp).Row
    If lastRow < DATA_START_ROW Then
        MsgBox "No data found in column " & NAME_COL & " from row " & DATA_START_ROW & ".", _
               vbInformation, "Data check"
        Exit Sub
    End If

    ' Build resource indexes from the configured reference sheets
    Set indexes = CreateObject("Scripting.Dictionary")
    If Not BuildAllIndexes(indexes) Then Exit Sub

    If FORMAT_ON Then Call SetLinkHeaders(srcWs)

    Application.StatusBar = "Matching..."
    DoEvents

    Dim i As Long
    For i = DATA_START_ROW To lastRow
        Dim nm As String
        nm = Trim(srcWs.Cells(i, NAME_COL).Value)
        If nm = "" Then
            If FORMAT_ON Then ClearLinks srcWs, i
            GoTo ContinueLoop
        End If

        Dim pct As Double
        pct = (i - DATA_START_ROW + 1) / (lastRow - DATA_START_ROW + 1) * 100
        Application.StatusBar = "Row " & i & "/" & lastRow & " (" & Format(pct, "0.0") & "%)"
        DoEvents

        Call FindLinksForProduct(srcWs, i, nm, indexes, found)
        processed = processed + 1
ContinueLoop:
    Next i

    If FORMAT_ON Then Call SetLinkFormat(srcWs, DATA_START_ROW, lastRow)
    Application.StatusBar = False
    Call ShowDone(lastRow - DATA_START_ROW + 1, processed, found, Timer - t0)
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Resource linking failed"
End Sub

' ============================================================
' Build the name -> resource-records index from all sheets
' ============================================================
Private Function BuildAllIndexes(ByRef indexes As Object) As Boolean
    Dim sheets As Variant
    sheets = Split(RESOURCE_SHEETS, ",")
    Dim s As Variant
    For Each s In sheets
        Dim shName As String
        shName = Trim(s)
        If shName = "" Then GoTo NextSheet
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(shName)
        On Error GoTo 0
        If ws Is Nothing Then
            MsgBox "Resource sheet '" & shName & "' not found in this workbook.", vbCritical, "Database error"
            BuildAllIndexes = False
            Exit Function
        End If
        Call BuildIndexForSheet(ws, indexes)
NextSheet:
    Next s
    If indexes.Count = 0 Then
        MsgBox "No resource records loaded from: " & RESOURCE_SHEETS, vbExclamation, "Database error"
        BuildAllIndexes = False
        Exit Function
    End If
    BuildAllIndexes = True
End Function

Private Sub BuildIndexForSheet(ByVal ws As Worksheet, ByRef indexes As Object)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, RES_NAME_COL).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        Dim rName As String
        rName = Trim(ws.Cells(i, RES_NAME_COL).Value)
        If rName = "" Then GoTo NextRow

        Dim rec As Object
        Set rec = CreateObject("Scripting.Dictionary")
        rec("DisplayName") = Trim(ws.Cells(i, RES_DISPLAY_COL).Value)
        rec("Code") = Trim(ws.Cells(i, RES_CODE_COL).Value)
        rec("Rate") = FormatRate(ws.Cells(i, RES_RATE_COL).Value)
        rec("Price") = ToDouble(ws.Cells(i, RES_PRICE_COL).Value)
        rec("Weight") = ToDouble(ws.Cells(i, RES_WEIGHT_COL).Value)
        rec("Link") = Trim(ws.Cells(i, RES_LINK_COL).Value)
        rec("Category") = Trim(ws.Cells(i, RES_CATEGORY_COL).Value)
        rec("Group") = Trim(ws.Cells(i, RES_GROUP_COL).Value)
        rec("Metric") = ToDouble(ws.Cells(i, RES_METRIC_COL).Value)

        ' Link filter: drop records whose link is excluded
        If LinkExcluded(rec("Link")) Then GoTo NextRow

        ' Handle "/" aliases in the resource name
        If InStr(rName, "/") > 0 Then
            Dim parts() As String
            parts = Split(rName, "/")
            Dim j As Long
            For j = LBound(parts) To UBound(parts)
                Dim singleName As String
                singleName = Trim(parts(j))
                If singleName <> "" Then Call AddResource(indexes, CleanName(singleName), rec)
            Next j
        Else
            Call AddResource(indexes, CleanName(rName), rec)
        End If
NextRow:
    Next i
    ' Sort each name's list by metric (ascending; invalid -> end)
    Call SortIndexesByMetric(indexes)
End Sub

Private Sub AddResource(ByRef indexes As Object, ByVal name As String, ByVal rec As Object)
    If Not indexes.Exists(name) Then
        Dim list As Variant
        list = Array()
        indexes.Add name, list
    End If
    Dim cur As Variant
    cur = indexes(name)
    ReDim Preserve cur(UBound(cur) + 1)
    Set cur(UBound(cur)) = rec
    indexes(name) = cur
End Sub

' Coerce a numeric rate to a percentage string; pass through text.
Private Function FormatRate(ByVal v As Variant) As String
    If IsNumeric(v) Then
        Dim d As Double
        d = CDbl(v)
        If d > 0 And d < 1 Then
            FormatRate = Format(d * 100, "0.00") & "%"
        ElseIf d >= 1 And d < 100 Then
            FormatRate = Format(d, "0.00") & "%"
        Else
            FormatRate = CStr(v)
        End If
    Else
        FormatRate = Trim(CStr(v))
    End If
End Function

Private Function ToDouble(ByVal v As Variant) As Double
    On Error Resume Next
    ToDouble = CDbl(v)
    If Err.Number <> 0 Then ToDouble = 0
    On Error GoTo 0
End Function

Private Function LinkExcluded(ByVal link As String) As Boolean
    Dim keys() As String
    keys = Split(EXCLUDE_LINK_KEYWORDS, ",")
    Dim k As Variant
    For Each k In keys
        If Trim(k) <> "" Then
            If InStr(1, link, Trim(k), vbTextCompare) > 0 Then
                LinkExcluded = True
                Exit Function
            End If
        End If
    Next k
    LinkExcluded = False
End Function

' Strip a trailing "-suffix" (e.g. "-black") from a resource name.
Private Function CleanName(ByVal nm As String) As String
    Dim s As String
    s = Trim(nm)
    Dim dash As Long
    dash = InStr(s, "-")
    If dash > 0 Then s = Trim(Left(s, dash - 1))
    CleanName = s
End Function

' Bubble-sort each name's record array by metric ascending.
Private Sub SortIndexesByMetric(ByRef indexes As Object)
    Dim key As Variant
    For Each key In indexes.Keys
        Dim arr As Variant
        arr = indexes(key)
        If UBound(arr) > 0 Then
            Dim sorted As Boolean
            sorted = False
            Do While Not sorted
                sorted = True
                Dim i As Long
                For i = LBound(arr) To UBound(arr) - 1
                    Dim m1 As Double, m2 As Double
                    m1 = MetricValue(arr(i))
                    m2 = MetricValue(arr(i + 1))
                    If m1 > m2 Then
                        Dim tmp As Object
                        Set tmp = arr(i)
                        Set arr(i) = arr(i + 1)
                        Set arr(i + 1) = tmp
                        sorted = False
                    End If
                Next i
            Loop
            indexes(key) = arr
        End If
    Next key
End Sub

Private Function MetricValue(ByVal rec As Object) As Double
    Dim m As Double
    m = rec("Metric")
    If m <= 0 Then m = 1E+308
    MetricValue = m
End Function

' ============================================================
' Matching primitives (two-tier)
' ============================================================
Private Function IsExactMatch(ByVal dbName As String, ByVal itemName As String) As Boolean
    Dim cleaned As String
    cleaned = CleanName(dbName)
    If InStr(cleaned, "/") > 0 Then
        Dim parts() As String
        parts = Split(cleaned, "/")
        Dim p As Variant
        For Each p In parts
            If StrComp(Trim(p), itemName, vbTextCompare) = 0 Then
                IsExactMatch = True
                Exit Function
            End If
        Next p
    Else
        IsExactMatch = (StrComp(cleaned, itemName, vbTextCompare) = 0)
    End If
End Function

' Database name must CONTAIN the item name before the boundary
' validation runs (so a more specific db entry matches a generic
' source term, e.g. "X-shaped desk leg" matches "desk leg").
Private Function IsIntelligentMatch(ByVal dbName As String, ByVal itemName As String) As Boolean
    Dim cleaned As String
    cleaned = CleanName(dbName)
    If InStr(1, cleaned, itemName, vbTextCompare) = 0 Then Exit Function
    If InStr(cleaned, "/") > 0 Then
        Dim parts() As String
        parts = Split(cleaned, "/")
        Dim p As Variant
        For Each p In parts
            If Trim(p) <> "" Then
                If IsValidMatch(Trim(p), itemName) Then
                    IsIntelligentMatch = True
                    Exit Function
                End If
            End If
        Next p
    Else
        IsIntelligentMatch = IsValidMatch(cleaned, itemName)
    End If
End Function

' Position-aware validation that prevents partial over-matches:
'   - item at the END of db name        -> OK   ("x-shaped desk leg")
'   - item at START + "-suffix"         -> OK   ("desk leg-black")
'   - item at START + extra word        -> NO   ("desk leg kit")
'   - item in the MIDDLE                -> NO   ("black desk leg cover")
' Plus a generic-suffix blacklist.
Private Function IsValidMatch(ByVal dbName As String, ByVal itemName As String) As Boolean
    Dim bl() As String
    bl = Split(EXCLUDE_SUFFIX_WORDS, ",")
    Dim w As Variant
    For Each w In bl
        If Trim(w) <> "" Then
            If InStr(1, dbName, Trim(w), vbTextCompare) > 0 Then
                If InStr(1, itemName, Trim(w), vbTextCompare) = 0 Then
                    IsValidMatch = False
                    Exit Function
                End If
            End If
        End If
    Next w

    Dim result As Boolean
    result = True
    Dim dbL As String, itL As String
    dbL = LCase(dbName)
    itL = LCase(itemName)

    If Right(dbL, Len(itL)) = itL Then
        result = True
    ElseIf Left(dbL, Len(itL)) = itL Then
        Dim remPart As String
        remPart = Mid(dbL, Len(itL) + 1)
        If Left(remPart, 1) = "-" Then
            result = True
        Else
            result = False
        End If
    Else
        result = False
    End If
    IsValidMatch = result
End Function

' ============================================================
' Per-product link resolution
' ============================================================
Private Sub FindLinksForProduct(ByVal ws As Worksheet, ByVal row As Long, _
                               ByVal itemName As String, ByVal indexes As Object, _
                               ByRef foundCount As Long)
    Dim exactLinks As Collection, fuzzyLinks As Collection
    Set exactLinks = New Collection
    Set fuzzyLinks = New Collection

    Dim key As Variant
    For Each key In indexes.Keys
        Dim isExact As Boolean
        isExact = IsExactMatch(key, itemName)
        If isExact Or IsIntelligentMatch(key, itemName) Then
            Dim arr As Variant
            arr = indexes(key)
            Dim i As Long
            For i = LBound(arr) To UBound(arr)
                Dim rec As Object
                Set rec = arr(i)
                If rec("Link") <> "" Then
                    If isExact Then
                        exactLinks.Add rec
                    Else
                        fuzzyLinks.Add rec
                    End If
                End If
            Next i
        End If
    Next key

    ' Prefer exact matches, then fill with fuzzy, capped at MAX_LINKS
    Dim finalList As Collection
    Set finalList = New Collection
    Dim c As Variant
    For Each c In exactLinks
        If finalList.Count < MAX_LINKS Then finalList.Add c Else Exit For
    Next c
    If finalList.Count < MAX_LINKS Then
        For Each c In fuzzyLinks
            If finalList.Count < MAX_LINKS Then finalList.Add c Else Exit For
        Next c
    End If

    If finalList.Count > 0 Then
        Dim colIdx As Long
        colIdx = ColumnNumber(OUTPUT_LINK_COL_START)
        For Each c In finalList
            Set rec = c
            Dim combined As String
            combined = BuildComposite(rec)
            ws.Cells(row, colIdx).Value = combined
            If HYPERLINK_ON Then
                Dim link As String
                link = rec("Link")
                If Left(link, 4) = "http" Then
                    ws.Hyperlinks.Add Anchor:=ws.Cells(row, colIdx), _
                                      Address:=link, TextToDisplay:=combined
                End If
            End If
            colIdx = colIdx + 1
        Next c
        foundCount = foundCount + 1
    Else
        If FORMAT_ON Then ClearLinks ws, row
    End If
End Sub

Private Function BuildComposite(ByVal rec As Object) As String
    Dim disp As String, code As String, rate As String, weight As String
    disp = rec("DisplayName")
    If disp = "" Then disp = "N/A"
    code = rec("Code")
    If code = "" Then code = "N/A"
    rate = rec("Rate")
    If rate = "" Then rate = "N/A"
    Dim w As Double
    w = rec("Weight")
    If w > 0 Then weight = Format(w, "0.00") Else weight = "N/A"
    BuildComposite = disp & COMPOSITE_SEP & code & COMPOSITE_SEP & rate & _
                     COMPOSITE_SEP & weight & COMPOSITE_SEP & rec("Link")
End Function

' ============================================================
' Output formatting
' ============================================================
Private Sub SetLinkHeaders(ByVal ws As Worksheet)
    Dim startCol As Long
    startCol = ColumnNumber(OUTPUT_LINK_COL_START)
    Dim i As Long
    For i = 0 To MAX_LINKS - 1
        With ws.Cells(HEADER_ROW, startCol + i)
            .Value = "Link " & (i + 1)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End With
    Next i
End Sub

Private Sub SetLinkFormat(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    Dim startCol As Long
    startCol = ColumnNumber(OUTPUT_LINK_COL_START)
    ws.Columns(startCol).Resize(, MAX_LINKS).ColumnWidth = 25.8
    Dim i As Long, c As Long
    For i = startRow To endRow
        For c = 0 To MAX_LINKS - 1
            With ws.Cells(i, startCol + c)
                .Font.Name = "Calibri"
                .Font.Size = 9
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlTop
                .WrapText = True
                With .Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
            End With
        Next c
        ws.Rows(i).RowHeight = 26.75
    Next i
End Sub

Private Sub ClearLinks(ByVal ws As Worksheet, ByVal row As Long)
    Dim startCol As Long
    startCol = ColumnNumber(OUTPUT_LINK_COL_START)
    Dim i As Long
    For i = 0 To MAX_LINKS - 1
        ws.Cells(row, startCol + i).Value = NO_LINK_VALUE
    Next i
End Sub

' ============================================================
' Helpers
' ============================================================
Private Function ColumnNumber(ByVal col As String) As Long
    Dim s As String
    s = UCase(col)
    Dim n As Long
    n = 0
    Dim i As Long
    For i = 1 To Len(s)
        n = n * 26 + (Asc(Mid(s, i, 1)) - 64)
    Next i
    ColumnNumber = n
End Function

Private Function LastColumnLetter() As String
    Dim n As Long
    n = ColumnNumber(OUTPUT_LINK_COL_START) + MAX_LINKS - 1
    Dim s As String
    s = ""
    Do While n > 0
        Dim m As Long
        m = (n - 1) Mod 26
        s = Chr(65 + m) & s
        n = (n - m - 1) \ 26
    Loop
    LastColumnLetter = s
End Function

' ============================================================
' Completion message
' ============================================================
Private Sub ShowDone(ByVal total As Long, ByVal processed As Long, _
                     ByVal found As Long, ByVal dur As Double)
    Dim s As String
    s = "Resource linking complete!" & vbCrLf & _
        "============================" & vbCrLf & _
        "Total rows:   " & total & vbCrLf & _
        "Processed:    " & processed & vbCrLf & _
        "Found links:  " & found & vbCrLf & _
        "============================" & vbCrLf & _
        "Elapsed:      " & Format(dur, "0.00") & " s" & vbCrLf & _
        "============================" & vbCrLf & _
        "Output: columns " & OUTPUT_LINK_COL_START & ".." & LastColumnLetter() & vbCrLf & _
        "Each cell: DisplayName=Code=Rate=Weight=Link"
    MsgBox s, vbInformation, "Done"
End Sub
