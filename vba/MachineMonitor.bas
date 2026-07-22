Attribute VB_Name = "MachineMonitor"
'==========================================================================
'  MachineMonitor.bas  —  VBA backend for MachineMonitor.xlsm
'
'  One-time install:
'     1. Open MachineMonitor.xlsx in Excel
'     2. Alt+F11 → File → Import File… → pick this .bas file
'     3. Close VBE, Save As Macro-Enabled Workbook (*.xlsm)
'     4. Alt+F8 → run "SetupWorkbook" once
'
'  Public macros:
'     SetupWorkbook       — installs shape buttons, hotkeys, event hooks
'     ImportCSV           — batch-load a CSV into the Data sheet
'     RefreshDashboard    — recompute the comparison charts
'     GenerateReport      — new sheet snapshot + PDF export
'     ClearData           — wipe the Data table (with confirm)
'     ToggleLive          — enable/disable 60-second auto-refresh
'==========================================================================
Option Explicit

Private Const SHT_DASH As String    = "Dashboard"
Private Const SHT_DATA As String    = "Data"
Private Const SHT_CONF As String    = "Config"
Private Const SHT_REPORTS As String = "Reports"
Private Const TBL_DATA As String    = "tblData"
Private Const TBL_MACH As String    = "tblMachines"

Private Const LIVE_INTERVAL_SEC As Long = 60

Public gLiveOn As Boolean
Public gNextTick As Double

'--------------------------------------------------------------------------
'  ONE-TIME SETUP
'--------------------------------------------------------------------------
Public Sub SetupWorkbook()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_DASH)

    '--- Remove any old shape buttons from a previous run
    Dim shp As Shape, i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left$(shp.Name, 3) = "btn" Then shp.Delete
    Next i

    '--- Rebuild the four action buttons over the styled cells
    AddButton ws, "btnImport",   ws.Range("J7:L7"),   "IMPORT  CSV",       "ImportCSV"
    AddButton ws, "btnRefresh",  ws.Range("J8:L8"),   "REFRESH  DASHBOARD","RefreshDashboard"
    AddButton ws, "btnReport",   ws.Range("N7:P7"),   "GENERATE  REPORT",  "GenerateReport"
    AddButton ws, "btnClear",    ws.Range("N8:P8"),   "CLEAR  DATA",       "ClearData"

    '--- Global hotkeys
    Application.OnKey "^+i", "ImportCSV"
    Application.OnKey "^+r", "RefreshDashboard"
    Application.OnKey "^+g", "GenerateReport"

    '--- Initial refresh so the charts reflect current filters
    RefreshDashboard

    MsgBox "Setup complete." & vbCrLf & _
           "Shortcuts: Ctrl+Shift+I, Ctrl+Shift+R, Ctrl+Shift+G.", _
           vbInformation, "Machine Monitor"
End Sub

Private Sub AddButton(ws As Worksheet, nm As String, rng As Range, _
                      caption As String, macroName As String)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                               rng.Left, rng.Top, rng.Width, rng.Height)
    s.Name = nm
    With s
        .Fill.ForeColor.RGB = RGB(232, 163, 61)     '  Amber
        .Line.ForeColor.RGB = RGB(200, 122, 31)
        .Line.Weight = 1.25
        .Shadow.Type = msoShadow21
        .OnAction = macroName
        With .TextFrame2
            .TextRange.Text = caption
            .TextRange.Font.Name = "Consolas"
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Size = 11
            .TextRange.Font.Fill.ForeColor.RGB = RGB(15, 26, 32)
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

'--------------------------------------------------------------------------
'  IMPORT CSV
'--------------------------------------------------------------------------
Public Sub ImportCSV()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    fd.Title = "Select machine-data CSV (columns: machine_id, tag, timestamp, metric1, metric2, metric3)"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "CSV files", "*.csv"
    If fd.Show <> -1 Then Exit Sub

    Dim path As String: path = fd.SelectedItems(1)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHT_DATA)
    Dim lo As ListObject: Set lo = ws.ListObjects(TBL_DATA)

    Dim ff As Integer: ff = FreeFile
    Open path For Input As #ff

    Dim line As String, parts() As String, isHeader As Boolean
    isHeader = True

    Dim added As Long: added = 0
    Dim newRow As ListRow

    Do While Not EOF(ff)
        Line Input #ff, line
        If Len(Trim$(line)) = 0 Then GoTo Continue

        parts = SplitCSV(line)
        If UBound(parts) < 5 Then GoTo Continue

        If isHeader Then
            ' Skip a header row if the 3rd column can't be parsed as a date
            isHeader = False
            If Not IsDate(parts(2)) Then GoTo Continue
        End If

        Set newRow = lo.ListRows.Add
        With newRow.Range
            .Cells(1, 1).Value = Trim$(parts(0))
            .Cells(1, 2).Value = Trim$(parts(1))
            If IsDate(parts(2)) Then
                .Cells(1, 3).Value = CDate(parts(2))
                .Cells(1, 3).NumberFormat = "yyyy-mm-dd hh:mm"
            Else
                .Cells(1, 3).Value = parts(2)
            End If
            .Cells(1, 4).Value = CDblSafe(parts(3))
            .Cells(1, 5).Value = CDblSafe(parts(4))
            .Cells(1, 6).Value = CDblSafe(parts(5))
        End With
        added = added + 1
Continue:
    Loop
    Close #ff

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    RefreshDashboard

    MsgBox added & " row(s) imported from " & vbCrLf & path, _
           vbInformation, "Import complete"
End Sub

Private Function SplitCSV(line As String) As String()
    ' Minimal CSV splitter that respects double quotes.
    Dim out() As String, buf As String, inQ As Boolean, i As Long, ch As String
    ReDim out(0 To 5)
    Dim col As Long: col = 0
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If ch = """" Then
            inQ = Not inQ
        ElseIf ch = "," And Not inQ Then
            If col <= UBound(out) Then out(col) = buf
            col = col + 1
            buf = ""
        Else
            buf = buf & ch
        End If
    Next i
    If col <= UBound(out) Then out(col) = buf
    SplitCSV = out
End Function

Private Function CDblSafe(s As String) As Double
    On Error Resume Next
    CDblSafe = CDbl(Trim$(s))
End Function

'--------------------------------------------------------------------------
'  REFRESH  DASHBOARD
'--------------------------------------------------------------------------
Public Sub RefreshDashboard()
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets(SHT_DASH)
    Dim wsX As Worksheet: Set wsX = ThisWorkbook.Worksheets(SHT_DATA)
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Worksheets(SHT_CONF)

    Dim dStart As Date, dEnd As Date, tag As String
    dStart = wsD.Range("StartDate").Value
    dEnd = wsD.Range("EndDate").Value
    tag = Trim$(CStr(wsD.Range("TagFilter").Value))

    ' --- 1. Discover machines belonging to the selected tag (from Config)
    Dim machines() As String
    ReDim machines(0 To 15)
    Dim mCount As Long: mCount = 0
    Dim r As Long
    For r = 5 To 20
        If UCase$(Trim$(CStr(wsC.Cells(r, 8).Value))) = "YES" _
           And Trim$(CStr(wsC.Cells(r, 4).Value)) = tag Then
            machines(mCount) = Trim$(CStr(wsC.Cells(r, 2).Value))
            mCount = mCount + 1
        End If
    Next r

    ' --- 2. Wipe staging area (R:AF, rows 4-500)
    Application.ScreenUpdating = False
    wsD.Range("R4:AF500").ClearContents

    ' --- 3. Header row on staging
    wsD.Cells(4, 18).Value = "Date"
    Dim i As Long
    For i = 0 To mCount - 1
        wsD.Cells(4, 19 + i).Value = machines(i)
    Next i

    If mCount = 0 Then GoTo Finalize

    ' --- 4. Build sorted unique dates + averages per machine per metric-focus
    Dim focus As String: focus = CStr(wsD.Range("MetricFocus").Value)
    Dim colOffset As Long
    ' colOffset = column index inside the DataBodyRange array
    ' arr(i,1)=machine_id  arr(i,2)=tag  arr(i,3)=timestamp
    ' arr(i,4)=Metric1     arr(i,5)=Metric2  arr(i,6)=Metric3
    Select Case focus
        Case "Metric 1", "All": colOffset = 4
        Case "Metric 2":         colOffset = 5
        Case "Metric 3":         colOffset = 6
        Case Else:               colOffset = 4
    End Select

    ' Collect date buckets (yyyy-mm-dd) -> row index
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim lo As ListObject: Set lo = wsX.ListObjects(TBL_DATA)
    Dim rng As Range: Set rng = lo.DataBodyRange
    If rng Is Nothing Then GoTo Finalize

    Dim arr As Variant: arr = rng.Value
    Dim n As Long: n = UBound(arr, 1)

    ' Sum & count matrices  [dateIdx, machineIdx]
    Dim sums() As Double, cnts() As Long
    ReDim sums(1 To 400, 0 To mCount - 1)
    ReDim cnts(1 To 400, 0 To mCount - 1)
    Dim dateList() As Date
    ReDim dateList(1 To 400)
    Dim dCount As Long: dCount = 0

    Dim mIdx As Long, dIdx As Long, ts As Date, mid As String, tg As String, key As String
    For i = 1 To n
        mid = CStr(arr(i, 1))
        tg  = CStr(arr(i, 2))
        If tg = tag Then
            If IsDate(arr(i, 3)) Then
                ts = CDate(arr(i, 3))
                If ts >= dStart And ts < dEnd + 1 Then
                    mIdx = MachineIndex(machines, mCount, mid)
                    If mIdx >= 0 Then
                        key = Format$(ts, "yyyy-mm-dd")
                        If dict.Exists(key) Then
                            dIdx = dict(key)
                        Else
                            dCount = dCount + 1
                            dict.Add key, dCount
                            dateList(dCount) = DateSerial(Year(ts), Month(ts), Day(ts))
                            dIdx = dCount
                        End If
                        Dim v As Double
                        v = CDblSafe(CStr(arr(i, colOffset)))
                        sums(dIdx, mIdx) = sums(dIdx, mIdx) + v
                        cnts(dIdx, mIdx) = cnts(dIdx, mIdx) + 1
                    End If
                End If
            End If
        End If
    Next i

    ' --- 5. Sort dates + write out
    If dCount = 0 Then GoTo Finalize
    QuickSortDates dateList, 1, dCount

    Dim rowOut As Long
    For rowOut = 1 To dCount
        wsD.Cells(4 + rowOut, 18).Value = dateList(rowOut)
        wsD.Cells(4 + rowOut, 18).NumberFormat = "yyyy-mm-dd"
        key = Format$(dateList(rowOut), "yyyy-mm-dd")
        dIdx = dict(key)
        For mIdx = 0 To mCount - 1
            If cnts(dIdx, mIdx) > 0 Then
                wsD.Cells(4 + rowOut, 19 + mIdx).Value = sums(dIdx, mIdx) / cnts(dIdx, mIdx)
            End If
        Next mIdx
    Next rowOut

    ' --- 6. Rebind charts to the new range
    RebindCharts wsD, mCount, dCount

Finalize:
    Application.ScreenUpdating = True
End Sub

Private Function MachineIndex(machines() As String, count As Long, mid As String) As Long
    Dim i As Long
    For i = 0 To count - 1
        If machines(i) = mid Then MachineIndex = i: Exit Function
    Next i
    MachineIndex = -1
End Function

Private Sub QuickSortDates(a() As Date, lo As Long, hi As Long)
    Dim i As Long, j As Long, piv As Date, tmp As Date
    i = lo: j = hi
    piv = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < piv: i = i + 1: Loop
        Do While a(j) > piv: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortDates a, lo, j
    If i < hi Then QuickSortDates a, i, hi
End Sub

Private Sub RebindCharts(ws As Worksheet, mCount As Long, dCount As Long)
    Dim co As ChartObject, ch As Chart
    Dim lastCol As Long: lastCol = 18 + mCount   ' col S = 19 for machine 1
    Dim lastRow As Long: lastRow = 4 + dCount

    For Each co In ws.ChartObjects
        Set ch = co.Chart
        ' Clear existing series
        Do While ch.SeriesCollection.count > 0
            ch.SeriesCollection(1).Delete
        Loop
        If mCount = 0 Or dCount = 0 Then GoTo NextChart

        Dim m As Long
        For m = 1 To mCount
            With ch.SeriesCollection.NewSeries
                .Name = "='" & ws.Name & "'!" & ws.Cells(4, 18 + m).Address
                .XValues = ws.Range(ws.Cells(5, 18), ws.Cells(lastRow, 18))
                .Values  = ws.Range(ws.Cells(5, 18 + m), ws.Cells(lastRow, 18 + m))
                .Format.Line.Weight = 1.75
                .Smooth = True
                Select Case (m - 1) Mod 4
                    Case 0: .Format.Line.ForeColor.RGB = RGB(63, 167, 184)
                    Case 1: .Format.Line.ForeColor.RGB = RGB(232, 163, 61)
                    Case 2: .Format.Line.ForeColor.RGB = RGB(155, 197, 61)
                    Case 3: .Format.Line.ForeColor.RGB = RGB(214, 69, 69)
                End Select
            End With
        Next m
NextChart:
    Next co
End Sub

'--------------------------------------------------------------------------
'  GENERATE  REPORT  (new sheet + PDF)
'--------------------------------------------------------------------------
Public Sub GenerateReport()
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets(SHT_DASH)
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHT_REPORTS)
    Dim wsX As Worksheet: Set wsX = ThisWorkbook.Worksheets(SHT_DATA)

    Dim dStart As Date: dStart = wsD.Range("StartDate").Value
    Dim dEnd As Date:   dEnd = wsD.Range("EndDate").Value
    Dim tag As String:  tag = CStr(wsD.Range("TagFilter").Value)

    Dim stamp As String: stamp = Format$(Now, "yyyymmdd_hhnnss")
    Dim shtName As String: shtName = "Rpt_" & tag & "_" & stamp
    If Len(shtName) > 31 Then shtName = Left$(shtName, 31)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Duplicate Dashboard as a static snapshot
    wsD.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
    Dim snap As Worksheet: Set snap = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
    snap.Name = shtName
    snap.Tab.Color = RGB(155, 197, 61)

    ' Convert formulas to values (freeze the snapshot)
    snap.UsedRange.Value = snap.UsedRange.Value

    ' Overwrite footer with report meta
    snap.Range("B48").Value = "  Report generated " & Format$(Now, "yyyy-mm-dd hh:mm") & _
                              "  |  Tag: " & tag & _
                              "  |  Range: " & Format$(dStart, "yyyy-mm-dd") & " → " & Format$(dEnd, "yyyy-mm-dd")

    ' PDF export next to the workbook
    Dim pdfPath As String
    pdfPath = ThisWorkbook.path & Application.PathSeparator & shtName & ".pdf"
    On Error Resume Next
    snap.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0

    ' Log to Reports sheet
    Dim r As Long: r = wsR.Cells(wsR.Rows.count, 2).End(xlUp).Row + 1
    If r < 5 Then r = 5
    wsR.Cells(r, 2).Value = Now
    wsR.Cells(r, 2).NumberFormat = "yyyy-mm-dd hh:mm"
    wsR.Cells(r, 3).Value = tag
    wsR.Cells(r, 4).Value = dStart:  wsR.Cells(r, 4).NumberFormat = "yyyy-mm-dd"
    wsR.Cells(r, 5).Value = dEnd:    wsR.Cells(r, 5).NumberFormat = "yyyy-mm-dd"
    wsR.Cells(r, 6).Value = CountTaggedMachines(tag)
    wsR.Cells(r, 7).Value = CountPointsInRange(tag, dStart, dEnd)
    wsR.Cells(r, 8).Value = shtName
    wsR.Cells(r, 9).Value = pdfPath

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Report created:" & vbCrLf & _
           "  • Sheet: " & shtName & vbCrLf & _
           "  • PDF:   " & pdfPath, _
           vbInformation, "Report ready"
End Sub

Private Function CountTaggedMachines(tag As String) As Long
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Worksheets(SHT_CONF)
    Dim r As Long, c As Long: c = 0
    For r = 5 To 20
        If Trim$(CStr(wsC.Cells(r, 4).Value)) = tag _
           And UCase$(Trim$(CStr(wsC.Cells(r, 8).Value))) = "YES" Then
            c = c + 1
        End If
    Next r
    CountTaggedMachines = c
End Function

Private Function CountPointsInRange(tag As String, d1 As Date, d2 As Date) As Long
    Dim wsX As Worksheet: Set wsX = ThisWorkbook.Worksheets(SHT_DATA)
    Dim lo As ListObject: Set lo = wsX.ListObjects(TBL_DATA)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim arr As Variant: arr = lo.DataBodyRange.Value
    Dim i As Long, cnt As Long
    For i = 1 To UBound(arr, 1)
        If CStr(arr(i, 2)) = tag Then
            If IsDate(arr(i, 3)) Then
                If CDate(arr(i, 3)) >= d1 And CDate(arr(i, 3)) < d2 + 1 Then cnt = cnt + 1
            End If
        End If
    Next i
    CountPointsInRange = cnt
End Function

'--------------------------------------------------------------------------
'  CLEAR  DATA
'--------------------------------------------------------------------------
Public Sub ClearData()
    If MsgBox("Delete every row in the Data table? This cannot be undone.", _
              vbYesNo + vbExclamation, "Clear data") <> vbYes Then Exit Sub

    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(SHT_DATA).ListObjects(TBL_DATA)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    RefreshDashboard
End Sub

'--------------------------------------------------------------------------
'  LIVE  UPDATES  (future feed hook)
'--------------------------------------------------------------------------
Public Sub ToggleLive()
    gLiveOn = Not gLiveOn
    If gLiveOn Then
        gNextTick = Now + TimeSerial(0, 0, LIVE_INTERVAL_SEC)
        Application.OnTime gNextTick, "AutoRefresh"
        MsgBox "Live refresh ON  (every " & LIVE_INTERVAL_SEC & " sec).", _
               vbInformation, "Machine Monitor"
    Else
        On Error Resume Next
        Application.OnTime gNextTick, "AutoRefresh", , False
        On Error GoTo 0
        MsgBox "Live refresh OFF.", vbInformation, "Machine Monitor"
    End If
End Sub

Public Sub AutoRefresh()
    If Not gLiveOn Then Exit Sub
    RefreshDashboard
    gNextTick = Now + TimeSerial(0, 0, LIVE_INTERVAL_SEC)
    Application.OnTime gNextTick, "AutoRefresh"
End Sub
