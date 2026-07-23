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
'     SetupWorkbook        — installs shape buttons, hotkeys
'     ImportCSV            — batch-load a CSV into the Data sheet
'     RefreshDashboard     — recompute picker + comparison charts
'     GenerateReport       — new sheet snapshot + PDF export
'     ClearData            — wipe the Data table (with confirm)
'     AddMachine           — append a machine from the Dashboard ADD panel
'     RefreshMachinePicker — populate the Dashboard's machine picker
'     ToggleLive           — enable/disable 60-second auto-refresh
'==========================================================================
Option Explicit

Private Const SHT_DASH As String    = "Dashboard"
Private Const SHT_DATA As String    = "Data"
Private Const SHT_CONF As String    = "Config"
Private Const SHT_REPORTS As String = "Reports"
Private Const TBL_DATA As String    = "tblData"
Private Const TBL_MACH As String    = "tblMachines"

' Picker is 2 groups of columns × 8 rows = 16 slots
'   Group 1: Show=B, ID=C, Name=D   rows 20..27
'   Group 2: Show=F, ID=G, Name=H   rows 20..27
Private Const PICK_R0 As Long = 20
Private Const PICK_ROWS As Long = 8

Private Const LIVE_INTERVAL_SEC As Long = 60
Public gLiveOn As Boolean
Public gNextTick As Double

'--------------------------------------------------------------------------
'  ONE-TIME SETUP
'--------------------------------------------------------------------------
Public Sub SetupWorkbook()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_DASH)

    Dim shp As Shape, i As Long
    For i = ws.Shapes.count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left$(shp.Name, 3) = "btn" Then shp.Delete
    Next i

    AddButton ws, "btnImport",  ws.Range("J7:L7"),  "IMPORT  CSV",        "ImportCSV"
    AddButton ws, "btnRefresh", ws.Range("J8:L8"),  "REFRESH  DASHBOARD", "RefreshDashboard"
    AddButton ws, "btnReport",  ws.Range("N7:P7"),  "GENERATE  REPORT",   "GenerateReport"
    AddButton ws, "btnClear",   ws.Range("N8:P8"),  "CLEAR  DATA",        "ClearData"
    AddButton ws, "btnAdd",     ws.Range("N16:P16"), "ADD  MACHINE",      "AddMachine"

    Application.OnKey "^+i", "ImportCSV"
    Application.OnKey "^+r", "RefreshDashboard"
    Application.OnKey "^+g", "GenerateReport"

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
        .Fill.ForeColor.RGB = RGB(232, 163, 61)
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
'  ADD  MACHINE  (from the Dashboard ADD panel)
'--------------------------------------------------------------------------
Public Sub AddMachine()
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets(SHT_DASH)
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Worksheets(SHT_CONF)

    Dim mid As String: mid = Trim$(CStr(wsD.Range("AddMID").Value))
    Dim nm As String:  nm  = Trim$(CStr(wsD.Range("AddName").Value))
    Dim tg As String:  tg  = Trim$(CStr(wsD.Range("AddTag").Value))
    Dim m1 As String:  m1  = Trim$(CStr(wsD.Range("AddM1").Value))
    Dim m2 As String:  m2  = Trim$(CStr(wsD.Range("AddM2").Value))
    Dim m3 As String:  m3  = Trim$(CStr(wsD.Range("AddM3").Value))

    If Len(mid) = 0 Or Len(nm) = 0 Or Len(tg) = 0 Then
        MsgBox "Fill in at least Machine ID, Name, and Tag.", _
               vbExclamation, "Add machine"
        Exit Sub
    End If

    Dim lo As ListObject: Set lo = wsC.ListObjects(TBL_MACH)

    ' Duplicate-ID guard
    If Not lo.DataBodyRange Is Nothing Then
        Dim arr As Variant: arr = lo.DataBodyRange.Value
        Dim i As Long
        For i = 1 To UBound(arr, 1)
            If UCase$(Trim$(CStr(arr(i, 1)))) = UCase$(mid) Then
                MsgBox "Machine ID '" & mid & "' already exists.", _
                       vbExclamation, "Add machine"
                Exit Sub
            End If
        Next i
    End If

    ' Find first empty row inside the table
    Dim r As Long, targetRow As Range
    Set targetRow = Nothing
    If lo.DataBodyRange Is Nothing Then
        Set targetRow = lo.ListRows.Add.Range
    Else
        For i = 1 To lo.DataBodyRange.Rows.count
            If Len(Trim$(CStr(lo.DataBodyRange.Cells(i, 1).Value))) = 0 Then
                Set targetRow = lo.DataBodyRange.Rows(i)
                Exit For
            End If
        Next i
        If targetRow Is Nothing Then
            If lo.DataBodyRange.Rows.count >= 16 Then
                MsgBox "The 16-machine limit has been reached.", _
                       vbExclamation, "Add machine"
                Exit Sub
            End If
            Set targetRow = lo.ListRows.Add.Range
        End If
    End If

    With targetRow
        .Cells(1, 1).Value = mid
        .Cells(1, 2).Value = nm
        .Cells(1, 3).Value = tg
        .Cells(1, 4).Value = IIf(Len(m1) = 0, "Metric 1", m1)
        .Cells(1, 5).Value = IIf(Len(m2) = 0, "Metric 2", m2)
        .Cells(1, 6).Value = IIf(Len(m3) = 0, "Metric 3", m3)
        .Cells(1, 7).Value = "Yes"
    End With

    ' Clear the input row
    wsD.Range("AddMID,AddName,AddTag,AddM1,AddM2,AddM3").ClearContents

    ' If the new machine belongs to the currently selected tag, add it to picker
    If tg = CStr(wsD.Range("TagFilter").Value) Then RefreshMachinePicker True

    MsgBox "Added " & mid & " (" & nm & ") under tag '" & tg & "'.", _
           vbInformation, "Add machine"
End Sub

'--------------------------------------------------------------------------
'  REFRESH  MACHINE  PICKER
'   forceReset = True  → wipe & repopulate with all Yes
'   forceReset = False → preserve existing Yes/No selections if tag unchanged
'--------------------------------------------------------------------------
Public Sub RefreshMachinePicker(Optional forceReset As Boolean = False)
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets(SHT_DASH)
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Worksheets(SHT_CONF)

    Dim tag As String
    tag = Trim$(CStr(wsD.Range("TagFilter").Value))

    Dim lastTag As String
    lastTag = CStr(wsD.Range("LastTag").Value)

    ' Snapshot current picker (mid → show) so we can preserve on same-tag refresh
    Dim prevSel As Object
    Set prevSel = CreateObject("Scripting.Dictionary")
    If Not forceReset And lastTag = tag Then
        Dim rr As Long, k As Long
        For rr = 0 To PICK_ROWS - 1
            k = PICK_R0 + rr
            SnapshotOne wsD, prevSel, "C" & k, "B" & k   ' group 1
            SnapshotOne wsD, prevSel, "G" & k, "F" & k   ' group 2
        Next rr
    End If

    ' Wipe picker area
    Dim rClear As Long
    For rClear = PICK_R0 To PICK_R0 + PICK_ROWS - 1
        wsD.Range("B" & rClear & ":D" & rClear).ClearContents
        wsD.Range("F" & rClear & ":H" & rClear).ClearContents
    Next rClear

    ' Gather machines from Config matching tag & Active=Yes
    Dim lo As ListObject: Set lo = wsC.ListObjects(TBL_MACH)
    If lo.DataBodyRange Is Nothing Then GoTo Finalize
    Dim arr As Variant: arr = lo.DataBodyRange.Value

    Dim i As Long, slot As Long: slot = 0
    Dim mid As String, nm As String, tg As String, act As String
    For i = 1 To UBound(arr, 1)
        mid = Trim$(CStr(arr(i, 1)))
        nm  = Trim$(CStr(arr(i, 2)))
        tg  = Trim$(CStr(arr(i, 3)))
        act = UCase$(Trim$(CStr(arr(i, 7))))
        If Len(mid) > 0 And tg = tag And act = "YES" Then
            Dim rowIdx As Long, groupCol As String
            If slot < PICK_ROWS Then
                rowIdx = PICK_R0 + slot
                wsD.Range("B" & rowIdx).Value = _
                    IIf(prevSel.Exists(mid), prevSel(mid), "Yes")
                wsD.Range("C" & rowIdx).Value = mid
                wsD.Range("D" & rowIdx).Value = nm
            ElseIf slot < PICK_ROWS * 2 Then
                rowIdx = PICK_R0 + (slot - PICK_ROWS)
                wsD.Range("F" & rowIdx).Value = _
                    IIf(prevSel.Exists(mid), prevSel(mid), "Yes")
                wsD.Range("G" & rowIdx).Value = mid
                wsD.Range("H" & rowIdx).Value = nm
            End If
            slot = slot + 1
        End If
    Next i

Finalize:
    wsD.Range("LastTag").Value = tag
End Sub

Private Sub SnapshotOne(ws As Worksheet, d As Object, midCell As String, showCell As String)
    Dim mid As String, sh As String
    mid = Trim$(CStr(ws.Range(midCell).Value))
    sh  = Trim$(CStr(ws.Range(showCell).Value))
    If Len(mid) > 0 And Not d.Exists(mid) Then d.Add mid, sh
End Sub

Private Function CollectSelectedMachines(ws As Worksheet) As Object
    ' Returns dict of  machine_id -> True  for every ticked slot in the picker.
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim rr As Long, k As Long, mid As String, sh As String
    For rr = 0 To PICK_ROWS - 1
        k = PICK_R0 + rr
        mid = Trim$(CStr(ws.Range("C" & k).Value))
        sh  = UCase$(Trim$(CStr(ws.Range("B" & k).Value)))
        If Len(mid) > 0 And sh = "YES" Then d(mid) = True
        mid = Trim$(CStr(ws.Range("G" & k).Value))
        sh  = UCase$(Trim$(CStr(ws.Range("F" & k).Value)))
        If Len(mid) > 0 And sh = "YES" Then d(mid) = True
    Next rr
    Set CollectSelectedMachines = d
End Function

'--------------------------------------------------------------------------
'  IMPORT  CSV
'--------------------------------------------------------------------------
Public Sub ImportCSV()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    fd.Title = "Select CSV (machine_id, tag, timestamp, metric1, metric2, metric3)"
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

    Dim line As String, parts() As String, isFirst As Boolean
    isFirst = True

    Dim added As Long, newRow As ListRow
    Do While Not EOF(ff)
        Line Input #ff, line
        If Len(Trim$(line)) = 0 Then GoTo Continue
        parts = SplitCSV(line)
        If UBound(parts) < 5 Then GoTo Continue

        If isFirst Then
            isFirst = False
            If Not IsDate(parts(2)) Then GoTo Continue  ' skip header
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

    ' 1. Populate/refresh the machine picker for the current tag
    RefreshMachinePicker False

    ' 2. Read filters
    Dim dStart As Date, dEnd As Date, tag As String
    dStart = wsD.Range("StartDate").Value
    dEnd   = wsD.Range("EndDate").Value
    tag    = CStr(wsD.Range("TagFilter").Value)

    ' 3. Collect ticked machines
    Dim sel As Object: Set sel = CollectSelectedMachines(wsD)
    Dim machines() As String, mCount As Long
    ReDim machines(0 To 15)
    mCount = 0

    Dim keys As Variant
    keys = sel.keys
    Dim i As Long
    For i = 0 To sel.count - 1
        machines(mCount) = CStr(keys(i))
        mCount = mCount + 1
    Next i

    ' 4. Clear staging (rows 30..500, cols R:AF)
    Application.ScreenUpdating = False
    wsD.Range("R30:AF500").ClearContents

    ' 5. Header row on staging (row 30)
    wsD.Cells(30, 18).Value = "Date"
    For i = 0 To mCount - 1
        wsD.Cells(30, 19 + i).Value = machines(i)
    Next i

    If mCount = 0 Then GoTo Rebind

    ' 6. Metric focus → column offset inside DataBody array
    Dim focus As String: focus = CStr(wsD.Range("MetricFocus").Value)
    Dim colOffset As Long
    Select Case focus
        Case "Metric 1": colOffset = 4
        Case "Metric 2": colOffset = 5
        Case "Metric 3": colOffset = 6
        Case Else:        colOffset = 4
    End Select

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lo As ListObject: Set lo = wsX.ListObjects(TBL_DATA)
    Dim rng As Range: Set rng = lo.DataBodyRange
    If rng Is Nothing Then GoTo Rebind

    Dim arr As Variant: arr = rng.Value
    Dim n As Long: n = UBound(arr, 1)

    Dim sums() As Double, cnts() As Long
    ReDim sums(1 To 400, 0 To mCount - 1)
    ReDim cnts(1 To 400, 0 To mCount - 1)
    Dim dateList() As Date
    ReDim dateList(1 To 400)
    Dim dCount As Long: dCount = 0

    Dim ts As Date, mid As String, tg As String, key As String, mIdx As Long, dIdx As Long
    Dim v As Double
    For i = 1 To n
        mid = CStr(arr(i, 1))
        tg  = CStr(arr(i, 2))
        If tg = tag And sel.Exists(mid) Then
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
                        v = CDblSafe(CStr(arr(i, colOffset)))
                        sums(dIdx, mIdx) = sums(dIdx, mIdx) + v
                        cnts(dIdx, mIdx) = cnts(dIdx, mIdx) + 1
                    End If
                End If
            End If
        End If
    Next i

    If dCount > 0 Then
        QuickSortDates dateList, 1, dCount
        Dim rowOut As Long
        For rowOut = 1 To dCount
            wsD.Cells(30 + rowOut, 18).Value = dateList(rowOut)
            wsD.Cells(30 + rowOut, 18).NumberFormat = "yyyy-mm-dd"
            key = Format$(dateList(rowOut), "yyyy-mm-dd")
            dIdx = dict(key)
            For mIdx = 0 To mCount - 1
                If cnts(dIdx, mIdx) > 0 Then
                    wsD.Cells(30 + rowOut, 19 + mIdx).Value = sums(dIdx, mIdx) / cnts(dIdx, mIdx)
                End If
            Next mIdx
        Next rowOut
    End If

Rebind:
    RebindCharts wsD, mCount, IIf(dCount > 0, dCount, 0)
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
    Dim lastRow As Long: lastRow = 30 + dCount

    For Each co In ws.ChartObjects
        Set ch = co.Chart
        Do While ch.SeriesCollection.count > 0
            ch.SeriesCollection(1).Delete
        Loop
        If mCount = 0 Or dCount = 0 Then GoTo NextChart

        Dim m As Long
        For m = 1 To mCount
            With ch.SeriesCollection.NewSeries
                .Name = "='" & ws.Name & "'!" & ws.Cells(30, 18 + m).Address
                .XValues = ws.Range(ws.Cells(31, 18), ws.Cells(lastRow, 18))
                .Values  = ws.Range(ws.Cells(31, 18 + m), ws.Cells(lastRow, 18 + m))
                .Format.Line.Weight = 1.75
                .Smooth = True
                Select Case (m - 1) Mod 8
                    Case 0: .Format.Line.ForeColor.RGB = RGB(63, 167, 184)
                    Case 1: .Format.Line.ForeColor.RGB = RGB(232, 163, 61)
                    Case 2: .Format.Line.ForeColor.RGB = RGB(155, 197, 61)
                    Case 3: .Format.Line.ForeColor.RGB = RGB(214, 69, 69)
                    Case 4: .Format.Line.ForeColor.RGB = RGB(120, 90, 200)
                    Case 5: .Format.Line.ForeColor.RGB = RGB(38, 122, 132)
                    Case 6: .Format.Line.ForeColor.RGB = RGB(200, 122, 31)
                    Case 7: .Format.Line.ForeColor.RGB = RGB(90, 130, 40)
                End Select
            End With
        Next m
NextChart:
    Next co
End Sub

'--------------------------------------------------------------------------
'  GENERATE  REPORT
'--------------------------------------------------------------------------
Public Sub GenerateReport()
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets(SHT_DASH)
    Dim wsR As Worksheet: Set wsR = ThisWorkbook.Worksheets(SHT_REPORTS)

    Dim dStart As Date: dStart = wsD.Range("StartDate").Value
    Dim dEnd As Date:   dEnd = wsD.Range("EndDate").Value
    Dim tag As String:  tag = CStr(wsD.Range("TagFilter").Value)

    Dim stamp As String: stamp = Format$(Now, "yyyymmdd_hhnnss")
    Dim shtName As String: shtName = "Rpt_" & tag & "_" & stamp
    If Len(shtName) > 31 Then shtName = Left$(shtName, 31)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    wsD.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
    Dim snap As Worksheet: Set snap = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
    snap.Name = shtName
    snap.Tab.Color = RGB(155, 197, 61)
    snap.UsedRange.Value = snap.UsedRange.Value

    snap.Range("B63").Value = "  Report generated " & Format$(Now, "yyyy-mm-dd hh:mm") & _
                              "  |  Tag: " & tag & _
                              "  |  Range: " & Format$(dStart, "yyyy-mm-dd") & _
                              " → " & Format$(dEnd, "yyyy-mm-dd")

    Dim pdfPath As String
    pdfPath = ThisWorkbook.path & Application.PathSeparator & shtName & ".pdf"
    On Error Resume Next
    snap.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0

    Dim r As Long: r = wsR.Cells(wsR.Rows.count, 2).End(xlUp).Row + 1
    If r < 5 Then r = 5
    wsR.Cells(r, 2).Value = Now: wsR.Cells(r, 2).NumberFormat = "yyyy-mm-dd hh:mm"
    wsR.Cells(r, 3).Value = tag
    wsR.Cells(r, 4).Value = dStart: wsR.Cells(r, 4).NumberFormat = "yyyy-mm-dd"
    wsR.Cells(r, 5).Value = dEnd:   wsR.Cells(r, 5).NumberFormat = "yyyy-mm-dd"
    wsR.Cells(r, 6).Value = SelectedMachineCount(wsD)
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

Private Function SelectedMachineCount(ws As Worksheet) As Long
    Dim d As Object: Set d = CollectSelectedMachines(ws)
    SelectedMachineCount = d.count
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
'  LIVE  UPDATES
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
