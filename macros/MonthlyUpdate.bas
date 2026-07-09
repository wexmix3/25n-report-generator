Attribute VB_Name = "MonthlyUpdate"
' =============================================================================
' 25N Coworking - Monthly Dashboard Update Macro
' =============================================================================

Option Explicit

' Module-level step tracker so ErrorHandler can report which step failed
Private mCurrentStep As String

' -- Sheet names --------------------------------------------------------------
Private Const SBM_SHEET        As String = "2026 Sales By Month"
Private Const P1_SHEET         As String = "P1. Dashboard"
Private Const P4_SHEET         As String = "P.4 Forecast & Budget to Actual"
Private Const P3_SHEET         As String = "P3. Occupancy"
Private Const INSTR_SHEET      As String = "Instructions"

' -- P1 snapshot column positions (VBA 1-based) -------------------------------
' Col B = Month serial, C = Occ% Spaces, D = Occ% SqFt
' Col E = Contracts, F = Flex Members, G = Mail Members
' Col H = (empty spacer), Col I = Meeting Hours
Private Const P1_COL_MONTH     As Long = 2
Private Const P1_COL_OCC_SPC   As Long = 3
Private Const P1_COL_OCC_SQFT  As Long = 4
Private Const P1_COL_CONTRACTS As Long = 5
Private Const P1_COL_FLEX      As Long = 6
Private Const P1_COL_MAIL      As Long = 7
Private Const P1_COL_MTG_HRS   As Long = 9
Private Const P1_SNAPSHOT_ROWS As Long = 12

' -- Raw data tabs to hide (source exports that should not be visible) --------
Private Const RAW_TABS As String = "Occupancy Data|Budget Actual Data"

' =============================================================================
' ENTRY POINT
' =============================================================================
Public Sub MonthlyUpdate()
    Dim bridgePath As String
    Dim bridgeWb   As Workbook
    Dim dashWb     As Workbook

    Set dashWb = ThisWorkbook

    bridgePath = PickDataBridgeFile()
    If bridgePath = "" Then
        MsgBox "No file selected -- update cancelled.", vbInformation, "Monthly Update"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    ' Save a timestamped backup before making any changes
    Call BackupWorkbook(dashWb)

    Set bridgeWb = Workbooks.Open(Filename:=bridgePath, ReadOnly:=True, UpdateLinks:=False)

    Dim meta As Object
    Set meta = ReadMetaSheet(bridgeWb)

    Dim monthCode         As String
    Dim monthSerial       As Long
    Dim occPctSpaces      As Double
    Dim occPctSqFt        As Double
    Dim contracts         As Long
    Dim flexMembers       As Long
    Dim mailMembers       As Long
    Dim meetingHours      As Double
    Dim meetingHoursPrior As Double
    Dim dailySpaceCurrent As Double
    Dim dailySpacePrior   As Double

    ' month_code: Excel can auto-convert "May-26" to a date serial on open.
    ' Detect that case and format it back to "mmm-yy".
    ' NOTE: VBA Or/And are NOT short-circuit -- use nested If to avoid CDbl on non-numeric values.
    Dim rawMc As Variant
    Dim needDateConvert As Boolean
    rawMc = meta("month_code")
    needDateConvert = False
    If IsDate(rawMc) Then
        needDateConvert = True
    ElseIf IsNumeric(rawMc) Then
        If CDbl(rawMc) > 40000 And CDbl(rawMc) < 60000 Then
            needDateConvert = True
        End If
    End If

    If needDateConvert Then
        On Error Resume Next
        monthCode = Format(CDate(rawMc), "mmm-yy")
        On Error GoTo 0
        If monthCode = "" Then monthCode = CStr(rawMc)
    Else
        monthCode = CStr(rawMc)
    End If
    monthSerial       = CLng(meta("month_serial"))
    occPctSpaces      = CDbl(meta("occupancy_pct_spaces"))
    occPctSqFt        = CDbl(meta("occupancy_pct_sqft"))
    contracts         = CLng(meta("contracts"))
    flexMembers       = CLng(meta("flex_members"))
    mailMembers       = CLng(meta("mail_members"))
    meetingHours      = CDbl(meta("meeting_hours"))
    meetingHoursPrior = CDbl(meta("meeting_hours_prior"))
    dailySpaceCurrent = CDbl(meta("daily_space_current"))
    dailySpacePrior   = CDbl(meta("daily_space_prior"))

    If monthSerial = 0 Then
        MsgBox "Data Bridge has no month data (month_serial = 0)." & Chr(13) & Chr(13) & _
               "Check that the correct source files were uploaded.", _
               vbExclamation, "Monthly Update"
        GoTo Cleanup
    End If

    Dim confirmed As VbMsgBoxResult
    confirmed = MsgBox("About to update the dashboard for: " & monthCode & Chr(13) & Chr(13) & _
                       "This will:" & Chr(13) & _
                       "  - Paste Sales By Month data" & Chr(13) & _
                       "  - Insert P.1 snapshot row (occupancy % auto-filled)" & Chr(13) & _
                       "  - Update P.4 formula column references" & Chr(13) & _
                       "  - Add P.3 column for " & monthCode & " (occupied spaces carry forward rate)" & Chr(13) & _
                       "  - Write P.2 meeting room hours & revenue (if Booking Summary uploaded)" & Chr(13) & _
                       "  - Refresh Prior Month Meeting Hours tab" & Chr(13) & _
                       "  - Repopulate Daily Space tab (year filter + transaction rows)" & Chr(13) & _
                       "  - Repopulate Flex Members tab (income detail rows)" & Chr(13) & _
                       "  - Repopulate Mail Members tab (income detail rows)" & Chr(13) & Chr(13) & _
                       "Note: Verify check figures = 0 on Flex, Mail, and Daily Space tabs after running.", _
                       vbYesNo + vbQuestion, "Monthly Update - " & monthCode)
    If confirmed = vbNo Then GoTo Cleanup

    mCurrentStep = "UpdateSalesByMonth"
    Call UpdateSalesByMonth(dashWb, bridgeWb, monthCode)
    mCurrentStep = "UpdateOccupancyDataTab"
    Call UpdateOccupancyDataTab(dashWb, bridgeWb)
    mCurrentStep = "InsertP1SnapshotRow"
    Call InsertP1SnapshotRow(dashWb, monthSerial, monthCode, occPctSpaces, occPctSqFt, contracts, flexMembers, mailMembers, meetingHours)
    mCurrentStep = "UpdateP4ColumnRefs"
    Call UpdateP4ColumnRefs(dashWb, monthSerial)
    mCurrentStep = "UpdateP1H4Formula"
    Call UpdateP1H4Formula(dashWb, monthSerial)
    mCurrentStep = "AddP3MonthColumn"
    Call AddP3MonthColumn(dashWb, bridgeWb, monthCode)
    mCurrentStep = "UpdateP2DailySpace"
    Call UpdateP2DailySpace(dashWb, monthSerial, dailySpaceCurrent, dailySpacePrior)
    mCurrentStep = "UpdateP2MeetingRoomData"
    Call UpdateP2MeetingRoomData(dashWb, bridgeWb, monthSerial)
    mCurrentStep = "UpdatePriorMonthMeetingHours"
    Call UpdatePriorMonthMeetingHours(dashWb, bridgeWb)
    mCurrentStep = "UpdateDailySpaceTab"
    Call UpdateDailySpaceTab(dashWb, bridgeWb, monthSerial)
    mCurrentStep = "UpdateFlexTab"
    Call UpdateFlexTab(dashWb, bridgeWb, monthSerial)
    mCurrentStep = "UpdateMailMembersTab"
    Call UpdateMailMembersTab(dashWb, bridgeWb, monthSerial)
    mCurrentStep = "DeleteStaleTabs"
    Call DeleteStaleTabs(dashWb)
    mCurrentStep = "UpdateYardiIncDetailTab"
    Call UpdateYardiIncDetailTab(dashWb, bridgeWb)
    mCurrentStep = "UpdateYardiAdjOccTab"
    Call UpdateYardiAdjOccTab(dashWb, bridgeWb)
    mCurrentStep = "FormatP1DateColumn"
    Call FormatP1DateColumn(dashWb)
    mCurrentStep = "HideRawTabs"
    Call HideRawTabs(dashWb)
    mCurrentStep = "RewriteInstructions"
    Call RewriteInstructions(dashWb)
    mCurrentStep = "RunPostValidation"
    Call RunPostValidation(dashWb, bridgeWb, monthCode, monthSerial)

    Dim mtgNote As String
    If meetingHours > 0 Then
        mtgNote = "Meeting Hours (col I): " & Format(meetingHours, "0.0") & " hrs"
    Else
        mtgNote = "Meeting Hours (col I): not found -- enter manually"
    End If
    Dim p2MtgNote As String
    If meetingHours > 0 Then
        p2MtgNote = "P.2 mtg room hrs/revenue: auto-filled"
    Else
        p2MtgNote = "P.2 mtg room hrs/revenue: enter manually (no Booking Summary)"
    End If
    MsgBox "Monthly update complete for " & monthCode & "." & Chr(13) & Chr(13) & _
           "AUTO-FILLED:" & Chr(13) & _
           "  P.1: Occ%, Contracts (" & contracts & "), Flex (" & flexMembers & "), Mail (" & mailMembers & ")" & Chr(13) & _
           "  " & mtgNote & Chr(13) & _
           "  " & p2MtgNote & Chr(13) & _
           "  P.2 Daily Space, P.3 column, P.4 refs, SBM data" & Chr(13) & _
           "  Daily Space tab, Flex Members tab, Mail Members tab" & Chr(13) & _
           "  Prior Month Mtg Hrs tab" & Chr(13) & Chr(13) & _
           "STILL MANUAL:" & Chr(13) & _
           "  1. Verify check figure = 0: Flex, Mail, Daily Space tabs" & Chr(13) & _
           "  2. P.3 -- verify any new tenant contract rates" & Chr(13) & _
           "  3. P.4 col F -- enter Budget figures from Voyager" & Chr(13) & _
           "  4. P.2 -- verify check figure = 0 for all months" & Chr(13) & _
           "  5. Save + backup to finance drive", _
           vbInformation, "Monthly Update - Done"

Cleanup:
    If Not bridgeWb Is Nothing Then bridgeWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.Calculate
    Exit Sub

ErrorHandler:
    MsgBox "Error in step [" & mCurrentStep & "]: " & Err.Description & Chr(13) & _
           "Error #" & Err.Number & " | Source: " & Err.Source, _
           vbCritical, "Monthly Update Error"
    Resume Cleanup
End Sub

' =============================================================================
' CLEANUP P.1 DUPLICATE ROWS
' Run this once to purge duplicate snapshot rows (e.g. after multiple test runs).
' Keeps the most recently inserted row for each month serial; deletes extras.
' =============================================================================
' =============================================================================
' DIAGNOSE SNAPSHOT ROW VALUES
' Shows exact Value, TypeName, and IsNumeric for col B rows 55-80.
' Run this to understand what InsertP1SnapshotRow actually wrote.
' =============================================================================
Public Sub DiagnoseSnapshotRows()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(P1_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "P.1 not found", vbExclamation: Exit Sub

    Dim report As String
    report = "Col B (P1_COL_MONTH=" & P1_COL_MONTH & ") values, rows 55-80:" & Chr(13) & String(55, "-")
    Dim r As Long
    For r = 55 To 80
        Dim v As Variant
        v = ws.Cells(r, P1_COL_MONTH).Value
        Dim hasF As String
        hasF = IIf(ws.Cells(r, P1_COL_MONTH).HasFormula, " [FORMULA:" & ws.Cells(r, P1_COL_MONTH).Formula & "]", "")
        Dim info As String
        info = "R" & r & ": |" & CStr(v) & "| Type=" & TypeName(v) & " Num=" & IsNumeric(v)
        If IsNumeric(v) And Not IsError(v) Then info = info & " CLng=" & CLng(v)
        info = info & hasF
        report = report & Chr(13) & info
    Next r
    MsgBox report, vbInformation, "Snapshot Row Diagnosis"
End Sub

Public Sub CleanupP1DuplicateRows()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(P1_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & P1_SHEET & "' not found.", vbExclamation, "Cleanup"
        Exit Sub
    End If

    ' Scan rows 200 → 10 using CellToDateSerial() which correctly handles the
    ' Date variant that Excel returns for date-formatted cells.
    ' (IsNumeric() returns False for Date variants -- the root cause of prior failures.)
    ' First occurrence of each serial encountered going downward (= newest row,
    ' closest to Average) is kept; earlier rows with same serial are duplicates.
    Dim seen         As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim deleteList() As Long
    Dim deleteCount  As Long
    deleteCount = 0
    ReDim deleteList(200)

    Dim r As Long
    For r = 200 To 10 Step -1
        Dim v As Variant
        v = ws.Cells(r, P1_COL_MONTH).Value
        Dim s As Long
        s = CellToDateSerial(v)
        If s > 0 Then
            Dim key As String
            key = CStr(s)   ' use the serial as the dictionary key
            If seen.Exists(key) Then
                If deleteCount <= UBound(deleteList) Then deleteList(deleteCount) = r
                deleteCount = deleteCount + 1
            Else
                seen(key) = r
            End If
        End If
NextScan:
    Next r

    If deleteCount = 0 Then
        MsgBox "No duplicate snapshot rows found -- P.1 is clean.", vbInformation, "Cleanup"
        Exit Sub
    End If

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Found " & deleteCount & " duplicate snapshot row(s) to delete." & Chr(13) & Chr(13) & _
                     "Proceed?", vbYesNo + vbQuestion, "Cleanup Duplicates")
    If confirm = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Delete highest row numbers first so lower row indices stay valid
    Dim i As Long
    Dim j As Long
    For i = 0 To deleteCount - 2
        For j = 0 To deleteCount - 2 - i
            If deleteList(j) < deleteList(j + 1) Then
                Dim tmp As Long
                tmp = deleteList(j)
                deleteList(j) = deleteList(j + 1)
                deleteList(j + 1) = tmp
            End If
        Next j
    Next i
    For i = 0 To deleteCount - 1
        ws.Rows(deleteList(i)).Delete
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Deleted " & deleteCount & " duplicate row(s). P.1 snapshot block is now clean.", _
           vbInformation, "Cleanup Complete"
End Sub

' =============================================================================
' DIAGNOSE P.1 FORMULA ERRORS
' Lists every cell in P.1 Dashboard that shows an error and its formula text.
' Run this to identify which formula patterns still need the RegExp fix.
' =============================================================================
Public Sub DiagnoseP1Errors()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(P1_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & P1_SHEET & "' not found.", vbExclamation, "Diagnose"
        Exit Sub
    End If

    Dim report    As String
    Dim errCount  As Long
    errCount = 0
    report = "ERROR CELLS IN P.1 DASHBOARD" & Chr(13) & String(50, "-") & Chr(13) & Chr(13)

    Dim cell As Range
    For Each cell In ws.UsedRange
        If IsError(cell.Value) Then
            errCount = errCount + 1
            Dim formulaText As String
            formulaText = ""
            If cell.HasFormula Then formulaText = cell.Formula
            report = report & "Cell: " & cell.Address & Chr(13)
            If formulaText <> "" Then
                report = report & "Formula: " & Left(formulaText, 200) & Chr(13)
            Else
                report = report & "(no formula -- direct error value)" & Chr(13)
            End If
            report = report & "Error: " & CStr(cell.Value) & Chr(13) & Chr(13)
        End If
    Next cell

    If errCount = 0 Then
        MsgBox "No formula errors found in P.1 Dashboard.", vbInformation, "Diagnose"
    Else
        report = report & String(50, "-") & Chr(13) & errCount & " error cell(s) found."
        MsgBox report, vbExclamation, "P.1 Errors - " & errCount & " found"
    End If
End Sub

' =============================================================================
' ONE-TIME SETUP -- run once after importing this macro
' Hides raw data tabs and rewrites the Instructions sheet
' =============================================================================
Public Sub RunSetup()
    Dim dashWb As Workbook
    Set dashWb = ThisWorkbook
    Application.ScreenUpdating = False
    Call HideRawTabs(dashWb)
    Call RewriteInstructions(dashWb)
    Call FormatP1DateColumn(dashWb)
    Application.ScreenUpdating = True
    MsgBox "Setup complete." & Chr(13) & Chr(13) & _
           "- Raw data tabs are now hidden" & Chr(13) & _
           "- Instructions tab rewritten" & Chr(13) & _
           "- P.1 date column formatted as Mon-YY", _
           vbInformation, "Setup"
End Sub

' =============================================================================
' READ META SHEET
' =============================================================================
Private Function ReadMetaSheet(wb As Workbook) As Object
    Dim ws As Worksheet
    Dim d  As Object
    Set d = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set ws = wb.Sheets("Meta")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Data Bridge is missing the 'Meta' sheet.", vbExclamation, "Monthly Update"
        Set ReadMetaSheet = d
        Exit Function
    End If

    Dim r As Long
    For r = 2 To 50
        Dim k As String
        k = Trim(CStr(ws.Cells(r, 1).Value))
        If k = "" Then Exit For
        d(k) = ws.Cells(r, 2).Value
    Next r

    Set ReadMetaSheet = d
End Function

' =============================================================================
' UPDATE SALES BY MONTH
' =============================================================================
Private Sub UpdateSalesByMonth(dashWb As Workbook, bridgeWb As Workbook, monthCode As String)
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet

    On Error Resume Next
    Set srcWs = bridgeWb.Sheets("SBM")
    Set dstWs = dashWb.Sheets(SBM_SHEET)
    On Error GoTo 0

    If srcWs Is Nothing Then Exit Sub
    If dstWs Is Nothing Then
        MsgBox "Sheet '" & SBM_SHEET & "' not found.", vbExclamation, "Monthly Update"
        Exit Sub
    End If

    Dim srcLastCol As Long
    srcLastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

    Dim srcMonthCols As Object
    Set srcMonthCols = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 2 To srcLastCol
        Dim mc As String
        mc = Trim(CStr(srcWs.Cells(1, c).Value))
        If mc <> "" Then srcMonthCols(mc) = c
    Next c

    Dim dstHeaderRow As Long
    dstHeaderRow = FindSBMHeaderRow(dstWs)
    If dstHeaderRow = 0 Then
        MsgBox "Cannot find month header row in '" & SBM_SHEET & "'.", vbExclamation, "Monthly Update"
        Exit Sub
    End If

    Dim dstLastCol As Long
    dstLastCol = dstWs.Cells(dstHeaderRow, dstWs.Columns.Count).End(xlToLeft).Column

    Dim dstMonthCols As Object
    Set dstMonthCols = CreateObject("Scripting.Dictionary")
    For c = 2 To dstLastCol
        Dim cellVal As Variant
        cellVal = dstWs.Cells(dstHeaderRow, c).Value
        Dim dstMc As String
        dstMc = ""
        If IsNumeric(cellVal) Then
            If CLng(cellVal) > 40000 Then
                dstMc = Format(CDate(cellVal), "mmm-yy")
            Else
                dstMc = Trim(CStr(cellVal))
            End If
        Else
            dstMc = Trim(CStr(cellVal))
        End If
        If dstMc <> "" Then dstMonthCols(dstMc) = c
    Next c

    Dim dstLastRow As Long
    dstLastRow = dstWs.Cells(dstWs.Rows.Count, 1).End(xlUp).Row
    Dim dstLabelRows As Object
    Set dstLabelRows = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = dstHeaderRow + 1 To dstLastRow
        Dim lbl As String
        lbl = LCase(Trim(CStr(dstWs.Cells(r, 1).Value)))
        If lbl <> "" Then dstLabelRows(lbl) = r
    Next r

    Dim srcLastRow As Long
    srcLastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    Dim updatedCells As Long
    updatedCells = 0

    Dim srcMcKey As Variant
    For Each srcMcKey In srcMonthCols.Keys
        Dim srcCol As Long
        srcCol = srcMonthCols(srcMcKey)
        Dim dstCol As Long
        dstCol = 0
        If dstMonthCols.Exists(srcMcKey) Then
            dstCol = dstMonthCols(srcMcKey)
        Else
            Dim dstKey As Variant
            For Each dstKey In dstMonthCols.Keys
                If LCase(dstKey) = LCase(srcMcKey) Then
                    dstCol = dstMonthCols(dstKey)
                    Exit For
                End If
            Next dstKey
        End If
        If dstCol = 0 Then GoTo NextMonth

        For r = 2 To srcLastRow
            Dim srcLbl As String
            srcLbl = LCase(Trim(CStr(srcWs.Cells(r, 1).Value)))
            If srcLbl = "" Then GoTo NextRow
            If dstLabelRows.Exists(srcLbl) Then
                Dim dstRow As Long
                dstRow = dstLabelRows(srcLbl)
                Dim srcVal As Variant
                srcVal = srcWs.Cells(r, srcCol).Value
                If IsNumeric(srcVal) Then
                    dstWs.Cells(dstRow, dstCol).Value = CDbl(srcVal)
                    updatedCells = updatedCells + 1
                End If
            End If
NextRow:
        Next r
NextMonth:
    Next srcMcKey

    Application.StatusBar = "SBM: updated " & updatedCells & " cells"
End Sub

' -- Helper: find header row with >= 3 month indicators ----------------------
Private Function FindSBMHeaderRow(ws As Worksheet) As Long
    Dim r        As Long
    Dim c        As Long
    Dim hitCount As Long
    Dim cellStr  As String
    For r = 1 To 20
        hitCount = 0
        Dim lastC As Long
        lastC = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lastC
            Dim v As Variant
            v = ws.Cells(r, c).Value
            If IsNumeric(v) Then
                If CLng(v) > 40000 And CLng(v) < 70000 Then hitCount = hitCount + 1
            Else
                cellStr = LCase(Trim(CStr(v)))
                If cellStr Like "jan*" Or cellStr Like "feb*" Or cellStr Like "mar*" Or _
                   cellStr Like "apr*" Or cellStr Like "may*" Or cellStr Like "jun*" Or _
                   cellStr Like "jul*" Or cellStr Like "aug*" Or cellStr Like "sep*" Or _
                   cellStr Like "oct*" Or cellStr Like "nov*" Or cellStr Like "dec*" Then
                    hitCount = hitCount + 1
                End If
            End If
        Next c
        If hitCount >= 3 Then
            FindSBMHeaderRow = r
            Exit Function
        End If
    Next r
    FindSBMHeaderRow = 0
End Function

' =============================================================================
' INSERT P.1 SNAPSHOT ROW
' Auto-fills: occupancy %, contracts, flex members, mail members, meeting hours.
' Meeting hours come from the Kube Booking Summary (if uploaded); left blank otherwise.
' =============================================================================
Private Sub InsertP1SnapshotRow(dashWb As Workbook, _
                                 monthSerial As Long, _
                                 monthCode As String, _
                                 occSpaces As Double, _
                                 occSqFt As Double, _
                                 contracts As Long, _
                                 flexMembers As Long, _
                                 mailMembers As Long, _
                                 meetingHrs As Double)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = dashWb.Sheets(P1_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & P1_SHEET & "' not found.", vbExclamation, "Monthly Update"
        Exit Sub
    End If

    ' Find AVERAGE row -- scan from BOTTOM so KPI-section AVERAGE formulas
    ' (rows 10-17) don't shadow the actual snapshot Average row (~row 76).
    Dim avgRow As Long
    avgRow = 0
    Dim r As Long
    For r = 200 To 10 Step -1
        Dim bChk As Variant
        bChk = ws.Cells(r, P1_COL_MONTH).Value
        If Not IsError(bChk) Then
            If Not IsNumeric(bChk) Then
                If InStr(LCase(Trim(CStr(bChk))), "average") > 0 Then
                    avgRow = r
                    Exit For
                End If
            End If
        End If
    Next r
    If avgRow = 0 Then
        For r = 200 To 10 Step -1
            If ws.Cells(r, P1_COL_OCC_SPC).HasFormula Then
                If InStr(UCase(ws.Cells(r, P1_COL_OCC_SPC).Formula), "AVERAGE(") > 0 Then
                    avgRow = r
                    Exit For
                End If
            End If
        Next r
    End If

    If avgRow = 0 Then
        MsgBox "Cannot find AVERAGE row in '" & P1_SHEET & "'." & Chr(13) & _
               "Please insert the snapshot row manually for " & monthCode & ".", _
               vbExclamation, "Monthly Update"
        Exit Sub
    End If

    ' Check for duplicate (scan full snapshot block, not just prior row)
    Dim existRow As Long
    existRow = FindExistingSnapshotRow(ws, monthSerial, avgRow)
    If existRow > 0 Then
        Dim replaceIt As VbMsgBoxResult
        replaceIt = MsgBox("Month " & monthCode & " already exists at row " & existRow & "." & Chr(13) & Chr(13) & _
                           "Replace the existing row?", vbYesNo + vbQuestion, "Duplicate Detected")
        If replaceIt = vbYes Then
            ws.Rows(existRow).Delete
            ' Re-find AVERAGE row after deletion (same bottom-up logic)
            avgRow = 0
            For r = 200 To 10 Step -1
                Dim bChk2 As Variant
                bChk2 = ws.Cells(r, P1_COL_MONTH).Value
                If Not IsError(bChk2) Then
                    If Not IsNumeric(bChk2) Then
                        If InStr(LCase(Trim(CStr(bChk2))), "average") > 0 Then
                            avgRow = r
                            Exit For
                        End If
                    End If
                End If
            Next r
            If avgRow = 0 Then
                For r = 200 To 10 Step -1
                    If ws.Cells(r, P1_COL_OCC_SPC).HasFormula Then
                        If InStr(UCase(ws.Cells(r, P1_COL_OCC_SPC).Formula), "AVERAGE(") > 0 Then
                            avgRow = r
                            Exit For
                        End If
                    End If
                Next r
            End If
            If avgRow = 0 Then
                MsgBox "Cannot re-find AVERAGE row after deleting duplicate.", vbExclamation, "Monthly Update"
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If

    ' ── Step 1: Freeze the prior completed-month row (avgRow-1) before insert ──
    ' That row may still have live formulas (=E12, =H12, =Flex!K6, ='Mail Members'!O16)
    ' from when it was the "current" row.  Freeze C/D/F/G to locked values now so
    ' rolling history cannot be corrupted when the KPI block updates next month.
    ' Then set I (Meeting Hours Sold) to the Prior Month tab formula.
    Dim priorRow As Long
    priorRow = avgRow - 1
    If priorRow > 0 Then
        Dim freezeCols As Variant
        freezeCols = Array(P1_COL_OCC_SPC, P1_COL_OCC_SQFT, P1_COL_FLEX, P1_COL_MAIL)
        Dim fc As Variant
        For Each fc In freezeCols
            With ws.Cells(priorRow, CLng(fc))
                If .HasFormula Then .Value = .Value
            End With
        Next fc
        ws.Cells(priorRow, P1_COL_MTG_HRS).Formula = "='Prior Month Meeting Hours'!Q3"
    End If

    ' ── Step 2: Freeze meeting hours two months back (avgRow-2) ─────────────
    ' That row's I was set to ='Prior Month Meeting Hours'!Q3 in the prior run.
    ' Lock it to its historical value now that another month has passed.
    Dim oldPriorRow As Long
    oldPriorRow = avgRow - 2
    If oldPriorRow > 0 Then
        With ws.Cells(oldPriorRow, P1_COL_MTG_HRS)
            If .HasFormula Then .Value = .Value
        End With
    End If

    ' ── Step 3: Insert + format new row ─────────────────────────────────────
    ws.Rows(avgRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Rows(avgRow - 1).Copy
    ws.Rows(avgRow).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Update B2 -- period header date at top of P.1 dashboard (plain date value)
    ws.Cells(2, P1_COL_MONTH).Value = CDate(monthSerial)

    ' ── Step 4: Write new current-month row with live formulas ───────────────
    ' C/D/F/G use the same formulas as the master row for the current month.
    ' E (contracts) is left BLANK -- finance team enters this manually each month.
    ' I (meeting hours) is left BLANK -- the NEXT month's run populates it via
    ' Step 1 above, setting I = ='Prior Month Meeting Hours'!Q3 once that month closes.
    ws.Cells(avgRow, P1_COL_MONTH).Value = CDate(monthSerial)
    ws.Cells(avgRow, P1_COL_OCC_SPC).Formula  = "=E12"
    ws.Cells(avgRow, P1_COL_OCC_SQFT).Formula = "=H12"
    ws.Cells(avgRow, P1_COL_CONTRACTS).ClearContents
    ws.Cells(avgRow, P1_COL_FLEX).Formula     = "=Flex!K6"
    ws.Cells(avgRow, P1_COL_MAIL).Formula     = "='Mail Members'!O16"
    ws.Cells(avgRow, P1_COL_MTG_HRS).ClearContents

    ' Extend AVERAGE formulas
    Dim newAvgRow As Long
    newAvgRow = avgRow + 1
    Dim col As Long
    For col = P1_COL_OCC_SPC To P1_COL_MTG_HRS
        If ws.Cells(newAvgRow, col).HasFormula Then
            Dim oldFm As String
            oldFm = ws.Cells(newAvgRow, col).Formula
            Dim newFm As String
            newFm = ExtendAverageFormula(oldFm, 1)
            If newFm <> oldFm Then ws.Cells(newAvgRow, col).Formula = newFm
        End If
    Next col

    ' Hide oldest row to maintain rolling 12-month window
    Dim firstDataRow As Long
    firstDataRow = FindP1FirstDataRow(ws, newAvgRow)
    If firstDataRow > 0 Then
        Dim visibleCount As Long
        visibleCount = 0
        For r = firstDataRow To newAvgRow - 1
            If Not ws.Rows(r).Hidden Then visibleCount = visibleCount + 1
        Next r
        If visibleCount > P1_SNAPSHOT_ROWS Then
            For r = firstDataRow To newAvgRow - 1
                If Not ws.Rows(r).Hidden Then
                    ws.Rows(r).Hidden = True
                    Exit For
                End If
            Next r
        End If
    End If

    Application.StatusBar = "P.1: snapshot row inserted for " & monthCode
End Sub

' -- Helper: find first data row in P.1 snapshot block -----------------------
Private Function FindP1FirstDataRow(ws As Worksheet, avgRow As Long) As Long
    Dim r          As Long
    Dim firstFound As Long
    firstFound = 0
    For r = avgRow - 1 To 1 Step -1
        Dim v As Variant
        v = ws.Cells(r, P1_COL_MONTH).Value
        Dim s As Long
        s = CellToDateSerial(v)
        If s > 0 Then
            firstFound = r       ' keep scanning up -- we want the topmost date row
        ElseIf firstFound > 0 Then
            Exit For             ' hit a non-date cell after finding dates -- done
        End If
    Next r
    FindP1FirstDataRow = firstFound
End Function

' -- Helper: extend AVERAGE formula end-row by delta -------------------------
Private Function ExtendAverageFormula(formula As String, delta As Long) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "([A-Z]+)(\$?)(\d+)\)"

    Dim matches As Object
    Set matches = re.Execute(formula)

    Dim result As String
    result = formula
    Dim i As Long
    For i = matches.Count - 1 To 0 Step -1
        Dim m As Object
        Set m = matches(i)
        Dim oldRef  As String
        Dim newRef  As String
        Dim rowNum  As Long
        oldRef = m.Value
        rowNum = CLng(m.SubMatches(2)) + delta
        newRef = m.SubMatches(0) & m.SubMatches(1) & CStr(rowNum) & ")"
        result = Left(result, m.FirstIndex) & newRef & Mid(result, m.FirstIndex + Len(oldRef) + 1)
    Next i
    ExtendAverageFormula = result
End Function

' =============================================================================
' CELL TO DATE SERIAL HELPER
' Excel returns Date variant (not numeric) for date-formatted cells.
' VBA IsNumeric() returns False for Date variants -- this helper handles both.
' Returns the Excel date serial (>40000) or 0 if the cell isn't a date.
' =============================================================================
Private Function CellToDateSerial(v As Variant) As Long
    If IsError(v) Then CellToDateSerial = 0: Exit Function
    If TypeName(v) = "Date" Then
        CellToDateSerial = CLng(CDbl(v))
        Exit Function
    End If
    If IsNumeric(v) Then
        Dim d As Double
        d = CDbl(v)
        If d > 40000 And d < 80000 Then CellToDateSerial = CLng(d)
    End If
End Function

' -- Helper: column number to letter -----------------------------------------
Private Function ColNumToLetter(colNum As Long) As String
    Dim letter    As String
    Dim n         As Long
    Dim remainder As Long
    n = colNum
    Do While n > 0
        remainder = (n - 1) Mod 26
        letter = Chr(65 + remainder) & letter
        n = (n - 1) \ 26
    Loop
    ColNumToLetter = letter
End Function

' =============================================================================
' FORMAT P.1 DATE COLUMN
' Applies mmm-yy format to the month serial cells in the snapshot block
' =============================================================================
Private Sub FormatP1DateColumn(dashWb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = dashWb.Sheets(P1_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' Find first and last data rows (col B = month serial)
    Dim r       As Long
    Dim firstR  As Long
    Dim lastR   As Long
    firstR = 0
    lastR = 0
    For r = 15 To 200
        Dim v As Variant
        v = ws.Cells(r, P1_COL_MONTH).Value
        If CellToDateSerial(v) > 0 Then
            If firstR = 0 Then firstR = r
            lastR = r
        End If
    Next r

    If firstR = 0 Then Exit Sub

    ' Apply mmm-yy format so date values show as "Jan-26", "Feb-26", etc.
    ' (Cells already have Date type; format ensures consistent display.)
    ws.Range(ws.Cells(firstR, P1_COL_MONTH), ws.Cells(lastR, P1_COL_MONTH)).NumberFormat = "mmm-yy"

    ' Fix the column I header label -- remove the "$" sign
    Dim hdr As String
    hdr = CStr(ws.Cells(firstR - 1, P1_COL_MTG_HRS).Value)
    If InStr(hdr, "$") > 0 Then
        ws.Cells(firstR - 1, P1_COL_MTG_HRS).Value = "Meeting Hours Sold"
    End If

    Application.StatusBar = "P.1: date column formatted"
End Sub

' =============================================================================
' HIDE RAW DATA TABS
' Hides source export tabs that should not be visible to the finance team
' =============================================================================
Private Sub HideRawTabs(dashWb As Workbook)
    Dim tabList() As String
    tabList = Split(RAW_TABS, "|")
    Dim ws  As Worksheet
    Dim t   As Integer
    For t = 0 To UBound(tabList)
        On Error Resume Next
        Set ws = dashWb.Sheets(Trim(tabList(t)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVeryHidden
            Set ws = Nothing
        End If
    Next t

    ' Always ensure Yardi tabs are visible -- they contain live SUMIF formulas
    ' the finance team needs to inspect.  Explicitly un-hide in case a previous
    ' run left them VeryHidden before this policy change.
    Dim yardiNames(1) As String
    yardiNames(0) = "WA Yardi Inc Detail By Account"
    yardiNames(1) = "WA Yardi Adj Occupancy Report"
    Dim yi As Integer
    For yi = 0 To 1
        On Error Resume Next
        Set ws = dashWb.Sheets(yardiNames(yi))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVisible
            Set ws = Nothing
        End If
    Next yi

    Application.StatusBar = "Raw tabs hidden; Yardi tabs visible"
End Sub

' -- Show raw tabs (admin use) ------------------------------------------------
Public Sub ShowRawTabs()
    Dim tabList() As String
    tabList = Split(RAW_TABS, "|")
    Dim ws  As Worksheet
    Dim t   As Integer
    For t = 0 To UBound(tabList)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(Trim(tabList(t)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVisible
            Set ws = Nothing
        End If
    Next t
    MsgBox "Raw data tabs are now visible.", vbInformation, "Admin"
End Sub

' =============================================================================
' REWRITE INSTRUCTIONS SHEET
' Replaces the plain-text instruction list with a formatted SOP table
' =============================================================================
Public Sub RewriteInstructions(dashWb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = dashWb.Sheets(INSTR_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ws.Cells.Clear
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10

    ' -- Title ----------------------------------------------------------------
    With ws.Range("A1:F1")
        .Merge
        .Value = "25N COWORKING - MONTHLY DASHBOARD UPDATE  |  Standard Operating Procedure"
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(15, 32, 68)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With

    ws.Range("A2:F2").Interior.Color = RGB(201, 168, 76)
    ws.Range("A2:F2").RowHeight = 4

    ' -- Column headers -------------------------------------------------------
    ws.Range("A3").Value = "Step"
    ws.Range("B3").Value = "Tab"
    ws.Range("C3").Value = "Action"
    ws.Range("D3").Value = "Source"
    ws.Range("E3").Value = "Auto / Manual"
    ws.Range("F3").Value = "Notes"

    With ws.Range("A3:F3")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .HorizontalAlignment = xlCenter
        .RowHeight = 18
    End With

    ' -- Section label helper -------------------------------------------------
    ' Section headers use a gold background row (not in the main data loop)
    Dim secColor As Long
    secColor = RGB(201, 168, 76)

    ' -- SOP rows (split into two arrays to stay under VBA's 24-continuation limit)
    ' Format: Step | Tab | Action | Source | Auto/Manual | Notes
    Dim stepsA As Variant
    stepsA = Array( _
        Array("—", "PREPARATION", "BEFORE RUNNING: gather source files. Required: (1) Voyager 12-Month Statement, (2) Kube Occupancy Report, (3) Income Detail - Flex, (4) Income Detail - Mail/Mailbox, (5) Income Detail - Daily Space. Optional: (6) Kube Booking Summary.", "", "MANUAL", "All 5 required files are needed for a complete close. Booking Summary (6) is optional -- if omitted, P.1 Meeting Hours and P.2 meeting room rows must be entered manually."), _
        Array("1", "Web Tool", "Upload all source files to the web tool at report-generator.vercel.app. Download the Data Bridge .xlsx file.", "Voyager + Kube", "MANUAL", "Open the downloaded Data Bridge and confirm: (a) Meta sheet has correct month_code, (b) Status sheet shows all files detected. If a file is missing or shows 'Unknown', re-export from Voyager/Kube and re-upload."), _
        Array("2", "Master .xlsm", "Click Developer > Macros > MonthlyUpdate > Run. Select the Data Bridge file when prompted. Confirm the month shown in the dialog matches the reporting period, then click Yes.", "Data Bridge .xlsx", "AUTO", "The macro automatically completes steps 3-17. A validation report appears immediately after -- review all checks before proceeding to the manual steps."), _
        Array("—", "AUTOMATED BY MACRO (Steps 3-17)", "The following steps run automatically when MonthlyUpdate is executed.", "", "AUTO", ""), _
        Array("3", "P.1 Dashboard", "Header date in cell B2 updated to the first day of the reporting month.", "Data Bridge: Meta", "AUTO", ""), _
        Array("4", "P.1 Dashboard", "Sales total formula in cell H4 updated to reference the current month column in P.2.", "Data Bridge: Meta", "AUTO", ""), _
        Array("5", "P.1 Dashboard", "New snapshot row inserted above the Average row. Col B = date. Cols C/D/F/G written as live formulas. Col E (contracts) left blank -- enter manually. Col I (meeting hours) set automatically next month. Prior month frozen. Oldest row hidden. Average formulas extended.", "Data Bridge: Meta + Occupancy", "AUTO", "Enter contracts (col E) manually each month from your Customer List. Meeting hours (col I) for the prior month populate automatically when next month's update runs."), _
        Array("6", "P.1 Dashboard", "KPI sq ft cells (F10-G12): Dedicated Desk, Private Office, and Total sq ft computed from the Kube Occupancy Report.", "Kube Occupancy Report", "AUTO", ""), _
        Array("7", "2026 Sales By Month", "All 12-month revenue and cost-of-sales rows pasted from the Voyager 12-Month Statement. Matched by exact account label.", "Voyager 12-Month Statement", "AUTO", "If a row label in Voyager does not match the dashboard, it will be skipped. Add new labels manually and re-run if needed."), _
        Array("8", "P.2 Monthly Income", "Daily Space revenue for current and prior month written from Income Detail. Meeting room HOURS USED and REVENUE COLLECTED rows written for current month from Booking Summary.", "Income Detail + Booking Summary", "AUTO", "Verify P.2 Daily Space amount matches sum of Day Pass charges in Income Detail. Verify meeting room totals match Booking Summary. Check figure row must equal zero."), _
        Array("9", "P.3 Occupancy", "New month column added. Occupied spaces written with actual Contract Rate. Vacant spaces set to $0.", "Kube Occupancy Report", "AUTO", "Verify rates for new tenants match their signed agreement. New tenants starting mid-month may need a prorated rate entered manually."), _
        Array("10", "P.4 Forecast", "Column references updated so Actual column reflects current month.", "Internal (SBM tab)", "AUTO", "Verify: Gross Profit in P.4 col B = current month Gross Profit in 2026 Sales By Month tab.") _
    )

    Dim stepsB As Variant
    stepsB = Array( _
        Array("11", "Daily Space", "Year filter cell updated to current year. Existing rows cleared. Current-month day pass transaction rows pasted from Income Detail.", "Income Detail - Daily Space", "AUTO", "Verify check figure = 0. If an unrecognized charge code appears, it will show in the data but not match a check-figure row -- add it manually."), _
        Array("12", "Prior Month Mtg Hrs", "Tab cleared and refreshed with current-month booking rows from the Booking Summary.", "Kube Booking Summary", "AUTO", "Skipped if no Booking Summary was uploaded."), _
        Array("13", "WA Yardi Inc Detail By Account", "Data rows (row 9+) replaced with income rows from the uploaded Yardi Income Detail file. Header row and column structure are never touched. SUMIF formulas in WA Yardi Adj Occupancy Report recalculate automatically.", "Data Bridge: YardiIncDetailRaw", "AUTO", "Skipped if no Yardi Income Detail file was uploaded. Upload the Voyager Income Detail by Account export and regenerate the Data Bridge."), _
        Array("14", "WA Yardi Adj Occupancy Report", "Space data rows refreshed. Formula columns (L+) preserved. External refs and abbreviated tab name refs rewritten every run so right-side SUMIF summary resolves against live tabs.", "Data Bridge: YardiAdjOccRaw", "AUTO", "SUMIF summary (cols Y+) references 'WA Yardi Inc Detail By Account'. If values show $0 after first run, open the tab and press F9 to force recalc. BA97/BA98 feed P.3 occupancy summary."), _
        Array("15", "Hidden Tabs", "Background tabs hidden: Occupancy Data, Budget Actual Data. Yardi tabs remain visible for inspection.", "Internal", "AUTO", "Run ShowRawTabs macro to reveal Occupancy Data and Budget Actual Data for troubleshooting."), _
        Array("16", "Validation Report", "Post-update checks run automatically: P.1 snapshot present, no formula errors, P.3 column added, Occupancy Data populated, Daily Space rows written, P.2 Daily Space filled.", "Internal", "AUTO", "If any check shows '!' address it before distributing. All checks should show checkmarks for a clean close."), _
        Array("17", "Flex + Mail Members", "Income detail rows pasted into Flex and Mail Members tabs. Year filter updated on both.", "Data Bridge: FlexRaw / MailRaw", "AUTO", "If FlexRaw or MailRaw are missing from the Data Bridge, that tab will not update -- re-upload Income by Detail and regenerate."), _
        Array("—", "MANUAL STEPS (Steps 18-22)", "The following steps require manual input after the macro completes.", "", "MANUAL", ""), _
        Array("18", "Verify Check Figures", "Confirm check figure = 0 on: Flex tab, Mail Members tab, Daily Space tab. If non-zero, a row is miscategorized or missing -- investigate the Income by Detail export.", "Kube Income by Detail", "MANUAL", ""), _
        Array("19", "P.4 Forecast", "Update column F (Budget) with current-period budget figures from Voyager.", "Voyager Budget Report", "MANUAL", "Verify variance in col G matches the Voyager variance report."), _
        Array("20", "P.2 Monthly Income", "Verify the check figure row equals zero for all months. If non-zero, a formula is missing the new account -- locate and fix.", "Income Detail", "MANUAL", ""), _
        Array("21", "Membership tabs", "Lockers / Premier / Gym / Parking: ensure col K counts equal col H charges. Confirm rate totals match.", "", "MANUAL", "Update check figure formulas to reference current month's Sales By Month column if needed."), _
        Array("22", "Master .xlsm", "Save. Copy to Dashboard Backup Folder on the finance drive.", "", "MANUAL", "Naming convention: Waco Sales Dashboard [Month] [Year].xlsm") _
    )

    Dim autoColor   As Long
    Dim manualColor As Long
    autoColor   = RGB(226, 239, 218)
    manualColor = RGB(255, 242, 204)

    Dim dataRowCount As Long
    dataRowCount = 0

    Dim rowNum As Long
    rowNum = 3  ' header row is row 3; first data row will be 4

    Dim i As Integer
    Dim stp As Variant
    Dim allBatches(1) As Variant
    allBatches(0) = stepsA
    allBatches(1) = stepsB

    Dim b As Integer
    For b = 0 To 1
        Dim batch As Variant
        batch = allBatches(b)
        For i = 0 To UBound(batch)
            rowNum = rowNum + 1
            stp = batch(i)

            ' Section header rows (step = "—") use a gold banner style
            If CStr(stp(0)) = "—" Then
                ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6)).Merge
                ws.Cells(rowNum, 1).Value = CStr(stp(1)) & IIf(stp(5) <> "", "  |  " & stp(5), "")
                ws.Cells(rowNum, 1).Font.Bold = True
                ws.Cells(rowNum, 1).Font.Color = RGB(15, 32, 68)
                ws.Cells(rowNum, 1).Font.Size = 10
                ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6)).Interior.Color = secColor
                ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6)).Borders(xlEdgeBottom).Color = RGB(180, 140, 50)
                ws.Rows(rowNum).RowHeight = 20
            Else
                ws.Cells(rowNum, 1).Value = stp(0)
                ws.Cells(rowNum, 2).Value = stp(1)
                ws.Cells(rowNum, 3).Value = stp(2)
                ws.Cells(rowNum, 4).Value = stp(3)
                ws.Cells(rowNum, 5).Value = stp(4)
                ws.Cells(rowNum, 6).Value = stp(5)

                ws.Rows(rowNum).RowHeight = 36
                ws.Cells(rowNum, 3).WrapText = True
                ws.Cells(rowNum, 6).WrapText = True

                If dataRowCount Mod 2 = 0 Then
                    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6)).Interior.Color = RGB(245, 245, 245)
                Else
                    ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6)).Interior.Color = RGB(255, 255, 255)
                End If
                dataRowCount = dataRowCount + 1

                If stp(4) = "AUTO" Then
                    ws.Cells(rowNum, 5).Interior.Color = autoColor
                    ws.Cells(rowNum, 5).Font.Color = RGB(55, 86, 35)
                Else
                    ws.Cells(rowNum, 5).Interior.Color = manualColor
                    ws.Cells(rowNum, 5).Font.Color = RGB(124, 79, 0)
                End If
                ws.Cells(rowNum, 5).HorizontalAlignment = xlCenter
                ws.Cells(rowNum, 5).Font.Bold = True

                ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 6)).Borders(xlEdgeBottom).Color = RGB(210, 210, 210)
            End If
        Next i
    Next b

    ' -- Diagnostic tools section ---------------------------------------------
    Dim diagStartRow As Long
    diagStartRow = rowNum + 2

    With ws.Range(ws.Cells(diagStartRow, 1), ws.Cells(diagStartRow, 6))
        .Merge
        .Value = "DIAGNOSTIC MACROS  |  Run from Developer > Macros if issues occur"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .RowHeight = 18
    End With

    Dim diagTools As Variant
    diagTools = Array( _
        Array("DiagnoseP1Errors", "Lists every cell in P.1 Dashboard showing a formula error (#DIV/0!, #REF!, etc.) with its formula text. Run this first if P.1 KPIs look wrong."), _
        Array("DiagnoseSnapshotRows", "Shows the raw value, type, and date serial for every cell in the P.1 12-month snapshot block (col B, rows 55-80). Run this if duplicate months appear or the Average row looks wrong."), _
        Array("DiagnoseOccupancyData", "Shows column headers, first 5 rows, and sq ft aggregates from the Occupancy Data tab. Run this if P.1 KPI sqft cells (F10-G12) are showing 0 after a macro run."), _
        Array("CleanupP1DuplicateRows", "Scans the P.1 snapshot for duplicate month rows (same month inserted more than once) and deletes the extras. Run this after a test run that created duplicates."), _
        Array("ShowRawTabs", "Makes the hidden background tabs visible (Occupancy Data, Budget Actual Data). WA Yardi tabs are always visible and do not need this macro."), _
        Array("RunSetup", "One-time setup: hides raw tabs, rewrites this Instructions sheet, formats the P.1 date column. Run this after first importing the macro module.") _
    )

    Dim d As Integer
    For d = 0 To UBound(diagTools)
        Dim dRow As Long
        dRow = diagStartRow + 1 + d
        ws.Cells(dRow, 1).Value = diagTools(d)(0)
        ws.Cells(dRow, 1).Font.Bold = True
        ws.Cells(dRow, 1).Font.Color = RGB(31, 73, 125)
        ws.Range(ws.Cells(dRow, 2), ws.Cells(dRow, 6)).Merge
        ws.Cells(dRow, 2).Value = diagTools(d)(1)
        ws.Cells(dRow, 2).WrapText = True
        ws.Rows(dRow).RowHeight = 28
        If d Mod 2 = 0 Then
            ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 6)).Interior.Color = RGB(235, 241, 250)
        Else
            ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 6)).Interior.Color = RGB(255, 255, 255)
        End If
        ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Range(ws.Cells(dRow, 1), ws.Cells(dRow, 6)).Borders(xlEdgeBottom).Color = RGB(210, 210, 210)
    Next d

    ' -- Calculation Reference section ----------------------------------------
    ' One row per automated value: explains source, calculation logic, and how
    ' the finance team can independently reconcile each number.
    Dim calcStartRow As Long
    calcStartRow = diagStartRow + UBound(diagTools) + 2

    With ws.Range(ws.Cells(calcStartRow, 1), ws.Cells(calcStartRow, 6))
        .Merge
        .Value = "CALCULATION REFERENCE  |  How each automated value is derived and how to reconcile it"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .RowHeight = 18
    End With

    Dim calcHdrRow As Long
    calcHdrRow = calcStartRow + 1
    ws.Cells(calcHdrRow, 1).Value = "Tab"
    ws.Cells(calcHdrRow, 2).Value = "Cell / Range"
    ws.Cells(calcHdrRow, 3).Value = "Value Written"
    ws.Cells(calcHdrRow, 4).Value = "Source (Data Bridge)"
    ws.Cells(calcHdrRow, 5).Value = "How Calculated"
    ws.Cells(calcHdrRow, 6).Value = "Reconcile Against"
    With ws.Range(ws.Cells(calcHdrRow, 1), ws.Cells(calcHdrRow, 6))
        .Font.Bold = True
        .Font.Color = RGB(15, 32, 68)
        .Interior.Color = RGB(201, 168, 76)
        .RowHeight = 16
    End With

    ' Format: Tab | Cell | Value Written | Source (Data Bridge) | How Calculated | Reconcile Against
    Dim calcRows As Variant
    calcRows = Array( _
        Array("P.1 Dashboard", "B2", "Date — 1st day of current month (e.g., 5/1/2026)", "Meta sheet  row: month_serial", "month_serial (Excel date serial) is converted to a date and written as a date value.", "Must match the first day of the month you are closing. If wrong, check that the correct source files were uploaded."), _
        Array("P.1 Dashboard", "H4", "Formula reference to current month column in P.2", "Meta sheet  row: month_code", "Macro finds the column letter for the current month in P.2 row 2, then rewrites H4 to reference that column (e.g., P2!G7 becomes P2!H7).", "H4 should equal the Gross Revenue cell in P.2 for the current month. If H4 is wrong, check that P.2 has a column header matching the current month name."), _
        Array("P.1 Dashboard", "Col B — new snapshot row", "Text month label (e.g., May-26)", "Meta sheet  row: month_code", "Taken directly from month_code. If Excel auto-converted month_code to a date serial on open, macro formats it back to mmm-yy text before writing.", "Should match the column header in P.2 for the current month. If mismatched, regenerate the Data Bridge and re-run."), _
        Array("P.1 Dashboard", "Col C — new snapshot row", "Occupancy % by spaces", "Meta sheet  row: occupancy_pct_spaces", "Occupied spaces ÷ total spaces from the Kube Occupancy Report. Occupied = rows with a non-blank space code and Status = Occupied.", "Should match the occupancy % (spaces) shown in the Kube Occupancy Report for the reporting month."), _
        Array("P.1 Dashboard", "Col D — new snapshot row", "Occupancy % by sq ft", "Meta sheet  row: occupancy_pct_sqft", "Sum of Sq Ft for Occupied rows ÷ sum of Sq Ft for all rows in the Kube report.", "Should match the sq ft occupancy % in the Kube Occupancy Report. Discrepancy = sq ft mismatch in Kube setup."), _
        Array("P.1 Dashboard", "Col E — new snapshot row", "LEFT BLANK — enter manually", "N/A", "Not written by the macro. Enter manually from your Customer List each month. The occupancy file reports 35 unique company names for May 2026, but the historical definition may differ — use your Customer List as the authority.", "Cross-check against your Customer List. Comped, trial, or multi-location members may require a ± adjustment from the raw occupancy count."), _
        Array("P.1 Dashboard", "Col F — new snapshot row", "Flex member count", "Meta sheet  row: flex_members", "Count of unique company names in the Income Detail file with a Flex charge code in the current Post Month.", "Should equal the number of data rows in the Flex tab after the macro runs. If count differs, check the Flex tab for duplicate invoices."), _
        Array("P.1 Dashboard", "Col G — new snapshot row", "Mail member count", "Meta sheet  row: mail_members", "Count of unique company names in the Income Detail file with a Mail / Mailbox / Virtual Office / Virtual Receptionist charge code in the current Post Month.", "Should equal the number of data rows in the Mail Members tab. If count differs, check the Mail Members tab for duplicate invoices."), _
        Array("P.1 Dashboard", "Col I — new snapshot row", "LEFT BLANK — auto-set by next month's run", "Prior Month Mtg Hrs tab  col Q3", "I is left blank when the new row is first inserted. The FOLLOWING month's macro run sets I = ='Prior Month Meeting Hours'!Q3 (sum of the CALCULATED HOURS column for the prior month's booking rows). The month before that gets its I frozen to a locked value.", "After the next month's macro runs, verify I equals the sum of the CALCULATED HOURS column on the Prior Month Mtg Hrs tab. If the tab was skipped (no Booking Summary uploaded), enter I manually from the Kube Booking Summary."), _
        Array("P.1 Dashboard", "F10, F11, G12 (sq ft KPIs)", "Dedicated Desk sq ft / Private Office sq ft / Total sq ft", "Occupancy sheet  Sq Ft column", "Sum of Sq Ft for rows where Space Type contains Dedicated Desk (F10) or Private Office (F11). G12 = total sq ft across all rows.", "Should match your floor plan or Kube property setup. If wrong, check Kube space type labels."), _
        Array("2026 Sales By Month", "All revenue + COS rows (all 12 months)", "12-month P&L by account label (Jan–Dec)", "SBM sheet  all month columns", "Each dashboard row is matched to a Voyager account row by label (exact match first, then fuzzy). All 12 month values pasted. Unmatched labels are skipped.", "Compare the SBM tab row totals to the Voyager 12-Month Statement. Open the Data Bridge Status sheet to see which labels were skipped."), _
        Array("P.2 Monthly Income", "Daily Space row — current month col", "Day pass / daily space revenue ($)", "Meta sheet  row: daily_space_current", "Sum of the Income column for all Income Detail rows with a Day Pass / Daily Space / Day Office charge code in the current Post Month.", "Must equal the sum of the Income column in the Daily Space tab for current-month rows. Check figure in the Daily Space tab must be zero."), _
        Array("P.2 Monthly Income", "Daily Space row — prior month col", "Day pass / daily space revenue ($) — prior month", "Meta sheet  row: daily_space_prior", "Same as above, filtered to the prior Post Month (e.g., April when closing May).", "Should match the Daily Space revenue recorded in the prior month's close. If prior month was entered manually, verify this overwrite is correct."), _
        Array("P.2 Monthly Income", "HOURS USED section — current month col", "Hours per meeting room", "MtgRoomHours sheet  col C (Hours)", "Sum of Total_Units_Charged grouped by room name, confirmed bookings only, current month. Room names are mapped to P.2 row labels (e.g., Hub/Terrace, Boardroom).", "Filter the Kube Booking Summary to current month + Confirmed status. Sum hours for each room and compare row by row to the HOURS USED section."), _
        Array("P.2 Monthly Income", "REVENUE COLLECTED section — current month col", "Revenue per meeting room ($)", "MtgRoomHours sheet  col D (Revenue)", "Sum of the revenue column (Sub Total, Revenue, Billable Amount, or similar depending on Kube export format) grouped by room name, current month confirmed bookings.", "Filter Booking Summary to current month + Confirmed. Sum revenue for each room. Grand total must match the REVENUE COLLECTED total row in P.2."), _
        Array("P.3 Occupancy", "Current month column — all space rows", "Contract Rate ($) per space", "Occupancy sheet  Contract Rate column", "Occupied spaces: actual Contract Rate from the Kube Occupancy Report. Vacant spaces: $0.", "Compare each occupied space rate to the signed lease agreement. New tenants starting mid-month may need a prorated rate entered manually."), _
        Array("P.4 Forecast", "Actual column — current month", "Current month actuals by account row", "2026 Sales By Month tab (formula reference, not pasted values)", "P.4 contains formulas that reference the SBM tab by column letter. Macro updates the column letter to point to the current month column (e.g., G  H when advancing one month).", "P.4 Actual grand total must equal the Grand Total in the current month column of 2026 Sales By Month. Verify by checking one row manually."), _
        Array("Daily Space tab", "All data rows", "Transaction-level day pass invoices for current month", "DailySpaceRaw sheet  all columns", "Direct copy of all Income Detail rows matching Day Pass / Daily Space / Day Office charge codes in the current Post Month. Deduplicated by Invoice #.", "Check figure row must equal zero. Sum of Income column must equal the P.2 Daily Space cell for the current month."), _
        Array("Flex tab", "All data rows", "Transaction-level Flex membership invoices for current month", "FlexRaw sheet  all columns", "Direct copy of all Income Detail rows matching Flex charge codes in the current Post Month. Deduplicated by Invoice #.", "Check figure must equal zero. Row count must match the Flex member count in P.1 col F (new snapshot row)."), _
        Array("Mail Members tab", "All data rows", "Transaction-level Mail membership invoices for current month", "MailRaw sheet  all columns", "Direct copy of all Income Detail rows matching Mail / Mailbox / Virtual Office / Virtual Receptionist charge codes in the current Post Month. Deduplicated by Invoice #.", "Check figure must equal zero. Row count must match the Mail member count in P.1 col G (new snapshot row)."), _
        Array("Prior Month Mtg Hrs", "All data rows", "Individual confirmed booking rows from the prior month", "BookingSummaryPriorRaw sheet  all columns", "All non-cancelled, non-25N-Staff booking rows with Check In in the prior month, copied verbatim from the Kube Booking Summary.", "Sum of the Hours column should equal the Meeting Hours value in the prior month P.1 snapshot row (col I, one row below the newly inserted row).") _
    )

    Dim cr As Integer
    For cr = 0 To UBound(calcRows)
        Dim crRow As Long
        crRow = calcHdrRow + 1 + cr
        Dim crData As Variant
        crData = calcRows(cr)
        ws.Cells(crRow, 1).Value = crData(0)
        ws.Cells(crRow, 2).Value = crData(1)
        ws.Cells(crRow, 3).Value = crData(2)
        ws.Cells(crRow, 4).Value = crData(3)
        ws.Cells(crRow, 5).Value = crData(4)
        ws.Cells(crRow, 6).Value = crData(5)
        ws.Rows(crRow).RowHeight = 44
        ws.Cells(crRow, 2).WrapText = True
        ws.Cells(crRow, 3).WrapText = True
        ws.Cells(crRow, 5).WrapText = True
        ws.Cells(crRow, 6).WrapText = True
        If cr Mod 2 = 0 Then
            ws.Range(ws.Cells(crRow, 1), ws.Cells(crRow, 6)).Interior.Color = RGB(245, 245, 245)
        Else
            ws.Range(ws.Cells(crRow, 1), ws.Cells(crRow, 6)).Interior.Color = RGB(255, 255, 255)
        End If
        ws.Range(ws.Cells(crRow, 1), ws.Cells(crRow, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Range(ws.Cells(crRow, 1), ws.Cells(crRow, 6)).Borders(xlEdgeBottom).Color = RGB(210, 210, 210)
    Next cr

    ' -- Version footer -------------------------------------------------------
    Dim footRow As Long
    footRow = calcHdrRow + UBound(calcRows) + 3
    ws.Cells(footRow, 1).Value = "Last updated: " & Format(Now, "mmmm yyyy") & "  |  Macro version: MonthlyUpdate v3  |  Maintained by finance team + macro administrator"
    ws.Cells(footRow, 1).Font.Italic = True
    ws.Cells(footRow, 1).Font.Color = RGB(130, 130, 130)
    ws.Cells(footRow, 1).Font.Size = 9

    ' -- Column widths --------------------------------------------------------
    ' A=16: fits "P.1 Dashboard" in Calc Ref; step numbers still fine in SOP.
    ' E=34: fits "How Calculated" sentences; "AUTO/MANUAL" badge still centered.
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 38
    ws.Columns("D").ColumnWidth = 24
    ws.Columns("E").ColumnWidth = 34
    ws.Columns("F").ColumnWidth = 44

    ' -- Freeze top rows ------------------------------------------------------
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(4, 1).Select
    ActiveWindow.FreezePanes = True

    Application.StatusBar = "Instructions tab rewritten"
End Sub

' =============================================================================
' ADD P.3 MONTH COLUMN
' P3. Occupancy is a time-series grid: rows = space codes, cols = months.
' For the new month: occupied spaces carry forward their prior-month rate;
' vacant spaces get 0.  If the month column already exists, skips safely.
' Data Bridge "Occupancy" sheet supplies current-month status per space code.
' =============================================================================
Private Sub AddP3MonthColumn(dashWb As Workbook, bridgeWb As Workbook, monthCode As String)
    Dim p3Step As String
    p3Step = "get sheets"
    On Error GoTo P3ErrorHandler

    Dim p3Ws  As Worksheet
    Dim occWs As Worksheet

    On Error Resume Next
    Set p3Ws  = dashWb.Sheets(P3_SHEET)
    Set occWs = bridgeWb.Sheets("Occupancy")
    On Error GoTo P3ErrorHandler

    If occWs Is Nothing Then
        Application.StatusBar = "P.3: no Occupancy sheet in Data Bridge -- skipping"
        Exit Sub
    End If
    If p3Ws Is Nothing Then
        MsgBox "Sheet '" & P3_SHEET & "' not found.", vbExclamation, "Monthly Update"
        Exit Sub
    End If

    ' P.3 header row is always row 3; month labels use "Mon-YYYY" format
    p3Step = "compute monthLabel / lastCol"
    Dim headerRow As Long
    headerRow = 3
    Dim monthLabel As String
    monthLabel = ExpandMonthCode(monthCode)  ' "May-26" -> "May-2026"

    Dim lastCol As Long
    lastCol = p3Ws.Cells(headerRow, p3Ws.Columns.Count).End(xlToLeft).Column

    ' Find existing column for this month, or use next available column.
    ' When a column already exists (e.g. from a prior hydration run with broken formulas)
    ' we still overwrite the data values -- writing .Value is always safe.
    p3Step = "find newCol / priorCol"
    Dim c      As Long
    Dim newCol As Long
    Dim priorCol As Long  ' column to read prior-month rate from (always the col before newCol)
    newCol = 0
    For c = 3 To lastCol
        Dim cellV As Variant
        cellV = p3Ws.Cells(headerRow, c).Value
        Dim cellStr2 As String
        cellStr2 = ""
        ' Headers may be stored as text ("May-2026"), numeric date serial (46143),
        ' or VBA Date type -- all must be normalised to "mmm-yyyy" for comparison.
        ' NOTE: VBA And/Or are NOT short-circuit -- nested If required.
        ' NOTE: IsNumeric returns False for VBA Date variants -- check TypeName separately.
        If IsError(cellV) Then
            cellStr2 = ""
        ElseIf TypeName(cellV) = "Date" Then
            On Error Resume Next
            cellStr2 = Format(cellV, "mmm-yyyy")
            On Error GoTo 0
        ElseIf IsNumeric(cellV) Then
            Dim cellDbl As Double
            cellDbl = CDbl(cellV)
            If cellDbl > 40000 Then
                On Error Resume Next
                cellStr2 = Format(CDate(cellDbl), "mmm-yyyy")
                On Error GoTo 0
            Else
                cellStr2 = LCase(Trim(CStr(cellV)))
            End If
        Else
            cellStr2 = LCase(Trim(CStr(cellV)))
        End If
        If LCase(cellStr2) = LCase(monthLabel) Then
            newCol = c
            Exit For
        End If
    Next c

    If newCol = 0 Then
        ' Brand-new column: prior month is lastCol
        newCol    = lastCol + 1
        priorCol  = lastCol
        p3Ws.Cells(headerRow, newCol).Value = monthLabel
    Else
        ' Existing column: prior month is one to the left
        priorCol = newCol - 1
        If priorCol < 3 Then priorCol = 3  ' safety floor -- never go before col C
        Application.StatusBar = "P.3: " & monthLabel & " column exists -- overwriting values"
    End If

    ' Detect column positions from Data Bridge Occupancy header row (row 1).
    ' Defaults match the standard Kube Occupancy Report column layout.
    p3Step = "detect bridge columns"
    Dim spCodeCol  As Long   ' Space # column
    Dim spStatCol  As Long   ' Status column
    Dim rateCol    As Long   ' Contract Rate column (0 = not found)
    spCodeCol = 3: spStatCol = 4: rateCol = 0
    Dim bridgeLastCol As Long
    bridgeLastCol = occWs.Cells(1, occWs.Columns.Count).End(xlToLeft).Column
    Dim bc As Long
    For bc = 1 To bridgeLastCol
        Dim bh As String
        bh = LCase(Trim(CStr(occWs.Cells(1, bc).Value)))
        If bh = "space #" Or bh = "space#" Or bh = "space code" Then spCodeCol = bc
        If bh = "status" Then spStatCol = bc
        ' "Contract Rate" contains both "contract" and "rate"; "Market Rate" does not
        If InStr(bh, "contract") > 0 And InStr(bh, "rate") > 0 Then rateCol = bc
    Next bc

    ' Build lookups: space code -> status and space code -> contract rate
    p3Step = "build spaceStatus / spaceRate"
    Dim spaceStatus As Object
    Dim spaceRate   As Object
    Set spaceStatus = CreateObject("Scripting.Dictionary")
    Set spaceRate   = CreateObject("Scripting.Dictionary")
    Dim occLastRow As Long
    occLastRow = occWs.Cells(occWs.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To occLastRow
        Dim spCode As String
        Dim spStat As String
        spCode = Trim(CStr(occWs.Cells(r, spCodeCol).Value))
        spStat = LCase(Trim(CStr(occWs.Cells(r, spStatCol).Value)))
        If spCode <> "" Then
            spaceStatus(spCode) = spStat
            If rateCol > 0 Then
                Dim rateVal As Double
                rateVal = 0
                On Error Resume Next
                rateVal = CDbl(occWs.Cells(r, rateCol).Value)
                On Error GoTo 0
                If rateVal > 0 Then spaceRate(spCode) = rateVal
            End If
        End If
    Next r

    ' Walk P3 data rows; fill new column based on occupancy status and contract rate
    p3Step = "walk p3 rows"
    Dim lastDataRow As Long
    lastDataRow = p3Ws.Cells(p3Ws.Rows.Count, 1).End(xlUp).Row
    Dim updated As Long
    updated = 0

    For r = headerRow + 1 To lastDataRow
        p3Step = "walk p3 rows r=" & r
        Dim spaceCode As String
        spaceCode = Trim(CStr(p3Ws.Cells(r, 1).Value))
        ' Skip blank rows
        If spaceCode = "" Then GoTo NextSpace

        ' Special Yardi-sourced summary rows (Add Coworkers / Discounts).
        ' Protection rules -- applied in order:
        '  1. If this cell already has a formula (user set it), never overwrite.
        '  2. If the prior month column has a formula for this row, copy it exactly
        '     (so user's one-time fix propagates to every future month automatically).
        '  3. Fallback: write the current known correct reference from the MASTER.
        If LCase(spaceCode) = "add coworkers" Or LCase(spaceCode) = "discounts" Then
            ' Rule 1: already set — keep it
            If p3Ws.Cells(r, newCol).HasFormula Then GoTo NextSpace
            ' Rule 2: propagate from the immediately prior month column
            If newCol > 1 Then
                Dim ycPrior As Range
                Set ycPrior = p3Ws.Cells(r, newCol - 1)
                If ycPrior.HasFormula Then
                    p3Ws.Cells(r, newCol).Formula = ycPrior.Formula
                    updated = updated + 1
                    GoTo NextSpace
                End If
            End If
            ' Rule 3: fallback formulas referencing current MASTER Adj Occ Report cells
            If LCase(spaceCode) = "add coworkers" Then
                p3Ws.Cells(r, newCol).Formula = "='WA Yardi Adj Occupancy Report'!AB10"
            Else
                p3Ws.Cells(r, newCol).Formula = "='WA Yardi Adj Occupancy Report'!AB11+'WA Yardi Adj Occupancy Report'!AB9"
            End If
            updated = updated + 1
            GoTo NextSpace
        End If

        ' Total row: write column SUM formula matching existing master pattern (=SUM({col}4:{col}101))
        If LCase(spaceCode) = "total" Then
            Dim colLet As String
            colLet = Split(p3Ws.Cells(1, newCol).Address(True, False), "$")(0)
            p3Ws.Cells(r, newCol).Formula = "=SUM(" & colLet & "4:" & colLet & "101)"
            GoTo NextSpace
        End If

        If spaceStatus.Exists(spaceCode) Then
            Dim isOcc As Boolean
            isOcc = (InStr(spaceStatus(spaceCode), "occup") > 0)
            If isOcc Then
                ' Prefer the current-month Contract Rate from the Data Bridge (exact rate
                ' for this tenant this month -- handles new tenants and rate changes).
                ' Fall back to prior-month carry-forward only when no rate is available.
                If spaceRate.Exists(spaceCode) And spaceRate(spaceCode) > 0 Then
                    p3Ws.Cells(r, newCol).Value = spaceRate(spaceCode)
                Else
                    Dim priorVal As Variant
                    priorVal = p3Ws.Cells(r, priorCol).Value
                    If IsNumeric(priorVal) And CDbl(priorVal) > 0 Then
                        p3Ws.Cells(r, newCol).Value = CDbl(priorVal)
                    End If
                    ' If no prior value and no rate (genuinely new space), leave blank
                    ' so finance team is prompted to enter the rate manually
                End If
            Else
                ' Vacant -- write 0
                p3Ws.Cells(r, newCol).Value = 0
            End If
            updated = updated + 1
        End If
NextSpace:
    Next r

    Application.StatusBar = "P.3: wrote " & monthLabel & " column (" & updated & " spaces)"
    Exit Sub

P3ErrorHandler:
    ' Re-raise with step info so MonthlyUpdate's handler shows the exact location
    Dim errNum As Long
    Dim errDesc As String
    errNum  = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    Err.Raise errNum, "[AddP3MonthColumn:" & p3Step & "]", errDesc
End Sub

' -- Helper: expand two-digit year in month code ("May-26" -> "May-2026") ----
Private Function ExpandMonthCode(monthCode As String) As String
    ' VBA's And is NOT short-circuit -- both sides always evaluate.
    ' Use nested If so parts(1) is only accessed when UBound confirms it exists.
    Dim parts() As String
    parts = Split(monthCode, "-")
    If UBound(parts) = 1 Then
        If Len(parts(1)) = 2 Then
            ExpandMonthCode = parts(0) & "-20" & parts(1)
            Exit Function
        End If
    End If
    ExpandMonthCode = monthCode
End Function

' =============================================================================
' UPDATE P.4 COLUMN REFERENCES
' =============================================================================
Private Sub UpdateP4ColumnRefs(dashWb As Workbook, monthSerial As Long)
    Dim p4Ws  As Worksheet
    Dim sbmWs As Worksheet

    On Error Resume Next
    Set p4Ws  = dashWb.Sheets(P4_SHEET)
    Set sbmWs = dashWb.Sheets(SBM_SHEET)
    On Error GoTo 0

    If p4Ws Is Nothing Or sbmWs Is Nothing Then Exit Sub

    Dim sbmHeaderRow As Long
    sbmHeaderRow = FindSBMHeaderRow(sbmWs)
    If sbmHeaderRow = 0 Then Exit Sub

    Dim newCol As String
    newCol = ""
    Dim c As Long
    For c = 1 To sbmWs.Cells(sbmHeaderRow, sbmWs.Columns.Count).End(xlToLeft).Column
        Dim v As Variant
        v = sbmWs.Cells(sbmHeaderRow, c).Value
        If IsNumeric(v) Then
            If CLng(v) = monthSerial Then
                newCol = ColNumToLetter(c)
                Exit For
            End If
        End If
    Next c

    If newCol = "" Then
        Application.StatusBar = "P.4 column refs: month column not found in SBM -- skipping"
        Exit Sub
    End If

    Dim sbmRef    As String
    sbmRef = "'" & SBM_SHEET & "'!"
    Dim colCounts As Object
    Set colCounts = CreateObject("Scripting.Dictionary")

    Dim cell As Range
    For Each cell In p4Ws.UsedRange
        If cell.HasFormula Then
            Dim f As String
            f = cell.Formula
            Dim pos As Long
            pos = InStr(f, sbmRef)
            Do While pos > 0
                Dim afterRef As String
                afterRef = Mid(f, pos + Len(sbmRef))
                If Left(afterRef, 1) = "$" Then afterRef = Mid(afterRef, 2)
                Dim colLetters As String
                colLetters = ""
                Dim i As Long
                For i = 1 To Len(afterRef)
                    Dim ch As String
                    ch = Mid(afterRef, i, 1)
                    If ch >= "A" And ch <= "Z" Then
                        colLetters = colLetters & ch
                    Else
                        Exit For
                    End If
                Next i
                If colLetters <> "" Then
                    If colCounts.Exists(colLetters) Then
                        colCounts(colLetters) = colCounts(colLetters) + 1
                    Else
                        colCounts(colLetters) = 1
                    End If
                End If
                pos = InStr(pos + 1, f, sbmRef)
            Loop
        End If
    Next cell

    Dim oldCol   As String
    Dim maxCount As Long
    oldCol = ""
    maxCount = 0
    Dim key As Variant
    For Each key In colCounts.Keys
        If colCounts(key) > maxCount Then
            maxCount = colCounts(key)
            oldCol = key
        End If
    Next key

    If oldCol = "" Or oldCol = newCol Then
        Application.StatusBar = "P.4 column refs: already current (" & newCol & ")"
        Exit Sub
    End If

    Dim updatedCount As Long
    updatedCount = 0
    For Each cell In p4Ws.UsedRange
        If cell.HasFormula Then
            Dim oldFm2 As String
            oldFm2 = cell.Formula
            Dim newFm2 As String
            newFm2 = ReplaceColumnInFormula(oldFm2, SBM_SHEET, oldCol, newCol)
            If newFm2 <> oldFm2 Then
                cell.Formula = newFm2
                updatedCount = updatedCount + 1
            End If
        End If
    Next cell

    Application.StatusBar = "P.4: updated " & updatedCount & " formula(s) -- " & oldCol & " -> " & newCol
End Sub

' -- Helper: replace SBM column reference in a formula string -----------------
Private Function ReplaceColumnInFormula(formula As String, _
                                        sheetName As String, _
                                        oldCol As String, _
                                        newCol As String) As String
    Dim result As String
    result = formula
    Dim variants(1) As String
    variants(0) = "'" & sheetName & "'!" & oldCol
    variants(1) = sheetName & "!" & oldCol

    Dim v As Integer
    For v = 0 To 1
        Dim searchStr  As String
        Dim replaceStr As String
        searchStr  = variants(v)
        replaceStr = Replace(searchStr, "!" & oldCol, "!" & newCol)
        Dim pos As Long
        pos = 1
        Do
            pos = InStr(pos, result, searchStr)
            If pos = 0 Then Exit Do
            Dim nextChar As String
            nextChar = ""
            If pos + Len(searchStr) <= Len(result) Then
                nextChar = Mid(result, pos + Len(searchStr), 1)
            End If
            If nextChar = "" Or (nextChar >= "0" And nextChar <= "9") Or nextChar = "$" Then
                result = Left(result, pos - 1) & replaceStr & Mid(result, pos + Len(searchStr))
                pos = pos + Len(replaceStr)
            Else
                pos = pos + Len(searchStr)
            End If
        Loop
    Next v

    ReplaceColumnInFormula = result
End Function

' =============================================================================
' UPDATE P.1 H4 FORMULA
' P.1 cell H4 contains =+'P2. Monthly Income'!G15 (or similar column letter).
' This formula must advance each month so it references the column in P.2 that
' holds the current month's total revenue.
' P.2 has a fixed Jan-Dec header layout (row 2, full month names).
' This sub finds the current month's column in P.2 and rewrites H4's formula
' using a RegExp replacement -- independent of whether P.4 has SBM references.
' =============================================================================
Private Sub UpdateP1H4Formula(dashWb As Workbook, monthSerial As Long)
    Dim p1Ws As Worksheet
    Dim p2Ws As Worksheet

    On Error Resume Next
    Set p1Ws = dashWb.Sheets(P1_SHEET)
    Set p2Ws = dashWb.Sheets("P2. Monthly Income")
    On Error GoTo 0

    If p1Ws Is Nothing Or p2Ws Is Nothing Then
        Application.StatusBar = "P.1 H4: P1 or P2 sheet not found -- skipping"
        Exit Sub
    End If

    ' Find the P.2 column for the current month by scanning the header row
    Dim curMonthName As String
    On Error Resume Next
    curMonthName = Format(CDate(monthSerial), "mmmm")
    On Error GoTo 0
    If curMonthName = "" Then
        Application.StatusBar = "P.1 H4: cannot parse month from serial " & monthSerial
        Exit Sub
    End If

    Dim p2Col As Long
    p2Col = 0
    Dim p2c As Long
    Dim p2LastCol As Long
    p2LastCol = p2Ws.Cells(2, p2Ws.Columns.Count).End(xlToLeft).Column
    For p2c = 1 To p2LastCol
        If LCase(Trim(CStr(p2Ws.Cells(2, p2c).Value))) = LCase(curMonthName) Then
            p2Col = p2c
            Exit For
        End If
    Next p2c

    If p2Col = 0 Then
        Application.StatusBar = "P.1 H4: '" & curMonthName & "' not found in P.2 header row -- skipping"
        Exit Sub
    End If

    Dim newColLetter As String
    newColLetter = ColNumToLetter(p2Col)

    ' Read H4 and rewrite with RegExp -- replace the column letter in any P2 reference
    Const H4_ROW As Long = 4
    Const H4_COL As Long = 8   ' col H

    If Not p1Ws.Cells(H4_ROW, H4_COL).HasFormula Then
        Application.StatusBar = "P.1 H4: cell H4 has no formula -- nothing to update"
        Exit Sub
    End If

    Dim h4Fm  As String
    h4Fm = p1Ws.Cells(H4_ROW, H4_COL).Formula

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    ' Match: 'P2. Monthly Income'!<COL><ROW>  or  P2. Monthly Income!<COL><ROW>
    re.Pattern    = "((?:'P2\. Monthly Income'|P2\. Monthly Income)!\$?)([A-Z]+)(\d+)"
    re.IgnoreCase = True
    re.Global     = True

    If Not re.Test(h4Fm) Then
        Application.StatusBar = "P.1 H4: no P2 reference found in H4 formula -- skipping"
        Exit Sub
    End If

    Dim matches As Object
    Set matches = re.Execute(h4Fm)
    Dim h4New As String
    h4New = h4Fm
    Dim mi As Long
    ' Replace matches in reverse order to preserve character positions
    For mi = matches.Count - 1 To 0 Step -1
        Dim m As Object
        Set m = matches(mi)
        Dim replacement As String
        replacement = m.SubMatches(0) & newColLetter & m.SubMatches(2)
        h4New = Left(h4New, m.FirstIndex) & replacement & Mid(h4New, m.FirstIndex + m.Length + 1)
    Next mi

    If h4New <> h4Fm Then
        On Error Resume Next
        p1Ws.Cells(H4_ROW, H4_COL).Formula = h4New
        On Error GoTo 0
        Application.StatusBar = "P.1 H4: updated to reference " & curMonthName & " (col " & newColLetter & ")"
    Else
        Application.StatusBar = "P.1 H4: already references " & curMonthName & " -- no change"
    End If
End Sub

' =============================================================================
' DIAGNOSE OCCUPANCY DATA
' Shows column headers and first 5 data rows from the Occupancy Data tab.
' Run this when P.1 KPI sqft cells (F10-G12) are returning 0 to verify
' that column detection is working correctly.
' =============================================================================
Public Sub DiagnoseOccupancyData()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Occupancy Data")
    If Not ws Is Nothing Then ws.Visible = xlSheetVisible
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Occupancy Data tab not found. Run MonthlyUpdate first.", vbExclamation, "Diagnose"
        Exit Sub
    End If

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim report As String
    report = "Occupancy Data: " & (lastRow - 1) & " rows, " & lastCol & " cols" & Chr(13) & _
             String(55, "-") & Chr(13) & "HEADERS:" & Chr(13)

    ' Show headers with col letters
    Dim c As Long
    For c = 1 To lastCol
        Dim hdr As String
        hdr = CStr(ws.Cells(1, c).Value)
        report = report & "  " & ColNumToLetter(c) & " (" & c & "): " & hdr & Chr(13)
    Next c

    ' Show first 5 data rows
    report = report & Chr(13) & "FIRST 5 ROWS (cols 1-8):" & Chr(13)
    Dim r As Long
    Dim maxRow As Long
    maxRow = Application.WorksheetFunction.Min(lastRow, 6)
    For r = 2 To maxRow
        Dim rowStr As String
        rowStr = "R" & r & ": "
        Dim maxC As Long
        maxC = Application.WorksheetFunction.Min(lastCol, 8)
        For c = 1 To maxC
            Dim v As Variant
            v = ws.Cells(r, c).Value
            rowStr = rowStr & "|" & TypeName(v) & ":" & CStr(v)
        Next c
        report = report & rowStr & Chr(13)
    Next r

    ' Detect which col has sqft
    Dim sqftColFound As Long
    sqftColFound = 0
    For c = 1 To lastCol
        Dim h As String
        h = LCase(Trim(CStr(ws.Cells(1, c).Value)))
        If InStr(h, " sq") > 0 Or h = "sqft" Or h = "sq ft" Or InStr(h, "square") > 0 Then
            sqftColFound = c
            Exit For
        End If
    Next c
    If sqftColFound > 0 Then
        report = report & Chr(13) & "Sq Ft col detected: " & ColNumToLetter(sqftColFound) & " (" & sqftColFound & ")"
        ' Show type distribution
        Dim typeColFound As Long
        typeColFound = 0
        For c = 1 To lastCol
            Dim th As String
            th = LCase(Trim(CStr(ws.Cells(1, c).Value)))
            If InStr(th, "type") > 0 And typeColFound = 0 Then typeColFound = c
        Next c
        If typeColFound > 0 Then
            report = report & Chr(13) & "Type col: " & ColNumToLetter(typeColFound) & " (" & typeColFound & ")"
            Dim ddSqFt As Double, poSqFt As Double, otherSqFt As Double
            For r = 2 To lastRow
                Dim sqftV As Variant
                sqftV = ws.Cells(r, sqftColFound).Value
                Dim sqftN As Double
                sqftN = 0
                On Error Resume Next
                sqftN = CDbl(CStr(sqftV))
                On Error GoTo 0
                Dim typeV As String
                typeV = LCase(Trim(CStr(ws.Cells(r, typeColFound).Value)))
                If InStr(typeV, "dedicated") > 0 Or InStr(typeV, "desk") > 0 Then
                    ddSqFt = ddSqFt + sqftN
                ElseIf InStr(typeV, "private") > 0 Or InStr(typeV, "office") > 0 Then
                    poSqFt = poSqFt + sqftN
                Else
                    otherSqFt = otherSqFt + sqftN
                End If
            Next r
            report = report & Chr(13) & "  Dedicated Desk sqft: " & ddSqFt
            report = report & Chr(13) & "  Private Office sqft: " & poSqFt
            report = report & Chr(13) & "  Other sqft:          " & otherSqFt
        End If
    Else
        report = report & Chr(13) & "WARNING: Sq Ft column NOT found -- check header names"
    End If

    ws.Visible = xlSheetVeryHidden
    MsgBox report, vbInformation, "Occupancy Data Diagnosis"
End Sub

' =============================================================================
' UPDATE P.2 MEETING ROOM HOURS & REVENUE
' Reads the MtgRoomHours sheet from the Data Bridge and writes hours and
' revenue values into P.2 Monthly Income for the current month column.
' Locates the "HOURS USED BY MONTH" and "REVENUE COLLECTED FOR COMPLETED
' BOOKINGS" section headers in P.2 col A, then fuzzy-matches room labels.
' =============================================================================
Private Sub UpdateP2MeetingRoomData(dashWb As Workbook, bridgeWb As Workbook, monthSerial As Long)
    Dim p2Step As String
    p2Step = "get sheets"
    On Error GoTo P2RoomErrorHandler

    Dim p2Ws  As Worksheet
    Dim mrhWs As Worksheet

    On Error Resume Next
    Set p2Ws  = dashWb.Sheets("P2. Monthly Income")
    Set mrhWs = bridgeWb.Sheets("MtgRoomHours")
    On Error GoTo P2RoomErrorHandler

    If mrhWs Is Nothing Then
        Application.StatusBar = "P.2 meeting rooms: no MtgRoomHours in Data Bridge -- skipping"
        Exit Sub
    End If
    If p2Ws Is Nothing Then
        Application.StatusBar = "P.2 meeting rooms: P2. Monthly Income not found -- skipping"
        Exit Sub
    End If

    ' ── Step 1: find section header rows ────────────────────────────────────────
    ' Section labels ("HOURS USED BY MONTH", "REVENUE COLLECTED FOR COMPLETED
    ' BOOKINGS") are in col B of the Meeting Rooms section.  Month names sit in
    ' the SAME row as each label, cols C onwards.  Scan cols A+B together so the
    ' code handles either layout without breaking.
    p2Step = "find section headers"
    Dim curMonthName    As String
    On Error Resume Next
    curMonthName = Format(CDate(monthSerial), "mmmm")
    On Error GoTo P2RoomErrorHandler

    Dim hoursStartRow   As Long
    Dim hoursHdrRow2    As Long
    Dim revenueStartRow As Long
    Dim p2LastRow       As Long
    hoursStartRow   = 0
    hoursHdrRow2    = 0
    revenueStartRow = 0
    p2LastRow = p2Ws.Cells(p2Ws.Rows.Count, 2).End(xlUp).Row

    Dim rr As Long
    For rr = 1 To Application.WorksheetFunction.Min(p2LastRow, 250)
        Dim secTxt As String
        secTxt = UCase(Trim(CStr(p2Ws.Cells(rr, 1).Value))) & _
                 UCase(Trim(CStr(p2Ws.Cells(rr, 2).Value)))
        If hoursHdrRow2 = 0 And InStr(secTxt, "HOURS USED") > 0 Then
            hoursHdrRow2  = rr
            hoursStartRow = rr + 1
        End If
        If revenueStartRow = 0 And InStr(secTxt, "REVENUE COLLECTED FOR COMPLETED") > 0 Then
            revenueStartRow = rr + 1
        End If
        If hoursHdrRow2 > 0 And revenueStartRow > 0 Then Exit For
    Next rr

    If hoursStartRow = 0 And revenueStartRow = 0 Then
        Application.StatusBar = "P.2 meeting rooms: HOURS USED / REVENUE COLLECTED headers not found"
        Exit Sub
    End If

    ' ── Step 2: find the May (current month) column ───────────────────────────
    ' Read month names from the HOURS USED header row (same row as the label).
    ' Months can be plain text ("January") or date serials formatted as month
    ' names -- handle both.  Cols start at C (col 3) for the meeting rooms block.
    Dim p2Col As Long
    p2Col = 0
    Dim scanRow As Long
    scanRow = IIf(hoursHdrRow2 > 0, hoursHdrRow2, revenueStartRow - 1)
    Dim p2c As Long
    For p2c = 2 To 20
        Dim p2CellVal   As Variant
        Dim p2CellMonth As String
        p2CellVal   = p2Ws.Cells(scanRow, p2c).Value
        p2CellMonth = ""
        If IsDate(p2CellVal) Then
            On Error Resume Next
            p2CellMonth = LCase(Format(CDate(p2CellVal), "mmmm"))
            On Error GoTo P2RoomErrorHandler
        ElseIf IsNumeric(p2CellVal) Then
            If CDbl(p2CellVal) > 40000 And CDbl(p2CellVal) < 60000 Then
                On Error Resume Next
                p2CellMonth = LCase(Format(CDate(p2CellVal), "mmmm"))
                On Error GoTo P2RoomErrorHandler
            End If
        Else
            p2CellMonth = LCase(Trim(CStr(p2CellVal)))
        End If
        If p2CellMonth <> "" And p2CellMonth = LCase(curMonthName) Then
            p2Col = p2c
            Exit For
        End If
    Next p2c

    If p2Col = 0 Then
        Application.StatusBar = "P.2 meeting rooms: '" & curMonthName & _
                                 "' not found in header row " & scanRow & " -- skipping"
        Exit Sub
    End If

    ' Read MtgRoomHours: col A = P2Label, col C = Hours, col D = Revenue
    p2Step = "read MtgRoomHours"
    Dim hoursMap   As Object
    Dim revenueMap As Object
    Set hoursMap   = CreateObject("Scripting.Dictionary")
    Set revenueMap = CreateObject("Scripting.Dictionary")

    Dim mrhLastRow As Long
    mrhLastRow = mrhWs.Cells(mrhWs.Rows.Count, 1).End(xlUp).Row
    Dim mrhRow As Long
    For mrhRow = 2 To mrhLastRow  ' row 1 = header
        Dim p2Lbl As String
        p2Lbl = LCase(Trim(CStr(mrhWs.Cells(mrhRow, 1).Value)))
        If p2Lbl = "" Then GoTo NextMrhRow
        Dim mrhHrs As Double
        Dim mrhRev As Double
        mrhHrs = 0: mrhRev = 0
        On Error Resume Next
        mrhHrs = CDbl(mrhWs.Cells(mrhRow, 3).Value)
        mrhRev = CDbl(mrhWs.Cells(mrhRow, 4).Value)
        On Error GoTo P2RoomErrorHandler
        hoursMap(p2Lbl) = mrhHrs
        revenueMap(p2Lbl) = mrhRev
NextMrhRow:
    Next mrhRow

    ' Write hours to P.2 hours section
    p2Step = "write hours"
    Dim hoursWritten As Long
    hoursWritten = 0
    If hoursStartRow > 0 Then
        Dim hr As Long
        For hr = hoursStartRow To hoursStartRow + 12
            Dim hRowLbl As String
            ' Labels are in col B for the Meeting Rooms section; col A is blank
            hRowLbl = LCase(Trim(CStr(p2Ws.Cells(hr, 2).Value)))
            If hRowLbl = "" Then hRowLbl = LCase(Trim(CStr(p2Ws.Cells(hr, 1).Value)))
            If hRowLbl = "" Then GoTo NextHoursRow
            ' Exit once we hit the blank spacer rows before the next section
            If hRowLbl = "total" Or InStr(hRowLbl, "revenue") > 0 Then Exit For
            Dim hKey As Variant
            For Each hKey In hoursMap.Keys
                Dim hKeyStr As String
                hKeyStr = LCase(CStr(hKey))
                If hKeyStr = hRowLbl Or InStr(hRowLbl, hKeyStr) > 0 Or InStr(hKeyStr, hRowLbl) > 0 Then
                    p2Ws.Cells(hr, p2Col).Value = hoursMap(hKey)
                    hoursWritten = hoursWritten + 1
                    Exit For
                End If
            Next hKey
NextHoursRow:
        Next hr
    End If

    ' Write revenue to P.2 revenue section
    p2Step = "write revenue"
    Dim revWritten As Long
    revWritten = 0
    If revenueStartRow > 0 Then
        Dim rv As Long
        For rv = revenueStartRow To revenueStartRow + 12
            Dim rRowLbl As String
            ' Labels are in col B for the Meeting Rooms section; col A is blank
            rRowLbl = LCase(Trim(CStr(p2Ws.Cells(rv, 2).Value)))
            If rRowLbl = "" Then rRowLbl = LCase(Trim(CStr(p2Ws.Cells(rv, 1).Value)))
            If rRowLbl = "" Then GoTo NextRevRow
            If rRowLbl = "total" Or InStr(rRowLbl, "day pass") > 0 Then Exit For
            Dim rKey As Variant
            For Each rKey In revenueMap.Keys
                Dim rKeyStr As String
                rKeyStr = LCase(CStr(rKey))
                If rKeyStr = rRowLbl Or InStr(rRowLbl, rKeyStr) > 0 Or InStr(rKeyStr, rRowLbl) > 0 Then
                    ' Write even if $0 — zero confirms the row was matched and processed
                    p2Ws.Cells(rv, p2Col).Value = revenueMap(rKey)
                    revWritten = revWritten + 1
                    Exit For
                End If
            Next rKey
NextRevRow:
        Next rv
    End If

    Application.StatusBar = "P.2 meeting rooms: wrote " & hoursWritten & " hours, " & revWritten & " revenue cells for " & curMonthName
    Exit Sub

P2RoomErrorHandler:
    Dim p2RoomErrNum As Long
    Dim p2RoomErrDesc As String
    p2RoomErrNum  = Err.Number
    p2RoomErrDesc = Err.Description
    On Error GoTo 0
    Err.Raise p2RoomErrNum, "[UpdateP2MeetingRoomData:" & p2Step & "]", p2RoomErrDesc
End Sub

' =============================================================================
' UPDATE DAILY SPACE TAB
' Clears the Daily Space income-detail rows and pastes current-month rows from
' the Data Bridge.  Also updates the "Select Year*" filter cell to the current
' year so the tab's check-figure formulas calculate against the right period.
' =============================================================================
Private Sub UpdateDailySpaceTab(dashWb As Workbook, bridgeWb As Workbook, monthSerial As Long)
    Dim dsRawWs As Worksheet
    Dim dsTabWs As Worksheet

    On Error Resume Next
    Set dsRawWs = bridgeWb.Sheets("DailySpaceRaw")
    Set dsTabWs = dashWb.Sheets("Daily Space")
    On Error GoTo 0

    If dsRawWs Is Nothing Then
        Application.StatusBar = "Daily Space tab: no DailySpaceRaw in Data Bridge -- skipping"
        Exit Sub
    End If
    If dsTabWs Is Nothing Then
        Application.StatusBar = "Daily Space tab: 'Daily Space' tab not found in dashboard -- skipping"
        Exit Sub
    End If

    ' Update the "Select Year*" filter cell.
    ' Scan rows 1-10 for a cell whose text starts with "select year" (col A or nearby),
    ' then update the numeric year value in the same row.
    Dim currentYear As Long
    On Error Resume Next
    currentYear = Year(CDate(monthSerial))
    On Error GoTo 0
    If currentYear < 2020 Then currentYear = Year(Now)

    Dim r As Long
    For r = 1 To 10
        Dim c As Long
        For c = 1 To 6
            If LCase(Trim(CStr(dsTabWs.Cells(r, c).Value))) Like "select year*" Then
                ' Year value is likely to the right; scan this row for a 4-digit year
                Dim fc As Long
                For fc = 1 To 10
                    Dim yv As Variant
                    yv = dsTabWs.Cells(r, fc).Value
                    If IsNumeric(yv) Then
                        If CLng(yv) >= 2000 And CLng(yv) <= 2100 Then
                            dsTabWs.Cells(r, fc).Value = currentYear
                            GoTo YearUpdated
                        End If
                    End If
                Next fc
            End If
        Next c
    Next r
YearUpdated:

    ' Find the header row in the Daily Space tab (contains "Invoice")
    Dim hdrRow As Long
    hdrRow = 0
    For r = 1 To 15
        Dim rc As Long
        For rc = 1 To 12
            If InStr(LCase(CStr(dsTabWs.Cells(r, rc).Value)), "invoice") > 0 Then
                hdrRow = r
                Exit For
            End If
        Next rc
        If hdrRow > 0 Then Exit For
    Next r

    If hdrRow = 0 Then
        Application.StatusBar = "Daily Space tab: header row not found -- skipping data paste"
        Exit Sub
    End If

    ' Clear existing data rows (hdrRow+1 onward), preserving header and summary rows
    ' that may exist after the data (check-figure formulas).  Clear only the block of
    ' consecutive non-blank rows immediately below the header.
    Dim clearStart As Long
    Dim clearEnd   As Long
    clearStart = hdrRow + 1
    clearEnd   = clearStart - 1
    For r = clearStart To clearStart + 5000
        Dim firstCell As String
        firstCell = Trim(CStr(dsTabWs.Cells(r, 1).Value))
        If firstCell <> "" Then
            clearEnd = r
        ElseIf clearEnd >= clearStart Then
            Exit For  ' first blank row after data block -- stop
        End If
    Next r

    ' Save any user formula columns in the first data row before clearing.
    ' Same save/restore pattern as UpdatePriorMonthMeetingHours.
    Dim dsTotalHdrCols As Long
    dsTotalHdrCols = dsTabWs.Cells(hdrRow, dsTabWs.Columns.Count).End(xlToLeft).Column
    Dim dsFmColList()   As Long
    Dim dsFmTemplList() As String
    Dim dsFmCount       As Long
    dsFmCount = 0
    If clearEnd >= clearStart Then
        Dim dsSc As Long
        For dsSc = 1 To dsTotalHdrCols
            If dsTabWs.Cells(clearStart, dsSc).HasFormula Then
                dsFmCount = dsFmCount + 1
                ReDim Preserve dsFmColList(1 To dsFmCount)
                ReDim Preserve dsFmTemplList(1 To dsFmCount)
                dsFmColList(dsFmCount)   = dsSc
                dsFmTemplList(dsFmCount) = dsTabWs.Cells(clearStart, dsSc).Formula
            End If
        Next dsSc
        dsTabWs.Range(dsTabWs.Cells(clearStart, 1), dsTabWs.Cells(clearEnd, 20)).ClearContents
    End If

    ' Paste DailySpaceRaw rows (skip row 1 = source headers; tab already has headers)
    Dim srcLastRow As Long
    srcLastRow = dsRawWs.Cells(dsRawWs.Rows.Count, 1).End(xlUp).Row
    If srcLastRow < 2 Then
        Application.StatusBar = "Daily Space tab: no rows in DailySpaceRaw -- cleared only"
        Exit Sub
    End If

    Dim srcLastCol As Long
    srcLastCol = dsRawWs.Cells(1, dsRawWs.Columns.Count).End(xlToLeft).Column

    dsRawWs.Range(dsRawWs.Cells(2, 1), dsRawWs.Cells(srcLastRow, srcLastCol)).Copy
    dsTabWs.Cells(hdrRow + 1, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Dim dsRowsWritten As Long
    dsRowsWritten = srcLastRow - 1

    ' Restore user formula columns after paste.
    If dsFmCount > 0 And dsRowsWritten > 0 Then
        Dim dsFi As Long
        For dsFi = 1 To dsFmCount
            Dim dsFCol  As Long:   dsFCol  = dsFmColList(dsFi)
            Dim dsFTmpl As String: dsFTmpl = dsFmTemplList(dsFi)
            If dsFTmpl <> "" Then
                dsTabWs.Cells(clearStart, dsFCol).Formula = dsFTmpl
                If dsRowsWritten > 1 Then
                    Dim dsFSrc As Range
                    Dim dsFDst As Range
                    Set dsFSrc = dsTabWs.Cells(clearStart, dsFCol)
                    Set dsFDst = dsTabWs.Range(dsTabWs.Cells(clearStart, dsFCol), _
                                               dsTabWs.Cells(clearStart + dsRowsWritten - 1, dsFCol))
                    dsFSrc.AutoFill Destination:=dsFDst
                End If
            End If
        Next dsFi
    End If

    Application.StatusBar = "Daily Space tab: pasted " & dsRowsWritten & " rows, year set to " & currentYear
End Sub

' =============================================================================
' UPDATE PRIOR MONTH MEETING HOURS TAB
' Clears the "Prior Month Meeting Hours" data rows (row 5+) and pastes the
' current-month raw booking rows from the Data Bridge.  Next month, these
' become the "prior month" rows used for P.2 recurring revenue reconciliation.
' =============================================================================
Private Sub UpdatePriorMonthMeetingHours(dashWb As Workbook, bridgeWb As Workbook)
    Dim rawWs   As Worksheet
    Dim priorWs As Worksheet

    On Error Resume Next
    Set rawWs   = bridgeWb.Sheets("BookingSummaryPriorRaw")
    Set priorWs = dashWb.Sheets("Prior Month Meeting Hours")
    On Error GoTo 0

    If rawWs Is Nothing Then
        Application.StatusBar = "Prior Month Mtg Hrs: no BookingSummaryPriorRaw in Data Bridge -- skipping"
        Exit Sub
    End If
    If priorWs Is Nothing Then
        Application.StatusBar = "Prior Month Mtg Hrs: tab not found -- skipping"
        Exit Sub
    End If

    ' Find the header row (scan for "Booking ID" in col A)
    Dim hdrRow As Long
    hdrRow = 4  ' default
    Dim hh As Long
    For hh = 1 To 15
        If InStr(1, CStr(priorWs.Cells(hh, 1).Value), "Booking", vbTextCompare) > 0 Then
            hdrRow = hh: Exit For
        End If
    Next hh
    Dim dataStart As Long
    dataStart = hdrRow + 1

    ' How many raw data columns does the bridge sheet have?
    Dim srcLastCol As Long
    srcLastCol = rawWs.Cells(1, rawWs.Columns.Count).End(xlToLeft).Column

    ' How many total header columns does the dashboard tab have (includes user formula cols)?
    Dim totalHdrCols As Long
    totalHdrCols = priorWs.Cells(hdrRow, priorWs.Columns.Count).End(xlToLeft).Column

    ' Save formula templates by scanning EVERY column in the first data row.
    ' User formula columns may be INTERLEAVED within the source range (e.g. "Days or Hours"
    ' and "CALCULATED HOURS" inserted between Duration and the native Hours column), so we
    ' cannot assume they sit only after srcLastCol.  Scanning all columns handles both cases.
    Dim fmColList()   As Long
    Dim fmTemplList() As String
    Dim fmCount       As Long
    fmCount = 0
    Dim existLastRow As Long
    existLastRow = priorWs.Cells(priorWs.Rows.Count, 1).End(xlUp).Row
    If existLastRow >= dataStart Then
        Dim scanCol As Long
        For scanCol = 1 To totalHdrCols
            If priorWs.Cells(dataStart, scanCol).HasFormula Then
                fmCount = fmCount + 1
                ReDim Preserve fmColList(1 To fmCount)
                ReDim Preserve fmTemplList(1 To fmCount)
                fmColList(fmCount)   = scanCol
                fmTemplList(fmCount) = priorWs.Cells(dataStart, scanCol).Formula
            End If
        Next scanCol
    End If

    ' Delete all existing data rows (clears both value and formula columns)
    Dim clearLastRow As Long
    clearLastRow = priorWs.Cells(priorWs.Rows.Count, 1).End(xlUp).Row
    If clearLastRow >= dataStart Then
        priorWs.Rows(dataStart & ":" & clearLastRow).Delete
    End If

    ' Paste prior-month rows from the Data Bridge (row 1 = header, skip it)
    Dim srcLastRow As Long
    srcLastRow = rawWs.Cells(rawWs.Rows.Count, 1).End(xlUp).Row
    If srcLastRow < 2 Then
        Application.StatusBar = "Prior Month Mtg Hrs: no prior-month booking rows -- cleared only"
        Exit Sub
    End If

    rawWs.Range(rawWs.Cells(2, 1), rawWs.Cells(srcLastRow, srcLastCol)).Copy
    priorWs.Cells(dataStart, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Dim rowsWritten As Long
    rowsWritten = srcLastRow - 1

    ' Re-apply saved formula columns (handles any user-added formula columns
    ' that existed BEFORE the run and weren't corrupted).
    If fmCount > 0 And rowsWritten > 0 Then
        Dim fi As Long
        For fi = 1 To fmCount
            Dim fCol  As Long
            Dim fTmpl As String
            fCol  = fmColList(fi)
            fTmpl = fmTemplList(fi)
            If fTmpl <> "" Then
                priorWs.Cells(dataStart, fCol).Formula = fTmpl
                If rowsWritten > 1 Then
                    Dim fillSrc As Range
                    Dim fillDst As Range
                    Set fillSrc = priorWs.Cells(dataStart, fCol)
                    Set fillDst = priorWs.Range( _
                        priorWs.Cells(dataStart, fCol), _
                        priorWs.Cells(dataStart + rowsWritten - 1, fCol))
                    fillSrc.AutoFill Destination:=fillDst
                End If
            End If
        Next fi
    End If

    ' Always hardcode the CALCULATED HOURS formula regardless of sheet state.
    ' The formula converts Duration text ("3.00 Hours" / "1 Days") to a number:
    '   Days  → number * 8
    '   Hours → number as-is
    ' We find both columns by header name so this is resilient to column shifts.
    Dim durationCol  As Long: durationCol  = 0
    Dim calcHrsCol   As Long: calcHrsCol   = 0
    Dim ch As Long
    For ch = 1 To totalHdrCols
        Dim hv As String
        hv = LCase(Trim(CStr(priorWs.Cells(hdrRow, ch).Value)))
        If InStr(hv, "duration") > 0 And durationCol = 0 Then durationCol = ch
        If InStr(hv, "calculated") > 0 And InStr(hv, "hour") > 0 And calcHrsCol = 0 Then calcHrsCol = ch
    Next ch

    If calcHrsCol > 0 And durationCol > 0 And rowsWritten > 0 Then
        Dim durLetter As String
        durLetter = Split(priorWs.Cells(1, durationCol).Address(True, False), "$")(0)
        ' =IF(ISNUMBER(SEARCH("Day",K5)), VALUE(TRIM(LEFT(K5,FIND(" ",K5)-1)))*8,
        '                                 VALUE(TRIM(LEFT(K5,FIND(" ",K5)-1))))
        Dim calcF As String
        calcF = "=IF(ISNUMBER(SEARCH(""Day""," & durLetter & CStr(dataStart) & "))," & _
                "VALUE(TRIM(LEFT(" & durLetter & CStr(dataStart) & ",FIND("" ""," & durLetter & CStr(dataStart) & ")-1)))*8," & _
                "VALUE(TRIM(LEFT(" & durLetter & CStr(dataStart) & ",FIND("" ""," & durLetter & CStr(dataStart) & ")-1))))"
        priorWs.Cells(dataStart, calcHrsCol).Formula = calcF
        If rowsWritten > 1 Then
            Dim cfSrc As Range
            Dim cfDst As Range
            Set cfSrc = priorWs.Cells(dataStart, calcHrsCol)
            Set cfDst = priorWs.Range(priorWs.Cells(dataStart, calcHrsCol), _
                                      priorWs.Cells(dataStart + rowsWritten - 1, calcHrsCol))
            cfSrc.AutoFill Destination:=cfDst
        End If
    End If

    Application.StatusBar = "Prior Month Mtg Hrs: pasted " & rowsWritten & " rows, Calculated Hours formula applied"
End Sub

' =============================================================================
' UPDATE FLEX MEMBERS TAB
' =============================================================================
Private Sub UpdateFlexTab(dashWb As Workbook, bridgeWb As Workbook, monthSerial As Long)
    Call UpdateMembershipTab(dashWb, bridgeWb, "Flex", "FlexRaw", monthSerial)
End Sub

' =============================================================================
' UPDATE MAIL MEMBERS TAB
' =============================================================================
Private Sub UpdateMailMembersTab(dashWb As Workbook, bridgeWb As Workbook, monthSerial As Long)
    Call UpdateMembershipTab(dashWb, bridgeWb, "Mail Members", "MailRaw", monthSerial)
End Sub

' =============================================================================
' DELETE STALE ABBREVIATED TABS
' Removes old abbreviated tab names ("WA Yardi Inc Detail By Acct" and
' "WA Yardi Adj Occupancy Rep") that were created before the macro was
' corrected to use the full names matching the SUMIF formula references.
' Safe to run even if the tabs don't exist.
' =============================================================================
Private Sub DeleteStaleTabs(dashWb As Workbook)
    Dim staleNames(1) As String
    staleNames(0) = "WA Yardi Inc Detail By Acct"
    staleNames(1) = "WA Yardi Adj Occupancy Rep"

    Dim i As Long
    For i = 0 To 1
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = dashWb.Sheets(staleNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Application.StatusBar = "Deleted stale tab: " & staleNames(i)
            Set ws = Nothing
        End If
    Next i
End Sub

' =============================================================================
' UPDATE WA YARDI INC DETAIL BY ACCOUNT TAB
' Replaces data rows (row 9+) with YardiIncDetailRaw from the Data Bridge.
' Header row (row 8) is never touched.  The SUMIF formulas in the separate
' "WA Yardi Adj Occupancy Report" tab auto-recalculate after the paste, feeding
' P.3 Occupancy rows 97 (Add Coworkers) and 98 (Discounts) for the new month.
' SAFETY: Only "WA Yardi Inc Detail By Account" is modified.
'         "WA Yardi Adj Occupancy Report" is NEVER touched — its formulas are preserved.
' =============================================================================
Private Sub UpdateYardiIncDetailTab(dashWb As Workbook, bridgeWb As Workbook)
    Dim rawWs As Worksheet
    Dim tabWs As Worksheet

    On Error Resume Next
    Set rawWs = bridgeWb.Sheets("YardiIncDetailRaw")
    Set tabWs = dashWb.Sheets("WA Yardi Inc Detail By Account")
    On Error GoTo 0

    If rawWs Is Nothing Then
        Application.StatusBar = "Yardi Inc Detail: no YardiIncDetailRaw in Data Bridge -- skipping (upload income detail file)"
        Exit Sub
    End If
    If tabWs Is Nothing Then
        Application.StatusBar = "Yardi Inc Detail: 'WA Yardi Inc Detail By Account' tab not found -- skipping"
        Exit Sub
    End If

    ' Find header row (scan for "Invoice" in first 15 rows)
    Dim hdrRow As Long
    hdrRow = 8  ' default per master file
    Dim hh As Long
    For hh = 1 To 15
        If InStr(1, CStr(tabWs.Cells(hh, 1).Value), "Invoice", vbTextCompare) > 0 Then
            hdrRow = hh: Exit For
        End If
    Next hh
    Dim dataStart As Long
    dataStart = hdrRow + 1

    ' ClearContents on a fixed large range (never Rows.Delete).
    ' Rows.Delete shifts row indices and causes Excel to shrink K1's =SUM(H9:H247)
    ' reference, which would make "Income By Detail Report" (L103 in Adj Occ tab)
    ' return a partial total after re-paste.  ClearContents leaves K1 untouched.
    Dim srcLastCol As Long
    srcLastCol = rawWs.Cells(1, rawWs.Columns.Count).End(xlToLeft).Column
    tabWs.Range(tabWs.Cells(dataStart, 1), tabWs.Cells(1000, srcLastCol)).ClearContents

    ' Paste fresh rows from bridge (bridge row 1 = headers, skip it)
    Dim srcLastRow As Long
    srcLastRow = rawWs.Cells(rawWs.Rows.Count, 1).End(xlUp).Row
    If srcLastRow < 2 Then
        Application.StatusBar = "Yardi Inc Detail: no rows in YardiIncDetailRaw -- cleared only"
        Exit Sub
    End If

    rawWs.Range(rawWs.Cells(2, 1), rawWs.Cells(srcLastRow, srcLastCol)).Copy
    tabWs.Cells(dataStart, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Application.StatusBar = "Yardi Inc Detail: pasted " & (srcLastRow - 1) & " rows -- WA Yardi Adj Occupancy Report formulas will recalculate"
End Sub

' =============================================================================
' UPDATE YARDI ADJ OCCUPANCY TAB
' Refreshes space-data rows in "WA Yardi Adj Occupancy Report" from YardiAdjOccRaw
' bridge sheet.  Column L may contain user-added SUMIF formulas (e.g. for vacant
' spaces).  All formula columns are saved before the paste and restored via
' AutoFill afterward — identical pattern to UpdatePriorMonthMeetingHours.
'
' Data section end is identified by col D containing "Occupied" or "Vacant"
' (never "Grand Total" or category labels), so the summary/formula section
' starting around row 101/107 is NEVER touched regardless of row count.
'
' If no YardiAdjOccRaw sheet is present in the bridge, sub exits silently.
' =============================================================================
Private Sub UpdateYardiAdjOccTab(dashWb As Workbook, bridgeWb As Workbook)
    Dim rawWs  As Worksheet
    Dim tabWs  As Worksheet

    On Error Resume Next
    Set rawWs = bridgeWb.Sheets("YardiAdjOccRaw")
    Set tabWs = dashWb.Sheets("WA Yardi Adj Occupancy Report")
    On Error GoTo 0

    If rawWs Is Nothing Then
        Application.StatusBar = "Yardi Adj Occ: no YardiAdjOccRaw in bridge -- skipping (upload Yardi Adj Occ report)"
        Exit Sub
    End If
    If tabWs Is Nothing Then
        Application.StatusBar = "Yardi Adj Occ: 'WA Yardi Adj Occupancy Report' not found -- skipping"
        Exit Sub
    End If

    ' ── Find header row (contains "Property" and "Space" or "Status" keywords) ─
    Dim hdrRow As Long
    hdrRow = 8  ' default per reference master
    Dim hh As Long
    For hh = 1 To 15
        Dim rStr As String
        rStr = LCase(CStr(tabWs.Cells(hh, 1).Value) & " " & CStr(tabWs.Cells(hh, 2).Value) & _
                     " " & CStr(tabWs.Cells(hh, 3).Value) & " " & CStr(tabWs.Cells(hh, 4).Value))
        If (InStr(rStr, "property") > 0 Or InStr(rStr, "space") > 0) And _
           (InStr(rStr, "status") > 0 Or InStr(rStr, "type") > 0 Or InStr(rStr, "capacity") > 0) Then
            hdrRow = hh: Exit For
        End If
    Next hh
    Dim dataStart As Long
    dataStart = hdrRow + 1

    ' ── Find end of data section ──────────────────────────────────────────────
    ' Use col D "Status" column: rows that have "Occupied" or "Vacant" are space
    ' data rows.  The Grand Total / summary rows never have Status values.
    Dim dataEnd As Long
    dataEnd = 0
    Dim lastUsed As Long
    lastUsed = tabWs.Cells(tabWs.Rows.Count, 1).End(xlUp).Row
    Dim dr As Long
    For dr = dataStart To WorksheetFunction.Min(200, lastUsed)
        Dim statusVal As String
        statusVal = LCase(Trim(CStr(tabWs.Cells(dr, 4).Value)))
        If statusVal = "occupied" Or statusVal = "vacant" Then
            dataEnd = dr
        End If
    Next dr
    If dataEnd < dataStart Then
        Application.StatusBar = "Yardi Adj Occ: no Occupied/Vacant rows found in tab -- skipping"
        Exit Sub
    End If

    ' ── Bridge source dimensions ──────────────────────────────────────────────
    Dim srcLastRow As Long
    srcLastRow = rawWs.Cells(rawWs.Rows.Count, 1).End(xlUp).Row
    Dim srcLastCol As Long
    srcLastCol = rawWs.Cells(1, rawWs.Columns.Count).End(xlToLeft).Column
    ' Cap at col K (11) -- the Yardi export contains only A-K as data columns.
    ' Cols L onwards are formula-only and must never be overwritten by the paste.
    If srcLastCol > 11 Then srcLastCol = 11

    ' ── Save formula templates across ALL columns in first data row ───────────
    ' Columns like L (SUMIF for vacant spaces) may have formulas interleaved with
    ' plain-value columns.  Save every formula column so we can restore after paste.
    Dim adjFmColList()   As Long
    Dim adjFmTemplList() As String
    Dim adjFmCount       As Long
    adjFmCount = 0
    Dim scanEnd As Long
    scanEnd = WorksheetFunction.Max(srcLastCol, tabWs.Cells(hdrRow, tabWs.Columns.Count).End(xlToLeft).Column)
    Dim sc As Long
    For sc = 1 To scanEnd
        If tabWs.Cells(dataStart, sc).HasFormula Then
            adjFmCount = adjFmCount + 1
            ReDim Preserve adjFmColList(1 To adjFmCount)
            ReDim Preserve adjFmTemplList(1 To adjFmCount)
            adjFmColList(adjFmCount)   = sc
            adjFmTemplList(adjFmCount) = tabWs.Cells(dataStart, sc).Formula
        End If
    Next sc

    ' ── Clear data rows in source columns only ────────────────────────────────
    tabWs.Range(tabWs.Cells(dataStart, 1), tabWs.Cells(dataEnd, srcLastCol)).ClearContents

    ' ── Paste bridge rows (skip bridge header row 1) ──────────────────────────
    If srcLastRow < 2 Then
        Application.StatusBar = "Yardi Adj Occ: no data rows in YardiAdjOccRaw -- cleared only"
        GoTo RestoreFormulas
    End If
    Dim rowsToPaste As Long
    rowsToPaste = WorksheetFunction.Min(srcLastRow - 1, dataEnd - dataStart + 1)
    rawWs.Range(rawWs.Cells(2, 1), rawWs.Cells(rowsToPaste + 1, srcLastCol)).Copy
    tabWs.Cells(dataStart, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

RestoreFormulas:
    ' ── Restore formula columns via AutoFill ──────────────────────────────────
    Dim rowsWritten As Long
    rowsWritten = WorksheetFunction.Max(0, rowsToPaste)
    If adjFmCount > 0 And rowsWritten > 0 Then
        Dim fi As Long
        For fi = 1 To adjFmCount
            Dim fCol  As Long:   fCol  = adjFmColList(fi)
            Dim fTmpl As String: fTmpl = adjFmTemplList(fi)
            If fTmpl <> "" Then
                tabWs.Cells(dataStart, fCol).Formula = fTmpl
                If rowsWritten > 1 Then
                    Dim fSrc As Range: Set fSrc = tabWs.Cells(dataStart, fCol)
                    Dim fDst As Range
                    Set fDst = tabWs.Range(tabWs.Cells(dataStart, fCol), tabWs.Cells(dataStart + rowsWritten - 1, fCol))
                    fSrc.AutoFill Destination:=fDst
                End If
            End If
        Next fi
    End If

    ' ── Rewrite external refs in formula cells to internal ────────────────────
    ' When the .xlsm was originally saved from the .xlsx master, formula cells
    ' retained external references like:
    '   =SUMIF('[Waco Sales Dashboard MASTER.xlsx]WA Yardi Inc Detail By Account'!D:D,...)
    ' Those return $0 unless the external .xlsx is open.  Rewrite them to internal:
    '   =SUMIF('WA Yardi Inc Detail By Account'!D:D,...)
    ' We scan the full used range of the tab (not just the data section) so that
    ' the summary rows (101, 103, 104, 107-117) are also fixed on first run.
    Dim reExt As Object
    Set reExt = CreateObject("VBScript.RegExp")
    reExt.Global      = True
    reExt.IgnoreCase  = True
    ' Matches '[Any File.xlsx]SheetName'! or '[Any File.xlsx]SheetName'!
    reExt.Pattern = "'\[[^\[\]]+\.xls[xm]?\]([^']+)'!"

    Dim fullLastRow As Long
    fullLastRow = tabWs.UsedRange.Row + tabWs.UsedRange.Rows.Count - 1
    Dim fullLastCol As Long
    fullLastCol = tabWs.UsedRange.Column + tabWs.UsedRange.Columns.Count - 1

    Dim rewriteCount As Long
    rewriteCount = 0
    Dim rr As Long, cc As Long
    For rr = 1 To fullLastRow
        For cc = 1 To fullLastCol
            Dim cel As Range
            Set cel = tabWs.Cells(rr, cc)
            If cel.HasFormula Then
                Dim oldF As String
                oldF = cel.Formula
                If reExt.Test(oldF) Then
                    ' Replace '[AnyFile.xlsm]SheetName'! → 'SheetName'!
                    Dim newF As String
                    newF = reExt.Replace(oldF, "'$1'!")
                    If newF <> oldF Then
                        cel.Formula = newF
                        rewriteCount = rewriteCount + 1
                    End If
                End If
            End If
        Next cc
    Next rr

    ' ── Rewrite abbreviated internal tab name references to full names ─────────
    ' MASTER files created from older templates reference shortened tab names:
    '   'WA Yardi Inc Detail By Acct'!  →  'WA Yardi Inc Detail By Account'!
    '   'WA Yardi Adj Occupancy Rep'!   →  'WA Yardi Adj Occupancy Report'!
    ' The SUMIF formulas in the right-side summary section (cols Y+) will return
    ' zero until these are fixed. This pass runs every time so it self-heals.
    Dim abbrevCount As Long
    abbrevCount = 0
    For rr = 1 To fullLastRow
        For cc = 1 To fullLastCol
            Set cel = tabWs.Cells(rr, cc)
            If cel.HasFormula Then
                Dim abOldF As String
                Dim abNewF As String
                abOldF = cel.Formula
                abNewF = abOldF
                abNewF = Replace(abNewF, "'WA Yardi Inc Detail By Acct'!", _
                                         "'WA Yardi Inc Detail By Account'!")
                abNewF = Replace(abNewF, "'WA Yardi Adj Occupancy Rep'!", _
                                         "'WA Yardi Adj Occupancy Report'!")
                If abNewF <> abOldF Then
                    On Error Resume Next
                    cel.Formula = abNewF
                    If Err.Number = 0 Then abbrevCount = abbrevCount + 1
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
        Next cc
    Next rr

    ' ── Write correct breakdown SUMIF formulas ────────────────────────────────
    Call FixYardiBreakdownFormulas(tabWs)

    Application.StatusBar = "Yardi Adj Occ: pasted " & rowsToPaste & " rows | " & _
                            rewriteCount & " ext refs + " & abbrevCount & " tab name refs rewritten to internal"
End Sub

' =============================================================================
' FIX YARDI ADJ OCC BREAKDOWN SUMIF FORMULAS
' Scans "WA Yardi Adj Occupancy Report" for the 9 known breakdown label strings
' and writes the correct SUMIF formula in the cell immediately to the right.
' Runs every macro execution (self-healing).
'
' The labels and correct lookup columns come from the FILLED reference file:
'   "Additional Coworker Dedicated Desk"  → SUMIF on Inc Detail col D (Item)
'   All other 8 labels                    → SUMIF on Inc Detail col F (Charge Code)
'
' Col H of Inc Detail contains numeric Income values (type=n).
' No VALUE() wrapper is needed or used.
'
' These formula cells are in the AA/AB+ area of the tab, already far outside
' the data-paste range (capped at col K=11), so they can never be overwritten.
' =============================================================================
Private Sub FixYardiBreakdownFormulas(tabWs As Worksheet)
    If tabWs Is Nothing Then Exit Sub

    ' label text → True = use Inc Detail col D (Item), False = col F (Charge Code)
    Dim labels(8)   As String
    Dim useColD(8)  As Boolean
    labels(0) = "Additional Coworker Dedicated Desk" : useColD(0) = True
    labels(1) = "Discounts Dedicated Desk/DPD"       : useColD(1) = False
    labels(2) = "Additional Coworker Private Offi"   : useColD(2) = False
    labels(3) = "Discounts Private Office"            : useColD(3) = False
    labels(4) = "Flex Unlimited Income"               : useColD(4) = False
    labels(5) = "Chamber of Commerce Flex"            : useColD(5) = False
    labels(6) = "Discounts Flex Memberships"          : useColD(6) = False
    labels(7) = "Corporate Membership"                : useColD(7) = False
    labels(8) = "Summer Flex"                         : useColD(8) = False

    Dim fixedCount As Long
    fixedCount = 0

    Dim ii As Integer
    For ii = 0 To 8
        Dim found As Range
        Set found = Nothing
        On Error Resume Next
        Set found = tabWs.UsedRange.Find(What:=labels(ii), LookAt:=xlWhole, MatchCase:=False)
        On Error GoTo 0
        If Not found Is Nothing Then
            Dim lookupCol As String
            lookupCol = IIf(useColD(ii), "D", "F")
            ' Formula references the label cell so the user can rename labels freely.
            ' Criteria cell address is relative (e.g. AA8) — no fill-down needed.
            Dim criteriaAddr As String
            criteriaAddr = found.Address(False, False)   ' e.g. "AA8" or "K107"
            ' SUMPRODUCT with VALUE() coercion handles col H whether it is stored
            ' as a number or as text (Yardi exports often paste H as text strings).
            ' SUMIF ignores text values in the sum range and returns 0 in that case.
            ' Range rows 9:2000 is generous and avoids whole-column SUMPRODUCT cost.
            Dim fml As String
            fml = "=SUMPRODUCT(('WA Yardi Inc Detail By Account'!" & lookupCol & "9:" & lookupCol & "2000=" & criteriaAddr & ")" & _
                  "*IFERROR(VALUE('WA Yardi Inc Detail By Account'!H9:H2000),0))"
            Dim valCell As Range
            Set valCell = tabWs.Cells(found.Row, found.Column + 1)
            On Error Resume Next
            valCell.Formula = fml
            If Err.Number = 0 Then fixedCount = fixedCount + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next ii

    Application.StatusBar = "Yardi breakdown SUMIFs: " & fixedCount & " of 9 formulas written"
End Sub

' =============================================================================
' UPDATE MEMBERSHIP TAB (shared helper for Flex and Mail Members)
' Finds the "Select Year*" label and corrects the year, then locates the
' "Invoice" header row, clears the existing data block, and pastes the rows
' from the named Data Bridge sheet.
' tabName       — exact dashboard sheet name (e.g. "Flex", "Mail Members")
' bridgeSheetName — sheet in Data Bridge (e.g. "FlexRaw", "MailRaw")
' monthSerial   — used to derive the year for the year filter cell
' =============================================================================
Private Sub UpdateMembershipTab(dashWb As Workbook, bridgeWb As Workbook, _
                                tabName As String, bridgeSheetName As String, _
                                monthSerial As Long)
    Dim rawWs  As Worksheet
    Dim tabWs  As Worksheet

    On Error Resume Next
    Set rawWs = bridgeWb.Sheets(bridgeSheetName)
    Set tabWs = dashWb.Sheets(tabName)
    On Error GoTo 0

    If rawWs Is Nothing Then
        Application.StatusBar = tabName & " tab: no " & bridgeSheetName & " in Data Bridge -- skipping"
        Exit Sub
    End If
    If tabWs Is Nothing Then
        Application.StatusBar = tabName & " tab: sheet not found in dashboard -- skipping"
        Exit Sub
    End If

    Dim targetYear As Long
    targetYear = Year(CDate(monthSerial - 25569 + #1/1/1900#))

    ' ── Step 1: update the year filter cell ──────────────────────────────────
    ' Scan rows 1-10 for a cell containing "Select Year" label, then look
    ' across that row for the first 4-digit number to update.
    Dim yearUpdated As Boolean
    yearUpdated = False
    Dim labelRow As Long
    labelRow = 0
    Dim r As Long, c As Long
    For r = 1 To 10
        For c = 1 To 20
            If InStr(1, CStr(tabWs.Cells(r, c).Value), "Select Year", vbTextCompare) > 0 Then
                labelRow = r
                Exit For
            End If
        Next c
        If labelRow > 0 Then Exit For
    Next r

    If labelRow > 0 Then
        For c = 1 To 30
            Dim cv As Variant
            cv = tabWs.Cells(labelRow, c).Value
            If IsNumeric(cv) Then
                If CLng(cv) >= 2000 And CLng(cv) <= 2099 Then
                    tabWs.Cells(labelRow, c).Value = targetYear
                    yearUpdated = True
                    Exit For
                End If
            End If
        Next c
    End If

    ' ── Step 2: find the "Invoice" header row ────────────────────────────────
    Dim hdrRow As Long
    hdrRow = 0
    For r = 1 To 30
        For c = 1 To 20
            If InStr(1, CStr(tabWs.Cells(r, c).Value), "Invoice", vbTextCompare) > 0 Then
                hdrRow = r
                Exit For
            End If
        Next c
        If hdrRow > 0 Then Exit For
    Next r

    If hdrRow = 0 Then
        Application.StatusBar = tabName & " tab: no Invoice header row found -- skipping data paste"
        Exit Sub
    End If

    ' ── Step 3: clear existing data rows ─────────────────────────────────────
    ' IMPORTANT: Only clear source-data columns (A through srcLastCol).
    ' Do NOT delete entire rows -- user-added formula columns (J-N on Flex,
    ' J-O on Mail Members) must never be destroyed by this macro.
    Dim dataStart As Long
    dataStart = hdrRow + 1

    Dim srcLastCol As Long
    srcLastCol = rawWs.Cells(1, rawWs.Columns.Count).End(xlToLeft).Column

    Dim lastDataRow As Long
    lastDataRow = tabWs.Cells(tabWs.Rows.Count, 1).End(xlUp).Row
    If lastDataRow >= dataStart Then
        tabWs.Range(tabWs.Cells(dataStart, 1), tabWs.Cells(lastDataRow, srcLastCol)).ClearContents
    End If

    ' ── Step 4: paste rows from Data Bridge ──────────────────────────────────
    Dim srcLastRow As Long
    srcLastRow = rawWs.Cells(rawWs.Rows.Count, 1).End(xlUp).Row
    If srcLastRow < 2 Then
        Application.StatusBar = tabName & " tab: no data rows in " & bridgeSheetName & " -- cleared only"
        Exit Sub
    End If

    ' (srcLastCol declared above in Step 3)

    ' Row 1 of bridgeSheetName is the header row — paste data rows (2+) only
    rawWs.Range(rawWs.Cells(2, 1), rawWs.Cells(srcLastRow, srcLastCol)).Copy
    tabWs.Cells(dataStart, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Dim rowsWritten As Long
    rowsWritten = srcLastRow - 1
    Application.StatusBar = tabName & " tab: pasted " & rowsWritten & " rows" & _
                            IIf(yearUpdated, " (year set to " & targetYear & ")", "")
End Sub

' =============================================================================
' FILE PICKER
' =============================================================================
Private Function PickDataBridgeFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the Data Bridge file from the web tool"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\Downloads\"
        If .Show = -1 Then
            PickDataBridgeFile = .SelectedItems(1)
        Else
            PickDataBridgeFile = ""
        End If
    End With
End Function

' =============================================================================
' BACKUP WORKBOOK
' Saves a timestamped copy to a _Backups subfolder next to the master file.
' Non-fatal: if backup fails, warns but does not abort the update.
' =============================================================================
Private Sub BackupWorkbook(dashWb As Workbook)
    On Error GoTo BackupFailed

    Dim masterDir  As String
    Dim backupDir  As String
    Dim timestamp  As String
    Dim backupPath As String

    Dim slashPos As Long
    slashPos = InStrRev(dashWb.FullName, "\")
    If slashPos = 0 Then
        ' Workbook has no saved path (opened from OneDrive root or not saved) --
        ' fall back to Desktop
        masterDir = Environ("USERPROFILE") & "\Desktop\"
    Else
        masterDir = Left(dashWb.FullName, slashPos)
    End If
    backupDir  = masterDir & "_Backups\"
    timestamp  = Format(Now, "YYYY-MM-DD_HHMM")
    backupPath = backupDir & Left(dashWb.Name, InStrRev(dashWb.Name, ".") - 1) & _
                 "_" & timestamp & ".xlsm"

    ' Create _Backups folder if it does not exist
    If Dir(backupDir, vbDirectory) = "" Then
        MkDir backupDir
    End If

    dashWb.SaveCopyAs Filename:=backupPath
    Application.StatusBar = "Backup saved: " & backupPath
    Exit Sub

BackupFailed:
    MsgBox "Warning: backup could not be saved." & Chr(13) & Chr(13) & _
           "Error: " & Err.Description & Chr(13) & Chr(13) & _
           "The monthly update will continue, but you should manually save a copy of this file first.", _
           vbExclamation, "Backup Warning"
End Sub

' =============================================================================
' UPDATE OCCUPANCY DATA TAB
' Writes current-month occupancy rows to the hidden "Occupancy Data" tab so
' P.1 KPI formulas referencing 'Occupancy Data'! work without external links.
' Replaces any stale [4]Occupancy! external-link references in P.1 formulas.
' =============================================================================
Private Sub UpdateOccupancyDataTab(dashWb As Workbook, bridgeWb As Workbook)
    Dim occDataWs As Worksheet
    Dim srcWs     As Worksheet

    On Error Resume Next
    Set occDataWs = dashWb.Sheets("Occupancy Data")
    Set srcWs     = bridgeWb.Sheets("Occupancy")
    On Error GoTo 0

    ' Create the hidden tab if it does not exist yet
    If occDataWs Is Nothing Then
        Set occDataWs = dashWb.Sheets.Add(After:=dashWb.Sheets(dashWb.Sheets.Count))
        occDataWs.Name = "Occupancy Data"
        occDataWs.Visible = xlSheetVeryHidden
    End If

    ' Clear existing data (keep row 1 as header placeholder)
    If occDataWs.Cells(occDataWs.Rows.Count, 1).End(xlUp).Row > 1 Then
        occDataWs.Rows("2:" & occDataWs.Rows.Count).Delete
    End If

    If srcWs Is Nothing Then
        Application.StatusBar = "Occupancy Data tab: no Occupancy sheet in Data Bridge -- cleared only"
        GoTo FixFormulas
    End If

    ' Copy header row
    Dim srcLastCol As Long
    srcLastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column
    srcWs.Rows(1).Copy
    occDataWs.Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Copy data rows
    Dim srcLastRow As Long
    srcLastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    If srcLastRow > 1 Then
        srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(srcLastRow, srcLastCol)).Copy
        occDataWs.Cells(2, 1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If

    Application.StatusBar = "Occupancy Data tab: " & (srcLastRow - 1) & " rows written"

    ' Convert Square Feet column from string to numeric.
    ' Kube source and xlsx-js-style both write sqft as left-aligned strings
    ' (e.g. "33", "43").  SUMIF on text returns 0; explicit CDbl conversion
    ' ensures SUMIF formulas and FixP1KpiErrors aggregation both work correctly.
    Dim odLastCol As Long
    odLastCol = occDataWs.Cells(1, occDataWs.Columns.Count).End(xlToLeft).Column
    Dim odSqftCol As Long
    odSqftCol = 0
    Dim oc As Long
    For oc = 1 To odLastCol
        Dim odHdr As String
        odHdr = LCase(Trim(CStr(occDataWs.Cells(1, oc).Value)))
        If InStr(odHdr, " sq") > 0 Or odHdr = "sqft" Or odHdr = "sq ft" Or InStr(odHdr, "square") > 0 Then
            odSqftCol = oc
            Exit For
        End If
    Next oc
    If odSqftCol > 0 And srcLastRow > 1 Then
        Dim odLastRow As Long
        odLastRow = occDataWs.Cells(occDataWs.Rows.Count, 1).End(xlUp).Row
        Dim or2 As Long
        For or2 = 2 To odLastRow
            Dim odVal As Variant
            odVal = occDataWs.Cells(or2, odSqftCol).Value
            If VarType(odVal) = vbString And Len(Trim(CStr(odVal))) > 0 Then
                Dim odNum As Double
                odNum = 0
                On Error Resume Next
                odNum = CDbl(CStr(odVal))
                On Error GoTo 0
                If odNum > 0 Then occDataWs.Cells(or2, odSqftCol).Value = odNum
            End If
        Next or2
    End If

    ' NOTE: "WA Yardi Adj Occupancy Report" is NOT mirrored here.
    ' That tab contains SUMIF formulas (rows 9-99, L column) that aggregate
    ' income from "WA Yardi Inc Detail By Account".  It is refreshed by
    ' UpdateYardiAdjOccTab which pastes data rows while preserving formula columns.
    ' Clearing it here would destroy those formulas.

FixFormulas:
    Dim p1Ws As Worksheet
    On Error Resume Next
    Set p1Ws = dashWb.Sheets(P1_SHEET)
    On Error GoTo 0
    If p1Ws Is Nothing Then Exit Sub

    ' Rewrite external-file refs on every sheet so formulas resolve against internal tabs.
    ' Handles two formats:
    '   Open workbook:   [4]'Sheet Name'!  or  [4]SheetName!   → 'SheetName'!
    '   Closed workbook: 'C:\path\[file.xlsx]SheetName'!       → 'SheetName'!
    ' Internal refs ('WA Yardi Adj Occupancy Report'! etc.) are left untouched.
    Dim reOpen   As Object
    Dim reClosed As Object
    Set reOpen   = CreateObject("VBScript.RegExp")
    Set reClosed = CreateObject("VBScript.RegExp")

    reOpen.Global      = True
    reOpen.IgnoreCase  = True
    reOpen.Pattern     = "\[\d+\]'?([^'!\[\]]+)'?!"   ' captures sheet name, drops workbook#

    reClosed.Global     = True
    reClosed.IgnoreCase = True
    reClosed.Pattern    = "'[^']*\[[^\[\]]+\.xls[xm]?\]([^']*)'!"  ' captures sheet name after ]

    Dim fixedCount As Long
    fixedCount = 0

    Dim scanWs  As Worksheet
    Dim fixCell As Range
    Dim fixF    As String
    Dim fixNewF As String

    ' Only scan sheets known to carry external-file refs.
    ' Scanning all sheets is too broad and hits array formulas / protected cells
    ' that break the iterator even with On Error Resume Next.
    Dim fixSheets(4) As String
    fixSheets(0) = P1_SHEET
    fixSheets(1) = "P2. Monthly Income"
    fixSheets(2) = "P3. Occupancy"
    fixSheets(3) = P4_SHEET
    fixSheets(4) = "WA Yardi Adj Occupancy Report"

    Application.Calculation  = xlCalculationManual
    Application.EnableEvents = False

    Dim fsi As Integer
    For fsi = 0 To 4
        Set scanWs = Nothing
        On Error Resume Next
        Set scanWs = dashWb.Sheets(fixSheets(fsi))
        On Error GoTo 0
        If Not scanWs Is Nothing Then
            On Error Resume Next
            For Each fixCell In scanWs.UsedRange
                Err.Clear
                If fixCell.HasFormula Then
                    fixF = fixCell.Formula
                    If Err.Number = 0 And Len(fixF) > 0 Then
                        fixNewF = fixF
                        If reOpen.Test(fixNewF)   Then fixNewF = reOpen.Replace(fixNewF,   "'$1'!")
                        If reClosed.Test(fixNewF) Then fixNewF = reClosed.Replace(fixNewF, "'$1'!")
                        If fixNewF <> fixF Then
                            Err.Clear
                            fixCell.Formula = fixNewF
                            If Err.Number = 0 Then fixedCount = fixedCount + 1
                        End If
                    End If
                End If
            Next fixCell
            On Error GoTo 0
        End If
    Next fsi

    Application.EnableEvents = True
    Application.Calculation  = xlCalculationAutomatic

    ' Fix any P.1 KPI cells that resolved to #REF! after the external-ref rewrite
    Dim occDataWs2 As Worksheet
    On Error Resume Next
    Set occDataWs2 = dashWb.Sheets("Occupancy Data")
    On Error GoTo 0
    If Not occDataWs2 Is Nothing Then
        Call FixP1KpiErrors(p1Ws, occDataWs2, dashWb)
    End If

    If fixedCount > 0 Then
        Application.StatusBar = "External refs: rewrote " & fixedCount & " formula(s) across all sheets"
    End If
End Sub

' =============================================================================
' FIX P.1 KPI ERRORS  (replaces the old FixP1RefCells stub)
' Fixes two distinct error types in the P.1 KPI section:
'
' 1. H10/H11/H12  #DIV/0!  -- "% of Total Sq Ft" for each workspace type.
'    Root cause: F10/G10-F12/G12 (total/occupied sqft) were supplied by the
'    original Yardi external link.  After rewriting, they resolve to 0.
'    Fix: compute DD/PO sqft aggregates from Occupancy Data; write to F/G rows;
'    add IFERROR wrapper on H rows as safety net when Kube file is absent.
'
' 2. C16/D16  #REF!  -- Meeting Space Revenues (prior two months).
'    Root cause: original Yardi external link was already #REF! in the template.
'    Fix: try to pull from 2026 Sales By Month tab; fall back to 0 (manual entry).
' =============================================================================
Private Sub FixP1KpiErrors(p1Ws As Worksheet, occDataWs As Worksheet, dashWb As Workbook)

    ' ── Part A: sqft data for Current Occupancy section (rows 10-12) ─────────
    ' Detect column positions from Occupancy Data header row (row 1) so this
    ' works regardless of which column order the Kube file uses.
    Dim statusCol As Long, typeCol As Long, sqftCol As Long
    statusCol = 4: typeCol = 5: sqftCol = 7    ' sensible defaults (0-based: 3,4,6)
    Dim occLastCol As Long
    occLastCol = occDataWs.Cells(1, occDataWs.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To occLastCol
        Dim h As String
        h = LCase(Trim(CStr(occDataWs.Cells(1, c).Value)))
        If h = "status" And statusCol = 4 Then statusCol = c
        If InStr(h, "type") > 0 And typeCol = 5 Then typeCol = c
        If (InStr(h, " sq") > 0 Or h = "sqft" Or h = "sq ft" Or h = "square") And sqftCol = 7 Then sqftCol = c
    Next c

    ' Accumulate sqft by workspace type
    Dim ddTotSqFt As Double, ddOccSqFt As Double
    Dim poTotSqFt As Double, poOccSqFt As Double
    Dim allTotSqFt As Double, allOccSqFt As Double
    Dim occLastRow As Long
    occLastRow = occDataWs.Cells(occDataWs.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = 2 To occLastRow
        Dim spStat As String
        Dim spType As String
        Dim sqFt   As Double
        spStat = LCase(Trim(CStr(occDataWs.Cells(r, statusCol).Value)))
        spType = LCase(Trim(CStr(occDataWs.Cells(r, typeCol).Value)))
        sqFt = 0
        On Error Resume Next
        sqFt = CDbl(occDataWs.Cells(r, sqftCol).Value)
        On Error GoTo 0
        Dim isOcc As Boolean
        isOcc = (InStr(spStat, "occup") > 0)

        If InStr(spType, "private") > 0 Or InStr(spType, "office") > 0 Then
            poTotSqFt = poTotSqFt + sqFt
            If isOcc Then poOccSqFt = poOccSqFt + sqFt
        ElseIf InStr(spType, "dedicated") > 0 Or InStr(spType, "desk") > 0 Then
            ddTotSqFt = ddTotSqFt + sqFt
            If isOcc Then ddOccSqFt = ddOccSqFt + sqFt
        End If
        allTotSqFt = allTotSqFt + sqFt
        If isOcc Then allOccSqFt = allOccSqFt + sqFt
    Next r

    ' Find P.1 rows for Dedicated Desk, Private Office, Total Spaces
    ' by scanning col A and B for type labels (rows 5-20)
    Const SQFT_TOT_COL As Long = 6   ' col F = Total Sq Ft
    Const SQFT_OCC_COL As Long = 7   ' col G = Occupied Sq Ft
    Const SQFT_PCT_COL As Long = 8   ' col H = % Sq Ft = G/F (formula already in place)
    Dim ddRow As Long, poRow As Long, totRow As Long
    ddRow = 10: poRow = 11: totRow = 12   ' fallback defaults
    For r = 5 To 20
        Dim lbl As String
        lbl = LCase(Trim(CStr(p1Ws.Cells(r, 1).Value)))
        If lbl = "" Then lbl = LCase(Trim(CStr(p1Ws.Cells(r, 2).Value)))
        If InStr(lbl, "dedicated") > 0 Then ddRow = r
        If InStr(lbl, "private") > 0 Then poRow = r
        If InStr(lbl, "total") > 0 And (poRow <> 11 Or r > 10) Then totRow = r
    Next r

    ' Write sqft values directly to F/G (only when data is available).
    ' ClearContents first to remove any SUMIF formula that references Occupancy Data
    ' col G -- SUMIF on string-typed sqft data returns 0, and formulas must not
    ' recalculate over our written values during Application.Calculate at end.
    If ddTotSqFt > 0 Then
        p1Ws.Cells(ddRow,  SQFT_TOT_COL).ClearContents
        p1Ws.Cells(ddRow,  SQFT_OCC_COL).ClearContents
        p1Ws.Cells(ddRow,  SQFT_TOT_COL).Value = ddTotSqFt
        p1Ws.Cells(ddRow,  SQFT_OCC_COL).Value = ddOccSqFt
    End If
    If poTotSqFt > 0 Then
        p1Ws.Cells(poRow,  SQFT_TOT_COL).ClearContents
        p1Ws.Cells(poRow,  SQFT_OCC_COL).ClearContents
        p1Ws.Cells(poRow,  SQFT_TOT_COL).Value = poTotSqFt
        p1Ws.Cells(poRow,  SQFT_OCC_COL).Value = poOccSqFt
    End If
    If allTotSqFt > 0 Then
        p1Ws.Cells(totRow, SQFT_TOT_COL).ClearContents
        p1Ws.Cells(totRow, SQFT_OCC_COL).ClearContents
        p1Ws.Cells(totRow, SQFT_TOT_COL).Value = allTotSqFt
        p1Ws.Cells(totRow, SQFT_OCC_COL).Value = allOccSqFt
    End If

    ' Add IFERROR wrapper to H rows so they show 0% instead of #DIV/0!
    ' when Kube file was not included in the current upload batch.
    Dim hRow As Long
    For hRow = ddRow To totRow
        If p1Ws.Cells(hRow, SQFT_PCT_COL).HasFormula Then
            Dim fm As String
            fm = p1Ws.Cells(hRow, SQFT_PCT_COL).Formula
            If Left(UCase(fm), 9) <> "=IFERROR(" Then
                On Error Resume Next
                p1Ws.Cells(hRow, SQFT_PCT_COL).Formula = "=IFERROR(" & Mid(fm, 2) & ",0)"
                On Error GoTo 0
            End If
        End If
    Next hRow

    ' ── Part B: Meeting Space Revenues row -- clear #REF! formulas ────────────
    ' Find the row by label first, then fall back to scanning for #REF! errors
    ' in the KPI section header area (rows 13-18).
    Dim meetRow As Long
    meetRow = 0
    For r = 13 To 18
        Dim rLbl As String
        rLbl = LCase(Trim(CStr(p1Ws.Cells(r, 1).Value)))
        If rLbl = "" Then rLbl = LCase(Trim(CStr(p1Ws.Cells(r, 2).Value)))
        If InStr(rLbl, "meeting") > 0 Then meetRow = r: Exit For
    Next r
    If meetRow = 0 Then
        For r = 13 To 18
            If IsError(p1Ws.Cells(r, 3).Value) Or IsError(p1Ws.Cells(r, 4).Value) Then
                meetRow = r: Exit For
            End If
        Next r
    End If
    If meetRow = 0 Then meetRow = 16

    ' Try to pull meeting revenue values from the 2026 Sales By Month tab
    Dim sbmWs As Worksheet
    On Error Resume Next
    Set sbmWs = dashWb.Sheets(SBM_SHEET)
    On Error GoTo 0

    ' Read month header row from P.1 (row above meetRow) to know which months
    ' are in col C and D -- scan for the row that has 3+ distinct month-like values
    Dim p1HdrRow As Long
    p1HdrRow = 0
    For r = meetRow - 3 To meetRow - 1
        Dim monthCount As Long
        monthCount = 0
        For c = 2 To 10
            Dim hdrVal As Variant
            hdrVal = p1Ws.Cells(r, c).Value
            If Not IsError(hdrVal) Then
                If TypeName(hdrVal) = "Date" Then
                    monthCount = monthCount + 1
                ElseIf IsNumeric(hdrVal) Then
                    If CDbl(hdrVal) > 40000 Then monthCount = monthCount + 1
                ElseIf Trim(CStr(hdrVal)) Like "???*" Then
                    monthCount = monthCount + 1
                End If
            End If
        Next c
        If monthCount >= 2 Then p1HdrRow = r: Exit For
    Next r

    ' For each formula-error cell in meetRow cols C-D: replace with value or 0
    For c = 3 To 4
        If p1Ws.Cells(meetRow, c).HasFormula Then
            If IsError(p1Ws.Cells(meetRow, c).Value) Then
                Dim replVal As Double
                replVal = 0

                ' Try to look up value from SBM using the month header in p1HdrRow
                If Not sbmWs Is Nothing And p1HdrRow > 0 Then
                    Dim hdrCell As Variant
                    hdrCell = p1Ws.Cells(p1HdrRow, c).Value
                    Dim targetSerial As Long
                    targetSerial = CellToDateSerial(hdrCell)
                    If targetSerial = 0 And IsNumeric(hdrCell) Then targetSerial = CLng(hdrCell)
                    If targetSerial > 0 Then
                        Dim sbmHdrRow As Long
                        sbmHdrRow = FindSBMHeaderRow(sbmWs)
                        If sbmHdrRow > 0 Then
                            Dim sbmCol As Long
                            sbmCol = 0
                            Dim sc As Long
                            For sc = 2 To sbmWs.Cells(sbmHdrRow, sbmWs.Columns.Count).End(xlToLeft).Column
                                Dim sv As Variant
                                sv = sbmWs.Cells(sbmHdrRow, sc).Value
                                If IsNumeric(sv) Then
                                    If CLng(sv) = targetSerial Then sbmCol = sc: Exit For
                                End If
                            Next sc
                            If sbmCol > 0 Then
                                Dim sr As Long
                                Dim sbmLastRow As Long
                                sbmLastRow = sbmWs.Cells(sbmWs.Rows.Count, 1).End(xlUp).Row
                                For sr = sbmHdrRow + 1 To sbmLastRow
                                    Dim sLbl As String
                                    sLbl = LCase(Trim(CStr(sbmWs.Cells(sr, 1).Value)))
                                    If InStr(sLbl, "meeting") > 0 And InStr(sLbl, "space") > 0 Then
                                        Dim sVal As Variant
                                        sVal = sbmWs.Cells(sr, sbmCol).Value
                                        If IsNumeric(sVal) Then replVal = CDbl(sVal)
                                        Exit For
                                    End If
                                Next sr
                            End If
                        End If
                    End If
                End If

                p1Ws.Cells(meetRow, c).ClearContents
                If replVal <> 0 Then
                    p1Ws.Cells(meetRow, c).Value = replVal
                Else
                    p1Ws.Cells(meetRow, c).Value = 0
                    ' Note: finance team enters actual meeting revenue from Kube Booking Summary
                End If
            End If
        End If
    Next c

    Application.StatusBar = "P.1 KPI errors fixed: sqft written, Meeting Revenue row cleared"
End Sub

' =============================================================================
' FIND EXISTING SNAPSHOT ROW
' Scans the full P.1 snapshot block for a row whose month serial matches.
' Returns the row number if found, 0 if not found.
' =============================================================================
Private Function FindExistingSnapshotRow(ws As Worksheet, _
                                          monthSerial As Long, _
                                          avgRow As Long) As Long
    ' NOTE: Excel returns Date variant (not numeric Long) for date-formatted cells.
    ' Use CellToDateSerial() which handles both Date and numeric variants.
    Dim r  As Long
    Dim v  As Variant
    For r = avgRow - 1 To 1 Step -1
        v = ws.Cells(r, P1_COL_MONTH).Value
        If IsError(v) Then
            ' skip broken formula cells
        ElseIf TypeName(v) = "Date" Or IsNumeric(v) Then
            Dim s As Long
            s = CellToDateSerial(v)
            If s > 0 Then
                If s = monthSerial Then
                    FindExistingSnapshotRow = r
                    Exit Function
                End If
            ElseIf IsNumeric(v) And CDbl(v) > 0 And CDbl(v) < 40000 Then
                Exit For  ' non-date number -- hit header area
            End If
        ElseIf Trim(CStr(v)) <> "" And LCase(Trim(CStr(v))) <> "average" Then
            Exit For  ' text label -- stop (but skip the word "average" itself)
        End If
    Next r
    FindExistingSnapshotRow = 0
End Function

' =============================================================================
' UPDATE P.2 DAILY SPACE
' P.2 Monthly Income has a fixed Jan-Dec column layout (row 2 headers = month names).
' Writes the current-month and prior-month Daily Space revenue from the Data Bridge
' into the appropriate cells.  Instructions: "Update the daily space section of P.2
' for the current and prior month using the income detail by account report."
' =============================================================================
Private Sub UpdateP2DailySpace(dashWb As Workbook, _
                                monthSerial As Long, _
                                dailySpaceCurrent As Double, _
                                dailySpacePrior As Double)
    ' Skip entirely if no daily space data was parsed
    If dailySpaceCurrent = 0 And dailySpacePrior = 0 Then
        Application.StatusBar = "P.2 Daily Space: no data from Data Bridge -- skipping"
        Exit Sub
    End If

    Dim p2Ws As Worksheet
    On Error Resume Next
    Set p2Ws = dashWb.Sheets("P2. Monthly Income")
    On Error GoTo 0
    If p2Ws Is Nothing Then
        Application.StatusBar = "P.2 Daily Space: P2. Monthly Income not found -- skipping"
        Exit Sub
    End If

    ' P.2 header row = row 2.  Month names are plain English ("January", "May", etc.).
    Const HDR_ROW       As Long = 2
    Const FIRST_MON_COL As Long = 2  ' col B = January
    Const LAST_MON_COL  As Long = 13 ' col M = December

    ' Convert month serial -> full month name (e.g. 46023 -> "May")
    Dim curMonthName  As String
    Dim priorMonthName As String
    On Error Resume Next
    curMonthName   = Format(CDate(monthSerial), "mmmm")
    ' Prior month: go back one month
    Dim priorDate As Date
    priorDate = DateSerial(Year(CDate(monthSerial)), Month(CDate(monthSerial)) - 1, 1)
    priorMonthName = Format(priorDate, "mmmm")
    On Error GoTo 0

    ' Find column indices for current and prior months in the P.2 header row
    Dim curCol   As Long
    Dim priorCol As Long
    curCol = 0: priorCol = 0
    Dim c As Long
    For c = FIRST_MON_COL To LAST_MON_COL
        Dim hdr As String
        hdr = Trim(CStr(p2Ws.Cells(HDR_ROW, c).Value))
        If LCase(hdr) = LCase(curMonthName)   Then curCol   = c
        If LCase(hdr) = LCase(priorMonthName) Then priorCol = c
    Next c

    If curCol = 0 Then
        Application.StatusBar = "P.2 Daily Space: '" & curMonthName & "' column not found in P.2"
        Exit Sub
    End If

    ' Find the Daily Space row by label in col A (scan rows 2-30)
    Dim dsRow As Long
    dsRow = 0
    Dim r As Long
    For r = HDR_ROW To 30
        If InStr(LCase(Trim(CStr(p2Ws.Cells(r, 1).Value))), "daily") > 0 Then
            dsRow = r
            Exit For
        End If
    Next r
    If dsRow = 0 Then
        Application.StatusBar = "P.2 Daily Space: 'Daily Space' row not found in P.2"
        Exit Sub
    End If

    ' Write values (only overwrite if non-zero to avoid blanking manually-entered data)
    If dailySpaceCurrent <> 0 Then
        p2Ws.Cells(dsRow, curCol).Value = dailySpaceCurrent
    End If
    If dailySpacePrior <> 0 And priorCol > 0 Then
        p2Ws.Cells(dsRow, priorCol).Value = dailySpacePrior
    End If

    Application.StatusBar = "P.2 Daily Space: wrote " & curMonthName & " (" & _
                             Format(dailySpaceCurrent, "$#,##0") & ") and " & _
                             priorMonthName & " (" & Format(dailySpacePrior, "$#,##0") & ")"
End Sub

' =============================================================================
' RUN POST VALIDATION
' Checks key outputs after the update and shows a summary report.
' Non-blocking: shows warnings but does not halt or roll back.
' =============================================================================
Private Sub RunPostValidation(dashWb As Workbook, bridgeWb As Workbook, monthCode As String, monthSerial As Long)
    Dim p1Ws As Worksheet
    Dim p3Ws As Worksheet

    On Error Resume Next
    Set p1Ws = dashWb.Sheets(P1_SHEET)
    Set p3Ws = dashWb.Sheets(P3_SHEET)
    On Error GoTo 0

    Dim checks()  As String
    Dim results() As String
    Dim icons()   As String
    Dim nChecks   As Long
    nChecks = 0
    ReDim checks(20)
    ReDim results(20)
    ReDim icons(20)

    ' -- Check 1: P.1 snapshot row inserted for this month --------------------
    Dim p1Inserted As Boolean
    Dim p1DupCount As Long
    p1Inserted = False
    p1DupCount = 0
    Dim p1FoundMonths As String
    p1FoundMonths = ""
    If Not p1Ws Is Nothing Then
        Dim r As Long
        For r = 10 To 200
            Dim v As Variant
            v = p1Ws.Cells(r, P1_COL_MONTH).Value
            Dim s As Long
            s = CellToDateSerial(v)
            If s > 0 Then
                Dim mc As String
                mc = ""
                On Error Resume Next
                mc = Format(CDate(s), "mmm-yy")
                On Error GoTo 0
                If mc <> "" Then
                    If LCase(mc) = LCase(monthCode) Then
                        If p1Inserted Then
                            p1DupCount = p1DupCount + 1
                        Else
                            p1Inserted = True
                        End If
                    End If
                    If InStr(p1FoundMonths, mc) = 0 Then
                        p1FoundMonths = p1FoundMonths & mc & " "
                    End If
                End If
            End If
        Next r
    End If
    checks(nChecks) = "P.1 snapshot row"
    If p1Inserted Then
        If p1DupCount > 0 Then
            results(nChecks) = monthCode & " present (" & p1DupCount + 1 & " copies -- run CleanupP1DuplicateRows)"
            icons(nChecks)   = "!"
        Else
            results(nChecks) = "Row for " & monthCode & " present"
            icons(nChecks)   = ChrW(10003)
        End If
    Else
        results(nChecks) = "NOT FOUND -- Months seen: " & IIf(p1FoundMonths <> "", p1FoundMonths, "(none)")
        icons(nChecks)   = "!"
    End If
    nChecks = nChecks + 1

    ' -- Check 2: P.1 KPI cells have no #VALUE! errors -----------------------
    Dim p1Errors As Long
    p1Errors = 0
    If Not p1Ws Is Nothing Then
        Dim cell As Range
        For Each cell In p1Ws.UsedRange
            If IsError(cell.Value) Then p1Errors = p1Errors + 1
        Next cell
    End If
    checks(nChecks)  = "P.1 formula errors"
    results(nChecks) = IIf(p1Errors = 0, "No errors", p1Errors & " cell(s) show errors -- check external links")
    icons(nChecks)   = IIf(p1Errors = 0, ChrW(10003), "!")
    nChecks = nChecks + 1

    ' -- Check 3: P.3 column for this month exists ----------------------------
    Dim p3ColFound As Boolean
    p3ColFound = False
    If Not p3Ws Is Nothing Then
        Dim monthLabel As String
        monthLabel = ExpandMonthCode(monthCode)
        Dim c As Long
        For c = 3 To p3Ws.Cells(3, p3Ws.Columns.Count).End(xlToLeft).Column
            Dim hdrV As Variant
            hdrV = p3Ws.Cells(3, c).Value
            Dim hdrStr As String
            hdrStr = ""
            If IsError(hdrV) Then
                hdrStr = ""
            ElseIf TypeName(hdrV) = "Date" Then
                On Error Resume Next
                hdrStr = Format(hdrV, "mmm-yyyy")
                On Error GoTo 0
            ElseIf IsNumeric(hdrV) Then
                Dim hdrD As Double
                hdrD = CDbl(hdrV)
                If hdrD > 40000 Then
                    On Error Resume Next
                    hdrStr = Format(CDate(hdrD), "mmm-yyyy")
                    On Error GoTo 0
                Else
                    hdrStr = LCase(Trim(CStr(hdrV)))
                End If
            Else
                hdrStr = LCase(Trim(CStr(hdrV)))
            End If
            If LCase(hdrStr) = LCase(monthLabel) Then
                p3ColFound = True
                Exit For
            End If
        Next c
    End If
    checks(nChecks)  = "P.3 month column"
    results(nChecks) = IIf(p3ColFound, monthCode & " column present", "Column NOT added -- check Occupancy sheet in Data Bridge")
    icons(nChecks)   = IIf(p3ColFound, ChrW(10003), "!")
    nChecks = nChecks + 1

    ' -- Check 4: Occupancy Data tab populated --------------------------------
    Dim occDataRows As Long
    occDataRows = 0
    Dim occDataWs As Worksheet
    On Error Resume Next
    Set occDataWs = dashWb.Sheets("Occupancy Data")
    On Error GoTo 0
    If Not occDataWs Is Nothing Then
        occDataRows = occDataWs.Cells(occDataWs.Rows.Count, 1).End(xlUp).Row - 1
    End If
    checks(nChecks)  = "Occupancy Data tab"
    results(nChecks) = IIf(occDataRows > 0, occDataRows & " rows written", "Empty -- P.1 KPIs may show #VALUE!")
    icons(nChecks)   = IIf(occDataRows > 0, ChrW(10003), "!")
    nChecks = nChecks + 1

    ' -- Check 5: Daily Space tab has data rows --------------------------------
    Dim dsTabWs2 As Worksheet
    Dim dsDataRows As Long
    dsDataRows = 0
    On Error Resume Next
    Set dsTabWs2 = dashWb.Sheets("Daily Space")
    On Error GoTo 0
    If Not dsTabWs2 Is Nothing Then
        ' Count non-blank rows below the header (header row found by "Invoice" scan)
        Dim dsHdrRow2 As Long
        dsHdrRow2 = 0
        Dim dsr As Long
        For dsr = 1 To 15
            Dim dsc As Long
            For dsc = 1 To 12
                If InStr(LCase(CStr(dsTabWs2.Cells(dsr, dsc).Value)), "invoice") > 0 Then
                    dsHdrRow2 = dsr: Exit For
                End If
            Next dsc
            If dsHdrRow2 > 0 Then Exit For
        Next dsr
        If dsHdrRow2 > 0 Then
            For dsr = dsHdrRow2 + 1 To dsHdrRow2 + 1000
                If Trim(CStr(dsTabWs2.Cells(dsr, 1).Value)) <> "" Then
                    dsDataRows = dsDataRows + 1
                Else
                    Exit For
                End If
            Next dsr
        End If
    End If
    checks(nChecks)  = "Daily Space tab"
    If dsTabWs2 Is Nothing Then
        results(nChecks) = "Tab not found in dashboard"
        icons(nChecks)   = "!"
    ElseIf dsDataRows = 0 Then
        results(nChecks) = "No data rows (Income Detail - Daily Space not uploaded?)"
        icons(nChecks)   = "!"
    Else
        results(nChecks) = dsDataRows & " transaction rows pasted"
        icons(nChecks)   = ChrW(10003)
    End If
    nChecks = nChecks + 1

    ' -- Check 6: P.2 Daily Space revenue filled for current month -------------
    Dim p2Ws2   As Worksheet
    Dim p2DsVal As Double
    p2DsVal = 0
    On Error Resume Next
    Set p2Ws2 = dashWb.Sheets("P2. Monthly Income")
    On Error GoTo 0
    If Not p2Ws2 Is Nothing Then
        ' Find current month column in row 2, then scan for "daily space" label in col A
        Dim p2MonthCol2 As Long
        p2MonthCol2 = 0
        Dim p2ccc As Long
        Dim p2CurMon As String
        p2CurMon = ""
        On Error Resume Next
        p2CurMon = Format(CDate(monthSerial), "mmmm")
        On Error GoTo 0
        For p2ccc = 1 To p2Ws2.Cells(2, p2Ws2.Columns.Count).End(xlToLeft).Column
            If LCase(Trim(CStr(p2Ws2.Cells(2, p2ccc).Value))) = LCase(p2CurMon) Then
                p2MonthCol2 = p2ccc: Exit For
            End If
        Next p2ccc
        If p2MonthCol2 > 0 Then
            Dim p2rrr As Long
            For p2rrr = 1 To p2Ws2.Cells(p2Ws2.Rows.Count, 1).End(xlUp).Row
                Dim p2lbl2 As String
                p2lbl2 = LCase(Trim(CStr(p2Ws2.Cells(p2rrr, 1).Value)))
                If InStr(p2lbl2, "daily space") > 0 Or InStr(p2lbl2, "day pass") > 0 Then
                    On Error Resume Next
                    p2DsVal = CDbl(p2Ws2.Cells(p2rrr, p2MonthCol2).Value)
                    On Error GoTo 0
                    Exit For
                End If
            Next p2rrr
        End If
    End If
    checks(nChecks)  = "P.2 Daily Space revenue"
    If p2DsVal <> 0 Then
        results(nChecks) = Format(p2DsVal, "$#,##0.00") & " for " & monthCode
        icons(nChecks)   = ChrW(10003)
    Else
        results(nChecks) = "Zero or blank -- verify Income Detail was uploaded"
        icons(nChecks)   = "!"
    End If
    nChecks = nChecks + 1

    ' -- Check 7: Flex tab has data rows ----------------------------------------
    Dim flexTabWs    As Worksheet
    Dim flexDataRows As Long
    Dim flexHdrRow   As Long
    Dim flexLastRow  As Long
    Dim fxr As Long, fxc As Long
    flexDataRows = 0
    flexHdrRow   = 0
    On Error Resume Next
    Set flexTabWs = dashWb.Sheets("Flex")
    On Error GoTo 0
    If Not flexTabWs Is Nothing Then
        For fxr = 1 To 30
            For fxc = 1 To 20
                If InStr(1, CStr(flexTabWs.Cells(fxr, fxc).Value), "Invoice", vbTextCompare) > 0 Then
                    flexHdrRow = fxr: Exit For
                End If
            Next fxc
            If flexHdrRow > 0 Then Exit For
        Next fxr
        If flexHdrRow > 0 Then
            flexLastRow  = flexTabWs.Cells(flexTabWs.Rows.Count, 1).End(xlUp).Row
            flexDataRows = IIf(flexLastRow > flexHdrRow, flexLastRow - flexHdrRow, 0)
        End If
    End If
    checks(nChecks) = "Flex Members tab data rows"
    If flexDataRows > 0 Then
        results(nChecks) = flexDataRows & " rows written"
        icons(nChecks)   = ChrW(10003)
    ElseIf flexTabWs Is Nothing Then
        results(nChecks) = "Tab not found"
        icons(nChecks)   = "!"
    Else
        results(nChecks) = "No rows -- upload Income by Detail (Flex charge types) and regenerate"
        icons(nChecks)   = "!"
    End If
    nChecks = nChecks + 1

    ' -- Check 8: Mail Members tab has data rows --------------------------------
    Dim mailTabWs    As Worksheet
    Dim mailDataRows As Long
    Dim mailHdrRow   As Long
    Dim mailLastRow  As Long
    Dim mlr As Long, mlc As Long
    mailDataRows = 0
    mailHdrRow   = 0
    On Error Resume Next
    Set mailTabWs = dashWb.Sheets("Mail Members")
    On Error GoTo 0
    If Not mailTabWs Is Nothing Then
        For mlr = 1 To 30
            For mlc = 1 To 20
                If InStr(1, CStr(mailTabWs.Cells(mlr, mlc).Value), "Invoice", vbTextCompare) > 0 Then
                    mailHdrRow = mlr: Exit For
                End If
            Next mlc
            If mailHdrRow > 0 Then Exit For
        Next mlr
        If mailHdrRow > 0 Then
            mailLastRow  = mailTabWs.Cells(mailTabWs.Rows.Count, 1).End(xlUp).Row
            mailDataRows = IIf(mailLastRow > mailHdrRow, mailLastRow - mailHdrRow, 0)
        End If
    End If
    checks(nChecks) = "Mail Members tab data rows"
    If mailDataRows > 0 Then
        results(nChecks) = mailDataRows & " rows written"
        icons(nChecks)   = ChrW(10003)
    ElseIf mailTabWs Is Nothing Then
        results(nChecks) = "Tab not found"
        icons(nChecks)   = "!"
    Else
        results(nChecks) = "No rows -- upload Income by Detail (Mail/Mailbox charge types) and regenerate"
        icons(nChecks)   = "!"
    End If
    nChecks = nChecks + 1

    ' -- Build report string --------------------------------------------------
    Dim report As String
    report = "POST-UPDATE VALIDATION -- " & monthCode & Chr(13) & String(44, "-") & Chr(13)
    Dim j As Long
    Dim hasWarnings As Boolean
    hasWarnings = False
    For j = 0 To nChecks - 1
        report = report & Chr(13) & icons(j) & "  " & checks(j) & Chr(13) & "   " & results(j) & Chr(13)
        If icons(j) = "!" Then hasWarnings = True
    Next j

    report = report & Chr(13) & String(44, "-")
    If hasWarnings Then
        report = report & Chr(13) & Chr(13) & "One or more items need attention." & Chr(13) & _
                 "Review the flagged items above before distributing the dashboard."
    Else
        report = report & Chr(13) & Chr(13) & "All checks passed. Save and backup to the finance drive."
    End If

    Dim icon As Long
    icon = IIf(hasWarnings, vbExclamation, vbInformation)
    MsgBox report, icon, "Validation - " & monthCode
End Sub
