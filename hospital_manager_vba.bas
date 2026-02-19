Attribute VB_Name = "HospitalManager"
Option Explicit

Private Const SHEET_HOSPITAL As String = "HOSPITAL"
Private Const SHEET_LOG As String = "SERVICE_LOG"
Private Const SHEET_FORM As String = "FORM"
Private Const SHEET_ISSUE As String = "ISSUE_MASTER"

Private Const TBL_HOSPITAL As String = "tblHospital"
Private Const TBL_LOG As String = "tblLog"
Private Const TBL_ISSUE As String = "tblIssueMaster"

Public Sub InitializeHospitalWorkbook()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanFail

    EnsureMasterSheets
    BuildHospitalSheet
    BuildServiceLogSheet
    BuildIssueMasterSheet
    BuildFormSheet
    SetupFormValidations

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "병원관리대장 초기화 완료", vbInformation
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "초기화 오류: " & Err.Description, vbCritical
End Sub

Public Sub SaveServiceLog()
    On Error GoTo ErrHandler

    Dim wsForm As Worksheet, wsLog As Worksheet
    Dim tbl As ListObject, newRow As ListRow

    Set wsForm = ThisWorkbook.Worksheets(SHEET_FORM)
    Set wsLog = ThisWorkbook.Worksheets(SHEET_LOG)
    Set tbl = wsLog.ListObjects(TBL_LOG)

    If Not ValidateForm(wsForm) Then Exit Sub

    Set newRow = tbl.ListRows.Add
    With newRow.Range
        .Cells(1, GetTableColumnIndex(tbl, "로그ID")).Value = GenerateLogId()
        .Cells(1, GetTableColumnIndex(tbl, "병원ID")).Value = Trim$(wsForm.Range("B4").Value)
        .Cells(1, GetTableColumnIndex(tbl, "일자")).Value = CDate(wsForm.Range("B5").Value)
        .Cells(1, GetTableColumnIndex(tbl, "유형")).Value = Trim$(wsForm.Range("B6").Value)
        .Cells(1, GetTableColumnIndex(tbl, "증상대분류")).Value = Trim$(wsForm.Range("B7").Value)
        .Cells(1, GetTableColumnIndex(tbl, "증상소분류")).Value = Trim$(wsForm.Range("B8").Value)
        .Cells(1, GetTableColumnIndex(tbl, "원인")).Value = Trim$(wsForm.Range("B9").Value)
        .Cells(1, GetTableColumnIndex(tbl, "조치대분류")).Value = Trim$(wsForm.Range("B10").Value)
        .Cells(1, GetTableColumnIndex(tbl, "조치소분류")).Value = Trim$(wsForm.Range("B11").Value)
        .Cells(1, GetTableColumnIndex(tbl, "사용부품")).Value = Trim$(wsForm.Range("B12").Value)
        .Cells(1, GetTableColumnIndex(tbl, "다운타임(분)")).Value = NzLong(wsForm.Range("B13").Value)
        .Cells(1, GetTableColumnIndex(tbl, "담당자")).Value = Trim$(wsForm.Range("B14").Value)
        .Cells(1, GetTableColumnIndex(tbl, "다음액션")).Value = Trim$(wsForm.Range("B15").Value)
        .Cells(1, GetTableColumnIndex(tbl, "메모")).Value = Trim$(wsForm.Range("B16").Value)
    End With

    ClearForm
    MsgBox "저장 완료", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "저장 오류: " & Err.Description, vbCritical
End Sub

Public Sub UpdateHospitalIdByName()
    Dim wsForm As Worksheet
    Set wsForm = ThisWorkbook.Worksheets(SHEET_FORM)

    Dim hospitalName As String
    hospitalName = Trim$(wsForm.Range("B3").Value)

    If Len(hospitalName) = 0 Then
        wsForm.Range("B4").ClearContents
    Else
        wsForm.Range("B4").Value = GetHospitalIdByName(hospitalName)
    End If
End Sub

Public Sub RefreshIssueDropdowns()
    SetupFormValidations
    MsgBox "증상/조치 드롭다운 갱신 완료", vbInformation
End Sub

Private Function ValidateForm(ByVal wsForm As Worksheet) As Boolean
    ValidateForm = False
    If Len(Trim$(wsForm.Range("B3").Value)) = 0 Then MsgBox "병원명은 필수입니다.", vbExclamation: Exit Function
    If Len(Trim$(wsForm.Range("B4").Value)) = 0 Then MsgBox "병원ID 확인이 필요합니다.", vbExclamation: Exit Function
    If Not IsDate(wsForm.Range("B5").Value) Then MsgBox "일자는 필수입니다.", vbExclamation: Exit Function
    If Len(Trim$(wsForm.Range("B7").Value)) = 0 Or Len(Trim$(wsForm.Range("B8").Value)) = 0 Then MsgBox "증상 대/소분류는 필수입니다.", vbExclamation: Exit Function
    If Len(Trim$(wsForm.Range("B10").Value)) = 0 Or Len(Trim$(wsForm.Range("B11").Value)) = 0 Then MsgBox "조치 대/소분류는 필수입니다.", vbExclamation: Exit Function
    ValidateForm = True
End Function

Private Sub ClearForm()
    With ThisWorkbook.Worksheets(SHEET_FORM)
        .Range("B3:B16").ClearContents
        .Range("B5").Value = Date
        .Range("B5").NumberFormat = "yyyy-mm-dd"
    End With
End Sub

Private Function GenerateLogId() As String
    Dim tbl As ListObject, todayPrefix As String, maxSeq As Long
    Dim i As Long, colIdx As Long, existingId As String, seqText As String

    todayPrefix = "LOG-" & Format(Date, "yyyymmdd") & "-"
    Set tbl = ThisWorkbook.Worksheets(SHEET_LOG).ListObjects(TBL_LOG)
    maxSeq = 0

    If Not tbl.DataBodyRange Is Nothing Then
        colIdx = GetTableColumnIndex(tbl, "로그ID")
        For i = 1 To tbl.DataBodyRange.Rows.Count
            existingId = CStr(tbl.DataBodyRange.Cells(i, colIdx).Value)
            If Left$(existingId, Len(todayPrefix)) = todayPrefix Then
                seqText = Right$(existingId, 4)
                If IsNumeric(seqText) Then If CLng(seqText) > maxSeq Then maxSeq = CLng(seqText)
            End If
        Next i
    End If

    GenerateLogId = todayPrefix & Format$(maxSeq + 1, "0000")
End Function

Private Function GetHospitalIdByName(ByVal hospitalName As String) As String
    Dim tbl As ListObject, nameIdx As Long, idIdx As Long, r As Long
    Set tbl = ThisWorkbook.Worksheets(SHEET_HOSPITAL).ListObjects(TBL_HOSPITAL)
    nameIdx = GetTableColumnIndex(tbl, "병원명")
    idIdx = GetTableColumnIndex(tbl, "병원ID")
    GetHospitalIdByName = ""
    If tbl.DataBodyRange Is Nothing Then Exit Function

    For r = 1 To tbl.DataBodyRange.Rows.Count
        If StrComp(Trim$(tbl.DataBodyRange.Cells(r, nameIdx).Value), hospitalName, vbTextCompare) = 0 Then
            GetHospitalIdByName = Trim$(tbl.DataBodyRange.Cells(r, idIdx).Value)
            Exit Function
        End If
    Next r
End Function

Private Function NzLong(ByVal value As Variant) As Long
    If IsNumeric(value) Then NzLong = CLng(value) Else NzLong = 0
End Function

Private Function GetTableColumnIndex(ByVal tbl As ListObject, ByVal headerName As String) As Long
    Dim i As Long
    For i = 1 To tbl.ListColumns.Count
        If StrComp(tbl.ListColumns(i).Name, headerName, vbTextCompare) = 0 Then
            GetTableColumnIndex = i
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 513, "GetTableColumnIndex", "테이블 컬럼 미존재: " & headerName
End Function

Private Sub EnsureMasterSheets()
    EnsureSheet SHEET_HOSPITAL
    EnsureSheet SHEET_LOG
    EnsureSheet SHEET_FORM
    EnsureSheet SHEET_ISSUE
End Sub

Private Sub EnsureSheet(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
End Sub

Private Sub BuildHospitalSheet(): BuildTable ThisWorkbook.Worksheets(SHEET_HOSPITAL), TBL_HOSPITAL, Array("병원ID", "병원명", "주소", "담당자", "연락처", "부서/층", "연결장비", "설치장비모델", "설치일", "계약형태", "비고"): End Sub
Private Sub BuildServiceLogSheet(): BuildTable ThisWorkbook.Worksheets(SHEET_LOG), TBL_LOG, Array("로그ID", "병원ID", "일자", "유형", "증상대분류", "증상소분류", "원인", "조치대분류", "조치소분류", "사용부품", "다운타임(분)", "담당자", "다음액션", "메모"): End Sub
Private Sub BuildIssueMasterSheet(): BuildTable ThisWorkbook.Worksheets(SHEET_ISSUE), TBL_ISSUE, Array("구분", "대분류", "소분류", "사용여부"): End Sub

Private Sub BuildFormSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORM)
    ws.Cells.Clear

    ws.Range("A1").Value = "서비스 이력 입력 폼"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A3:A16").Value = Application.WorksheetFunction.Transpose(Array("병원명", "병원ID", "일자", "유형(방문/원격)", "증상대분류", "증상소분류", "원인", "조치대분류", "조치소분류", "사용부품", "다운타임(분)", "담당자", "다음액션", "메모"))
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 32
    ws.Range("B5").Value = Date
    ws.Range("B5").NumberFormat = "yyyy-mm-dd"
End Sub

Private Sub SetupFormValidations()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_FORM)

    BuildIssueListsOnForm

    ws.Range("B3:B12").Validation.Delete
    AddValidationList ws.Range("B3"), "=" & TBL_HOSPITAL & "[병원명]"
    AddValidationList ws.Range("B6"), "방문,원격"
    AddValidationList ws.Range("B7"), "=FORM!$H$2:$H$500"
    AddValidationList ws.Range("B8"), "=FORM!$I$2:$I$500"
    AddValidationList ws.Range("B10"), "=FORM!$J$2:$J$500"
    AddValidationList ws.Range("B11"), "=FORM!$K$2:$K$500"
End Sub

Private Sub BuildIssueListsOnForm()
    Dim wsForm As Worksheet, tbl As ListObject
    Dim mapSymMajor As Object, mapSymMinor As Object, mapActMajor As Object, mapActMinor As Object
    Dim r As Long, category As String, major As String, minor As String, useYn As String

    Set wsForm = ThisWorkbook.Worksheets(SHEET_FORM)
    Set tbl = ThisWorkbook.Worksheets(SHEET_ISSUE).ListObjects(TBL_ISSUE)

    Set mapSymMajor = CreateObject("Scripting.Dictionary")
    Set mapSymMinor = CreateObject("Scripting.Dictionary")
    Set mapActMajor = CreateObject("Scripting.Dictionary")
    Set mapActMinor = CreateObject("Scripting.Dictionary")

    If Not tbl.DataBodyRange Is Nothing Then
        For r = 1 To tbl.DataBodyRange.Rows.Count
            category = Trim$(tbl.DataBodyRange.Cells(r, GetTableColumnIndex(tbl, "구분")).Value)
            major = Trim$(tbl.DataBodyRange.Cells(r, GetTableColumnIndex(tbl, "대분류")).Value)
            minor = Trim$(tbl.DataBodyRange.Cells(r, GetTableColumnIndex(tbl, "소분류")).Value)
            useYn = UCase$(Trim$(tbl.DataBodyRange.Cells(r, GetTableColumnIndex(tbl, "사용여부")).Value))

            If useYn = "Y" Then
                If category = "증상" Then
                    If Len(major) > 0 Then mapSymMajor(major) = 1
                    If Len(minor) > 0 Then mapSymMinor(minor) = 1
                ElseIf category = "조치" Then
                    If Len(major) > 0 Then mapActMajor(major) = 1
                    If Len(minor) > 0 Then mapActMinor(minor) = 1
                End If
            End If
        Next r
    End If

    wsForm.Range("H:K").ClearContents
    wsForm.Range("H1:K1").Value = Array("증상대분류", "증상소분류", "조치대분류", "조치소분류")
    DumpDictKeys wsForm, mapSymMajor, "H2"
    DumpDictKeys wsForm, mapSymMinor, "I2"
    DumpDictKeys wsForm, mapActMajor, "J2"
    DumpDictKeys wsForm, mapActMinor, "K2"
End Sub

Private Sub DumpDictKeys(ByVal ws As Worksheet, ByVal dict As Object, ByVal startCell As String)
    Dim keys As Variant, i As Long
    If dict.Count = 0 Then Exit Sub
    keys = dict.Keys
    For i = LBound(keys) To UBound(keys)
        ws.Range(startCell).Offset(i, 0).Value = keys(i)
    Next i
End Sub

Private Sub AddValidationList(ByVal target As Range, ByVal formulaText As String)
    target.Validation.Delete
    target.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formulaText
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
End Sub

Private Sub BuildTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant)
    Dim i As Long, colCount As Long, tbl As ListObject
    colCount = UBound(headers) - LBound(headers) + 1

    ws.Cells.Clear
    For i = 0 To colCount - 1
        ws.Cells(1, i + 1).Value = headers(i)
        ws.Cells(1, i + 1).Font.Bold = True
        ws.Columns(i + 1).ColumnWidth = 14
    Next i

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    If Not tbl Is Nothing Then tbl.Delete

    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, colCount)), , xlYes)
    tbl.Name = tableName
    tbl.TableStyle = "TableStyleMedium2"

    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Rows(1).Delete
    ws.Rows(1).AutoFilter
End Sub
