Attribute VB_Name = "DatePicker"
Option Explicit

Private Const DAY_BUTTON_PREFIX As String = "DAY_"
Private Const TOTAL_DAY_BUTTONS As Long = 42

Private Const DIALOG_NAME As String = "Date Picker"
Private Const BTN_SUPER_PREV As String = "Btn_Super_Prev"
Private Const BTN_PREV As String = "Btn_Prev"
Private Const BTN_YEAR As String = "Btn_Year"
Private Const BTN_MONTH As String = "Btn_Month"
Private Const BTN_NEXT As String = "Btn_Next"
Private Const BTN_SUPER_NEXT As String = "Btn_Super_Next"
Private Const BTN_TODAY As String = "Btn_Today"
Private Const BTN_OK As String = "Btn_OK"
Private Const BTN_CANCEL As String = "Btn_Cancel"

Private DialogSheet As DialogSheet
Private CurrentDate As Date
Private DatePickerStartYear As Long
Private DatePickerOffset As Long
Private DateMap As Object

Private Enum PickerMode
    dpNormal = 0
    dpYear = 1
    dpMonth = 2
End Enum
Private CurrentMode As PickerMode

Sub Calendar(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Set DialogSheet = Nothing
    On Error Resume Next
    Set DialogSheet = ThisWorkbook.DialogSheets(DIALOG_NAME)
    On Error GoTo ErrorHandler

    If DialogSheet Is Nothing Then
        MsgBox "找不到对话表 '" & DIALOG_NAME & "'。请确认名称是否正确。", vbExclamation, "错误"
        Exit Sub
    End If

    InitializeDatePicker
    DialogSheet.Show

    Exit Sub
ErrorHandler:
    MsgBox "无法显示日历对话框: " & Err.Description, vbExclamation, "错误"
End Sub


Private Sub InitializeDatePicker()
    Dim i As Long

    Set DateMap = CreateObject("Scripting.Dictionary")

    With DialogSheet
        On Error Resume Next
        .Buttons(BTN_SUPER_PREV).OnAction = "BtnSuperPrev"
        .Buttons(BTN_PREV).OnAction = "BtnPrev"
        .Buttons(BTN_YEAR).OnAction = "BtnYear"
        .Buttons(BTN_MONTH).OnAction = "BtnMonth"
        .Buttons(BTN_NEXT).OnAction = "BtnNext"
        .Buttons(BTN_SUPER_NEXT).OnAction = "BtnSuperNext"
        .Buttons(BTN_TODAY).OnAction = "Today"
        .Buttons(BTN_OK).OnAction = "OK"
        .Buttons(BTN_CANCEL).OnAction = "Cancel"
        For i = 1 To TOTAL_DAY_BUTTONS
            .Buttons(DAY_BUTTON_PREFIX & i).OnAction = "ButtonDate"
        Next i
        For i = 1 To 12
            .Buttons("Btn" & i).OnAction = "ButtonYearAndMonth"
        Next i
        On Error GoTo 0
    End With

    CurrentDate = Date
    DatePickerStartYear = Year(Date) - 6
    DatePickerOffset = DatePickerStartYear
    CurrentMode = dpNormal
    NormalMode
    UpdateCalendar CurrentDate
End Sub

Private Sub UpdateCalendar(Optional ByVal myDate As Date)
    Dim i As Long
    Dim MonthFirstDay As Date
    Dim MonthLastDay As Date
    Dim DatePickerFirstDay As Date

    If myDate = 0 Then myDate = Date
    MonthFirstDay = DateSerial(Year(myDate), Month(myDate), 1)
    MonthLastDay = DateSerial(Year(myDate), Month(myDate) + 1, 0)

    DatePickerFirstDay = MonthFirstDay - Weekday(MonthFirstDay, vbUseSystemDayOfWeek) + 1

    With DialogSheet
        .Buttons(BTN_YEAR).Caption = CStr(Year(myDate)) & "年"
        .Buttons(BTN_MONTH).Caption = CStr(Month(myDate)) & "月"
        For i = 1 To TOTAL_DAY_BUTTONS
            UpdateDayButton .Buttons(DAY_BUTTON_PREFIX & i), DatePickerFirstDay + i - 1, MonthFirstDay, MonthLastDay
        Next i
    End With
End Sub

Private Sub UpdateDayButton(ByRef DayButton As Object, ByVal ButtonDate As Date, ByVal MonthFirstDay As Date, ByVal MonthLastDay As Date)
    With DayButton
        .Caption = CStr(Day(ButtonDate))
        .Enabled = (ButtonDate >= MonthFirstDay And ButtonDate <= MonthLastDay)
        .Accelerator = ""
        .OnAction = "ButtonDate"
        If .Enabled And ButtonDate = Date Then .Accelerator = "T"
        If Not DateMap Is Nothing Then
            DateMap(.Name) = Format(ButtonDate, "yyyy-mm-dd")
        End If
    End With
End Sub

Sub ButtonDate()
    On Error GoTo ErrHandler
    Dim callerBtn As String
    callerBtn = Application.Caller
    
    If Not DateMap Is Nothing Then
        If DateMap.Exists(callerBtn) Then
            ActiveCell.Value = CDate(DateMap(callerBtn))
        Else
            ActiveCell.Value = DateSerial(Val(DialogSheet.Buttons(BTN_YEAR).Caption), _
                                          Val(DialogSheet.Buttons(BTN_MONTH).Caption), _
                                          Val(DialogSheet.Buttons(callerBtn).Caption))
        End If
    Else
        ActiveCell.Value = DateSerial(Val(DialogSheet.Buttons(BTN_YEAR).Caption), _
                                      Val(DialogSheet.Buttons(BTN_MONTH).Caption), _
                                      Val(DialogSheet.Buttons(Application.Caller).Caption))
    End If

    On Error Resume Next
    DialogSheet.Hide
    Set DateMap = Nothing
    Exit Sub

ErrHandler:
    MsgBox "选择日期时出错: " & Err.Description, vbExclamation, "错误"
End Sub

Private Sub ButtonYearAndMonth()
    Dim callerCaption As String
    callerCaption = DialogSheet.Buttons(Application.Caller).Caption
    Select Case CurrentMode
    Case dpYear
        DialogSheet.Buttons(BTN_YEAR).Caption = callerCaption
        UpdateCalendar DateSerial(Val(callerCaption), Val(DialogSheet.Buttons(BTN_MONTH).Caption), 1)
        NormalMode
        CurrentMode = dpNormal
        
    Case dpMonth
        DialogSheet.Buttons(BTN_MONTH).Caption = callerCaption
        UpdateCalendar DateSerial(Val(DialogSheet.Buttons(BTN_YEAR).Caption), Val(callerCaption), 1)
        NormalMode
        CurrentMode = dpNormal
    End Select
End Sub

Private Sub BtnSuperPrev()
    Select Case CurrentMode
    Case dpNormal
        ChangeMonthYear 0, -1
    Case dpYear
        DatePickerOffset = DatePickerOffset - 12
        UpdateYearButtons
    End Select
End Sub

Private Sub BtnPrev()
    Select Case CurrentMode
    Case dpNormal
        ChangeMonthYear -1, 0
    Case dpYear
        DatePickerOffset = DatePickerOffset - 3
        UpdateYearButtons
    End Select
End Sub

Private Sub BtnNext()
    Select Case CurrentMode
    Case dpNormal
        ChangeMonthYear 1, 0
    Case dpYear
        DatePickerOffset = DatePickerOffset + 3
        UpdateYearButtons
    End Select
End Sub

Private Sub BtnSuperNext()
    Select Case CurrentMode
    Case dpNormal
        ChangeMonthYear 0, 1
    Case dpYear
        DatePickerOffset = DatePickerOffset + 12
        UpdateYearButtons
    End Select
End Sub

Private Sub ChangeMonthYear(ByVal deltaMonth As Long, ByVal deltaYear As Long)
    Dim y As Long, m As Long
    With DialogSheet
        y = Val(.Buttons(BTN_YEAR).Caption)
        m = Val(.Buttons(BTN_MONTH).Caption)
    End With
    Dim newDate As Date
    newDate = DateSerial(y + deltaYear, m + deltaMonth, 1)
    UpdateCalendar newDate
End Sub

Private Sub UpdateYearButtons()
    Dim i As Long
    With DialogSheet
        For i = 1 To 12
            .Buttons("Btn" & i).Caption = CStr(DatePickerOffset + i - 1) & "年"
        Next i
    End With
End Sub

Private Sub UpdateMonthButtons()
    Dim i As Long
    With DialogSheet
        For i = 1 To 12
            .Buttons("Btn" & i).Caption = CStr(i) & "月"
        Next i
    End With
End Sub

Private Sub BtnYear()
    Select Case CurrentMode
    Case dpNormal
        CurrentMode = dpYear
        YearMode
    Case dpYear
        CurrentMode = dpNormal
        NormalMode
    Case dpMonth
        CurrentMode = dpYear
        YearMode
    End Select
End Sub

Private Sub BtnMonth()
    Select Case CurrentMode
    Case dpNormal
        CurrentMode = dpMonth
        MonthMode
    Case dpMonth
        CurrentMode = dpNormal
        NormalMode
    Case dpYear
        CurrentMode = dpMonth
        MonthMode
    End Select
End Sub

Private Sub Today()
    CurrentDate = Date
    UpdateCalendar CurrentDate
End Sub

Private Sub OK()
    On Error Resume Next
    DialogSheet.Hide
    Set DateMap = Nothing
End Sub

Private Sub Cancel()
    On Error Resume Next
    DialogSheet.Hide
    Set DateMap = Nothing
End Sub

Private Sub NormalMode()
    Dim i As Long
    With DialogSheet
        For i = 1 To 7
            .Buttons("Weekday_" & i).Enabled = False
            .Buttons("Weekday_" & i).Visible = True
        Next i
        For i = 1 To TOTAL_DAY_BUTTONS
            .Buttons(DAY_BUTTON_PREFIX & i).Visible = True
        Next i
        For i = 1 To 12
            .Buttons("Btn" & i).Visible = False
        Next i
        .Buttons(BTN_SUPER_PREV).Enabled = True
        .Buttons(BTN_PREV).Enabled = True
        .Buttons(BTN_YEAR).Visible = True
        .Buttons(BTN_MONTH).Visible = True
        .Buttons(BTN_NEXT).Visible = True
        .Buttons(BTN_SUPER_NEXT).Visible = True
        .Buttons(BTN_TODAY).Visible = True
        .Buttons(BTN_OK).Visible = True
        .Buttons(BTN_CANCEL).Visible = True
    End With
End Sub

Private Sub YearMode()
    Dim i As Long
    With DialogSheet
        For i = 1 To 7
            .Buttons("Weekday_" & i).Visible = False
        Next i
        For i = 1 To TOTAL_DAY_BUTTONS
            .Buttons(DAY_BUTTON_PREFIX & i).Visible = False
        Next i
        For i = 1 To 12
            .Buttons("Btn" & i).Visible = True
        Next i
        UpdateYearButtons
        .Buttons(BTN_SUPER_PREV).Enabled = True
        .Buttons(BTN_SUPER_NEXT).Enabled = True
        .Buttons(BTN_PREV).Enabled = True
        .Buttons(BTN_NEXT).Enabled = True
        .Buttons(BTN_TODAY).Visible = False
        .Buttons(BTN_OK).Visible = False
        .Buttons(BTN_CANCEL).Visible = False
    End With
End Sub

Private Sub MonthMode()
    Dim i As Long
    With DialogSheet
        For i = 1 To 7
            .Buttons("Weekday_" & i).Visible = False
        Next i
        For i = 1 To TOTAL_DAY_BUTTONS
            .Buttons(DAY_BUTTON_PREFIX & i).Visible = False
        Next i
        For i = 1 To 12
            .Buttons("Btn" & i).Visible = True
        Next i
        UpdateMonthButtons
        .Buttons(BTN_SUPER_PREV).Enabled = False
        .Buttons(BTN_SUPER_NEXT).Enabled = False
        .Buttons(BTN_PREV).Enabled = False
        .Buttons(BTN_NEXT).Enabled = False
        .Buttons(BTN_TODAY).Visible = False
        .Buttons(BTN_OK).Visible = False
        .Buttons(BTN_CANCEL).Visible = False
    End With
End Sub

