Option Explicit

Sub Colorize()
    On Error GoTo SomeError
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim UseColor As Boolean, UseBorder As Boolean
    Dim Cols(), LCols As Integer, UCols As Integer, HeaderRowsCount As Integer
    Dim fixed As Integer
    Dim random As Integer


    '############
    '############ Configuration
    HeaderRowsCount = 1
    Cols = [{1, 2}]

    UseColor = True
        fixed = 150
        random = 105
        
    UseBorder = True
    '############
    '############

  
    ' Reason of using random colors: To keep it useful even if data filtered or sorted differently
    ' "‫"دلیل استفاده از رنگ‌های تصادفی، حفظ کارایی علیرغم فیلتر کردن یا تغییر Sorting‌ه

    ' "‫"حالتی که فقط طبق سلول‌های دیده شده رفتار کنه اجراش سخت‌تره، ...
    ' "‫"...چون Rows برعکس Cells، برای Range‌های دارای چند Area بصورت یکباره قابل فراخوانی نیست

    Randomize Timer

    LCols = LBound(Cols)
    UCols = UBound(Cols)

    Dim A As Range
    Set A = ThisWorkbook.ActiveSheet.UsedRange

    Dim i As Integer, j As Integer, isChanged As Boolean
    Dim curColor As Long

    If UseColor Then
        curColor = RGB(fixed + (Rnd() * random), fixed + (Rnd() * random), fixed + (Rnd() * random))
    End If

    For i = HeaderRowsCount + 1 To A.Rows.Count
        If UseColor Then
            A.Rows(i).Interior.Color = curColor
        End If
        
        isChanged = False
        
        For j = LCols To UCols
            If A.Rows(i).Cells(1, Cols(j)) <> A.Rows(i + 1).Cells(1, Cols(j)) Then
                isChanged = True
                Exit For
            End If
        Next
        
        If isChanged Then
            If UseColor Then
                curColor = RGB(fixed + (Rnd() * random), fixed + (Rnd() * random), fixed + (Rnd() * random))
            End If
            If UseBorder Then
                a.Rows(i).Borders(xlEdgeBottom).Weight = 4
            End If
        Else
            If UseBorder Then
                a.Rows(i).Borders(xlEdgeBottom).Weight = 2
            End If
        End If
    Next

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
SomeError:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error!", vbCritical

End Sub

