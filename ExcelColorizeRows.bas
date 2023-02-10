Option Explicit

Dim UseColor As Boolean, UseBorder As Boolean, BreakRows As Boolean, AddHeaderCols As Boolean
Dim Cols(), LCols As Integer, UCols As Integer, HeaderRowsCount As Integer
Dim fixed As Integer, random As Integer
Dim InANewRow As Boolean, Delimiter As String, ChangeStyle As Boolean, InCol As Integer

Dim A As Range
Dim i As Integer, j As Integer, k As Integer, isChanged As Boolean, headerData As String
Dim curColor As Long

Sub Colorize()
    On Error GoTo SomeError
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual


    '#############################################
    '############### Configuration ###############
    
    Cols = [{1, 2}]
    HeaderRowsCount = 1      ' Default: 1
    
    UseColor = True          ' Default: True
        fixed = 150          ' Default: 150
        random = 105         ' Default: 105
        'Total must be less than 256
    
    UseBorder = True         ' Default: True
    BreakRows = False        ' Default: False
    
    AddHeaderCols = False    ' Default: False
        InANewRow = True     ' Default: True
        Delimiter = " - "    ' Default: " - "
        ChangeStyle = True   ' Default: True
        InCol = -1           ' Default: -1 (auto)
    
    '#############################################
    '#############################################
    
    
    ' Reason of using random colors: To keep it useful even if the data filtered or sorted differently

    ' It's harder to make the code work for visible cells only, because unlike Cells, ...
    ' ... rows are not callable for multiple `Area`s at once

    Randomize Timer

    LCols = LBound(Cols)
    UCols = UBound(Cols)


    Set A = Me.UsedRange
    ' Also `Me` alone is used in code

    If AddHeaderCols Then
        If MsgBox("Setting `AddHeaderCols = True` adds data to the worksheet and we recommend you to do a backup before. Proceed?", vbYesNo) = vbNo Then
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Exit Sub
        End If
        
        If InCol = -1 Then
            If InANewRow Then
                InCol = 1
            Else
                InCol = A.Columns.Count + 1
            End If
        End If
    End If

    If BreakRows Then
        Me.ResetAllPageBreaks
    End If

    If UseColor Then
        curColor = RGB(fixed + (Rnd() * random), fixed + (Rnd() * random), fixed + (Rnd() * random))
    End If
    
    doAddHeaderRow HeaderRowsCount
    
    i = HeaderRowsCount + 1
    Do Until i > A.Rows.Count
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
        
        ' Second condition is used especially at the last row
        If isChanged And WorksheetFunction.CountA(A.Rows(i + 1)) > 0 Then
            If UseColor Then
                curColor = RGB(fixed + (Rnd() * random), fixed + (Rnd() * random), fixed + (Rnd() * random))
            End If
            If UseBorder Then
                A.Rows(i).Borders(xlEdgeBottom).Weight = 4
            End If
            If BreakRows Then
                Me.HPageBreaks.Add Before:=A.Rows(i + 1)
                A.Rows(i + 1).PageBreak = xlPageBreakManual
            End If
            doAddHeaderRow i
        
        Else
            If UseBorder Then
                A.Rows(i).Borders(xlEdgeBottom).Weight = 2
            End If
        End If
        i = i + 1
    Loop
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
SomeError:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error!", vbCritical

End Sub

Function doAddHeaderRow(ByRef roww As Integer)
    If AddHeaderCols Then
        k = roww + 1
        
        headerData = ""
        For j = LCols To UCols - 1
            headerData = headerData & A.Rows(k).Cells(1, Cols(j)) & Delimiter
        Next
        headerData = headerData & A.Rows(k).Cells(1, Cols(UCols))
        
        If InANewRow Then
            A.Rows(k).Insert
            If UseColor Then
                A.Rows(k).Interior.Color = curColor
            End If
            If UseBorder Then
                A.Rows(k).Borders(xlEdgeBottom).Weight = 2
            End If
            roww = k
        End If
    
        A.Rows(k).Cells(1, InCol) = headerData
        If ChangeStyle Then
            A.Rows(k).Cells(1, InCol).Font.Bold = True
            A.Rows(k).Cells(1, InCol).Font.Size = A.Rows(k).Cells(1, InCol).Font.Size * 1.2
        End If
    End If
End Function
