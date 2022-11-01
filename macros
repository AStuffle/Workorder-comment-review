Sub A_HIGHLIGHT_ALL()
    Application.ScreenUpdating = False
    
    Range("Report_PQ[FailureRemark_LongDescription]").Select
    Call KeyList
    Call ExcList
    
    ActiveSheet.ListObjects("Report_PQ").Range.AutoFilter Field:=19
    Range("Report_PQ[[#Headers],[FailureRemark_LongDescription]]").Select
    ActiveSheet.ListObjects("Report_PQ").Range.AutoFilter Field:=19, Criteria1 _
        :=RGB(255, 0, 0), Operator:=xlFilterFontColor
        
    Application.ScreenUpdating = True
    Range("S1").Select
End Sub

Sub A_HIGHLIGHT_ONE()
    Application.ScreenUpdating = False
    
    Range("Report_PQ[FailureRemark_LongDescription]").Select
    Call KeyWord
    
    ActiveSheet.ListObjects("Report_PQ").Range.AutoFilter Field:=19
    Range("Report_PQ[[#Headers],[FailureRemark_LongDescription]]").Select
    ActiveSheet.ListObjects("Report_PQ").Range.AutoFilter Field:=19, Criteria1 _
        :=RGB(255, 0, 0), Operator:=xlFilterFontColor
        
    Application.ScreenUpdating = True
    Range("S1").Select
End Sub

Sub Clear()
    Application.ScreenUpdating = False
    
    Range("Report_PQ[FailureRemark_LongDescription]").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.Font.Underline = xlUnderlineStyleNone
    Selection.Font.Bold = False
    ActiveSheet.ListObjects("Report_PQ").Range.AutoFilter Field:=19
    
    Application.ScreenUpdating = True
    Range("S1").Select
End Sub

Sub Update()
    Application.ScreenUpdating = False
    
    ActiveWorkbook.Connections("Query - Fleet_WO_Comments").Refresh
    ActiveWorkbook.Connections("Query - Report_PQ").Refresh
    Sheets("Report_PQ").Select
    
    Call Clear
    Application.ScreenUpdating = True
    Range("S1").Select
End Sub

Sub KeyList()
    Application.ScreenUpdating = False
    
    Dim Rng As Range
    Dim Itm As Range
    Dim cFnd As String
    Dim xTmp As String
    Dim x As Long
    Dim m As Long
    Dim y As Long
    
For Each Itm In Range("KeyList[Included]")
    cFnd = Trim(LCase(Itm.Value))
    
        y = Len(cFnd)
        For Each Rng In Selection
            With Rng
                m = UBound(Split(Rng.Value, cFnd))
                If m > 0 Then
                  xTmp = ""
                    For x = 0 To m - 1
                        xTmp = xTmp & Split(Rng.Value, cFnd)(x)
                        .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.ColorIndex = 3
                        .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.Bold = True
                        .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.Underline = True
                        xTmp = xTmp & cFnd
                    Next
                End If
            End With
        Next Rng
    
Next Itm

    Application.ScreenUpdating = True
End Sub

Sub ExcList()
    Application.ScreenUpdating = False
    
    Dim Rng As Range
    Dim Itm As Range
    Dim cFnd As String
    Dim xTmp As String
    Dim x As Long
    Dim m As Long
    Dim y As Long
    
For Each Itm In Range("ExcList[Excluded]")
    cFnd = Trim(LCase(Itm.Value))
    
        y = Len(cFnd)
        For Each Rng In Selection
            With Rng
                m = UBound(Split(Rng.Value, cFnd))
                If m > 0 Then
                  xTmp = ""
                    For x = 0 To m - 1
                        xTmp = xTmp & Split(Rng.Value, cFnd)(x)
                        .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.ColorIndex = 0
                        .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.Bold = False
                        .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.Underline = False
                        xTmp = xTmp & cFnd
                    Next
                End If
            End With
        Next Rng
    
Next Itm

    Application.ScreenUpdating = True
End Sub

Sub KeyWord()
    Application.ScreenUpdating = False
    
    Dim Rng As Range
    Dim cFnd As String
    Dim xTmp As String
    Dim x As Long
    Dim m As Long
    Dim y As Long
    
    cFnd = Trim(LCase(InputBox("Enter the text string to highlight, use lower case only.")))
    y = Len(cFnd)
    For Each Rng In Selection
        With Rng
            m = UBound(Split(Rng.Value, cFnd))
            If m > 0 Then
              xTmp = ""
                For x = 0 To m - 1
                    xTmp = xTmp & Split(Rng.Value, cFnd)(x)
                    .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.ColorIndex = 3
                    .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.Bold = True
                    .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.Underline = True
                    xTmp = xTmp & cFnd
                Next
            End If
        End With
    Next Rng
    
    Application.ScreenUpdating = True
End Sub
