Sub easy()
    Dim ws As Worksheet
    Dim i As Long
    Dim n As Long
    Dim name As String
    Dim total As Double
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        i = 2
        n = 2
        While ws.Cells(i, 1).Value <> ""
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                total = total + ws.Cells(i, 7).Value
                ws.Cells(n, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(n, 12).Value = total
                total = 0
                n = n + 1
            Else
                total = total + ws.Cells(i, 7).Value
            End If
            i = i + 1
        Wend
    Next
    Application.ScreenUpdating = True
End Sub

Sub moderate()
    Dim ws As Worksheet
    Dim i As Long
    Dim n As Long
    Dim year_start As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        i = 3
        n = 2
        year_start = ws.Cells(2, 3).Value
        While ws.Cells(i, 1).Value <> ""
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                yearly_change = ws.Cells(i, 6).Value - year_start
                    If year_start <> 0 Then
                        percent_change = yearly_change / year_start
                        ws.Cells(n, 11).Value = percent_change
                    Else
                        ws.Cells(n, 11).Value = "N/A"
                    End If
                year_start = ws.Cells(i + 1, 3).Value
                ws.Cells(n, 10).Value = yearly_change
                n = n + 1
            End If
            i = i + 1
        Wend
        n = 2
        While ws.Cells(n, 10).Value <> ""
            If ws.Cells(n, 10).Value >= 0 Then
                ws.Cells(n, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(n, 10).Interior.ColorIndex = 3
            End If
            n = n + 1
        Wend
    Next
    Application.ScreenUpdating = True
End Sub

Sub hard()
    Dim ws As Worksheet
    Dim i As Long
    Dim n As Long
    Dim total As Double
    Dim total_pos As Double
    Dim max As Double
    Dim max_pos As Double
    Dim min As Double
    Dim min_pos As Double
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        i = 3
        total = ws.Cells(2, 12).Value
        While ws.Cells(i, 12).Value <> ""
            If ws.Cells(i + 1, 12).Value > total Then
                total = ws.Cells(i + 1, 12).Value
                ws.Cells(4, 17).Value = total
                total_pos = i + 1
                ws.Cells(4, 16).Value = ws.Cells(total_pos, 9).Value
            End If
            i = i + 1
        Wend
        
        i = 3
        max = ws.Cells(2, 11).Value
        While ws.Cells(i, 11).Value <> ""
            If ws.Cells(i + 1, 11).Value > max Then
                max = ws.Cells(i + 1, 11).Value
                ws.Cells(2, 17).Value = max
                max_pos = i + 1
                ws.Cells(2, 16).Value = ws.Cells(max_pos, 9).Value
            End If
            i = i + 1
        Wend
        
        i = 3
        min = ws.Cells(2, 11).Value
        While ws.Cells(i, 11).Value <> ""
            If ws.Cells(i + 1, 11).Value < min Then
                min = ws.Cells(i + 1, 11).Value
                ws.Cells(3, 17).Value = min
                min_pos = i + 1
                ws.Cells(3, 16).Value = ws.Cells(min_pos, 9).Value
            End If
            i = i + 1
        Wend
    Next
    Application.ScreenUpdating = True
End Sub

