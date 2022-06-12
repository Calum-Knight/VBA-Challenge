# VBA-challenge
 VBA Homework (Week 2)

Sub Stocks()

Dim First As Double
Dim Last As Double
Dim Volume As Double
Dim Diff As Double
Dim Percent As Double
Dim i As Long
Dim j As Integer
Dim ws As Integer


ws = Application.Sheets.Count

For k = 1 To ws


Worksheets(k).Select

'set up headings
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

j = 2

For i = 2 To Range("A2", Range("A2").End(xlDown)).Count + 1

'find all unique tickers and place in column i
If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
    Cells(j, 9).Value = Cells(i, 1).Value
    First = Cells(i, 3).Value
    Volume = Cells(i, 7).Value
End If

If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then              'if ticker is last of it's kind
    Last = Cells(i, 6).Value                                    'set last value
    Volume = Volume + Cells(i, 7).Value                         'update volume
    Cells(j, 12).Value = Volume                                 'print volume in cell
    Diff = Last - First                                         'calculate variance value
    Cells(j, 10) = Diff                                         'print variance value
    Percent = (Diff / First)                                    'calculate %
    Cells(j, 11).Value = Percent                                'print %
    Cells(j, 11).NumberFormat = "0.00%"                         'format as %
        If Cells(j, 11) > 0 Then
        Cells(j, 11).Interior.Color = RGB(0, 255, 0)            'format as green
        ElseIf Cells(j, 11) < 0 Then
        Cells(j, 11).Interior.Color = RGB(255, 0, 0)            'format as red
        End If
        
    j = j + 1
End If

If Cells(i, 1).Value = Cells(i - 1, 1).Value Then               'if not last ticker of its kind update volumne value
    Volume = Volume + Cells(i, 7).Value

End If

Next i

Columns("L").ColumnWidth = 20                                   'format column width

Next k

Call Summary                                                    'Call on Summary sub
    

End Sub

--------------------------------

Sub Summary()

Dim vol As Double
Dim inc As Double
Dim dec As Double
Dim inc_t As String
Dim dec_t As String
Dim vol_t As String
Dim ws As Integer


ws = Application.Sheets.Count

For k = 1 To ws


Worksheets(k).Select
'set initial variable values
inc = Cells(2, 11).Value
dec = Cells(2, 11).Value
inc_t = Cells(2, 9).Value
dec_t = Cells(2, 9).Value
vol = Cells(2, 12).Value
vol_t = Cells(2, 9).Value

    For l = 2 To Range("K2", Range("K2").End(xlDown)).Count + 1
    
    If Cells(l, 11) > inc Then          'update inc variables
        inc = Cells(l, 11).Value
        inc_t = Cells(l, 9).Value
    
    ElseIf Cells(l, 11) < dec Then      'update dec variables
        dec = Cells(l, 11).Value
        dec_t = Cells(l, 9).Value
    
    End If
    
    If Cells(l, 12).Value > vol Then    'update vol variables
        vol = Cells(l, 12).Value
        vol_t = Cells(l, 9).Value
        
        End If
    
    Next l

Cells(2, 17).Value = inc                'print inc variable
Cells(2, 17).NumberFormat = "0.00%"     'format as %
Cells(3, 17).Value = dec                'print dec variable
Cells(3, 17).NumberFormat = "0.00%"     'format as %
Cells(4, 17).Value = vol                'print vol variable
Cells(2, 16).Value = inc_t              'print ticker variable
Cells(3, 16).Value = dec_t              'print ticker variable
Cells(4, 16).Value = vol_t              'print ticker variable

Columns("Q").ColumnWidth = 20           'format column width

Next k


End Sub

